<#
    .SYNOPSIS 
     Manages contacts in two Exchange organizations based on mail-enabled users in the other organization.
	.DESCRIPTION
	 This script takes users from one Exchange Org and creates contacts of those users in another
	 Exchange Org and vice versa. The Exchane Org can be On Premise or Office 365.
     It will specifically perform the following:
	  - Contacts are created for new users.
	  - Contacts are deleted if the source user no longer meets the filter requirements.
	  - Contacts are updated with changed information.
	.EXAMPLE
	 .\GALSync.ps1
	  Runs the script in Read-Only mode, the default setting. No users will be modified.
    .EXAMPLE
	 .\GALSync.ps1 -ReadOnly:$false
	 Runs the script in WRITE mode, causing changes to be committed.
	.EXAMPLE
	 .\GALSync.ps1 -LogFileDir "C:\Logs"
	  Runs the script in Read-Only mode, specifying where the log file will be created. No users will be modified.
	.NOTES
	 - Requires .Net 4.5 and PowerShell 4.0
	 - A user account is needed in each tenant with Global Administrator permissions or on in each Exchange
        organization with Recipient Administrator permissions.
     - The passwords for the user account should be encrypted using an AES key generated using the following 
        command:
        $KeyFile = ".\PathtoKey\AES.key"
        $Key = New-Object Byte[] 32   # You can use 16, 24, or 32 for AES
        [Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($Key)
        $Key | out-file $KeyFile
	 - The passwords for these user accounts must be stored in secure files using the command:
        $PasswordFile = "\\Machine1\SharedPath\Password.txt"
        $KeyFile = ".\PathtoKey\AES.key"
        $Key = Get-Content $KeyFile
		Read-Host -AsSecureString | ConvertFrom-SecureString -key $Key | Out-File $PasswordFile
     - Uses customAttribute5 to track contacts that were created by GalSync, populates with source samAccountName
	 - Created by Andy Meyers, Anexinet
        Updated by Ned Bellavance, Anexinet
		Adapted from original written by Carol Wapshere
	 - Updated on 10/26/2015
	 - Version 1.1	  
#>

[CmdletBinding()]
Param(
 [Parameter(Mandatory=$False,Position=1)]
   [Switch]$ReadOnly = $true,
 [Parameter(Mandatory=$False,Position=2)]
   [String]$LogFileDir = "C:\Scripts\logs"
)

Import-Module ActiveDirectory

### --- GLOBAL DEFINITIONS ---

#Set to location of encryption key
$Key = Get-Content "C:\Scripts\GalSync.key"

#User accounts in UPN format
$FirstUser = ""
$SecondUser = ""

#Location of password files
$FirstPWFile = "C:\Scripts\firstpwfile.txt"
$SecondPWFile = "C:\Scripts\secondpwfile.txt"

#Name of an Exchange CAS server, leave empty for Exchange Online
$FirstExchangeServer = ""
$SecondExchangeServer = ""

#Set to true if using Exchange Online
$FirstExchangeOnline = $true
$SecondExchangeOnline = $false

#Location of source accounts to create contacts from, leave blank if using Exchange Online
$FirstSourceOU = ""
$SecondSourceOU = ""

#Location of target contacts, leave blank if using Exchange Online
$FirstTargetOU = ""
$SecondTargetOU = ""

## The following list of attributes will be copied from User to Contact, verify that attributes are availble from the Set-Contact command
$arrAttribs = 'DisplayName','Company','FirstName','MobilePhone','PostalCode','LastName','StateOrProvince','StreetAddress','Phone','Title','CountryOrRegion','City','Fax','Office','Notes'

## The following filter is used by Get-Recipient to decide which users will have contacts.
$strSelectUsers = 'HiddenFromAddressListsEnabled -eq $false -and -not DisplayName -eq "Administrator" -and CustomAttribute5 -eq "GalSync"'


### --- FUNCTION TO WRITE OUT THE DATA ABOUT AN OBJECT BEFORE IT IS DELETED ---

Function WriteObjectInfo{
    [cmdletbinding()]

    param(
        $sourceObj,
        $outFile
    )
    Add-Content -Path $outFile -Value "Name:$($sourceObj.Name)"
    Add-Content -Path $outFile -Value "Alias:$($sourceObj.Alias)"
    Add-Content -Path $outFile -Value "DisplayName:$($sourceObj.DisplayName)"
    Add-Content -Path $outFile -Value "DistinguishedName:$($sourceObj.DistinguishedName)"
    Add-Content -Path $outFile -Value "Identity:$($sourceObj.Identity)"
    Add-Content -Path $outFile -Value "LegacyExchangeDN:$($sourceObj.LegacyExchangeDN)"
    if($sourceObj.RecipientType -eq "UserMailbox"){
      Add-Content -Path $outFile -Value "samAccountName:$($sourceObj.samAccountName)"  
    }
    Add-Content -Path $outFile -Value "CustomAttribute5:$($sourceObj.customAttribute5)"
    Add-Content -Path $outFile -Value "PrimarySmtpAddress:$($sourceObj.PrimarySmtpAddress)"
    Add-Content -Path $outFile -Value "ExternalEmailAddress:$($sourceObj.ExternalEmailAddress)"
    foreach($address in $sourceObj.EmailAddresses){
        Add-Content -Path $outFile -Value "EmailAddress:$address"
    }
    Add-Content -Path $outFile -Value "`n"
}


### --- FUNCTION TO ADD, DELETE AND MODIFY CONTACTS IN TARGET DOMAIN BASED ON SOURCE USERS ---

Function SyncContacts
{
  [CmdletBinding()]
  PARAM($sourceUser,
   $sourcePWFile, 
   $targetUser, 
   $targetPWFile,
   $sourceExchangeOnline,
   $targetExchangeOnline,
   $sourceExchangeServer,
   $targetExchangeServer,
   $sourceOU,
   $targetOU,
   $sourceMailDomains
   )
  END
    {
		$colUsers = @()
		$colContacts = @()
		$colAddContact = @()
		$colDelContact = @()
		$colUpdContact = @()

		$arrUserMail = @()
		$arrContactMail = @()
        $arrUserName = @()

		Try
		{
			# Setup the remote PowerShell session to the source Exchange or tenant
			$sourceCredential =  New-Object -Typename System.Management.Automation.PSCredential -Argumentlist $sourceUser, (Get-Content $sourcePWFile | ConvertTo-SecureString -Key $Key)
            if($sourceExchangeOnline){
                Write-Verbose "Connecting to Exchange online for source tenant"
    			$sourceSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $sourceCredential -Authentication Basic -AllowRedirection -ErrorAction Stop
            }
            else{
                Write-Verbose "Connecting to Exchange on-premise using source: $sourceExchangeServer"
                $sourceSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$sourceExchangeServer/powershell" -Credential $sourceCredential -ErrorAction Stop
            }
		}
		Catch
		{
			Out-File -InputObject "$(Get-Date -Format MM.dd.yyyy-HH:mm:ss);$($WriteMode);ERROR;Source;$($sourceUser.Split('@')[1]);$($sourceUser);Remoting;;;Error creating source session for $($sourceUser.Split('@')[1]): $($Error[0])" -FilePath $LogFilePath -Append
			Exit
		}
		
		Try
		{
			# Setup the remote PowerShell session to the target Exchange or tenant
			$targetCredential =  New-Object -Typename System.Management.Automation.PSCredential -Argumentlist $targetUser,(Get-Content $targetPWFile | ConvertTo-SecureString -Key $Key)
            if($targetExchangeOnline){
                Write-Verbose "Connecting to Exchange online for target tenant"
    			$targetSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $targetCredential -Authentication Basic -AllowRedirection -ErrorAction Stop
            }
            else{
                Write-Verbose "Connecting to Exchange on-premise using target: $targetExchangeServer"
                $targetSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$targetExchangeServer/powershell" -Credential $targetCredential -ErrorAction Stop
            }
		}
		Catch
		{
			Out-File -InputObject "$(Get-Date -Format MM.dd.yyyy-HH:mm:ss);$($WriteMode);ERROR;Target;$($targetUser.Split('@')[1]),$($targetUser);Remoting;;;Error creating target session for $($targetUser.Split('@')[1]): $($Error[0])" -FilePath $LogFilePath -Append
			Exit
		}
		
		### ENUMERATE USERS

		# Get all users with mailboxes or mail enabled users on the source
		Try
			{
                if($sourceExchangeOnline){
                    Write-Verbose "Getting all User Mailboxes from Exchange online with recipient filter $strSelectUsers"
                    $colUsers += Invoke-Command -Session $sourceSession -ScriptBlock {param ($strSelectUsers) Get-Recipient -Filter $strSelectUsers -RecipientType UserMailbox,MailUser -ResultSize Unlimited | Get-User -ResultSize Unlimited} -ArgumentList $strSelectUsers -ErrorAction Stop
                }
                else{
                    Write-Verbose "Getting all User Mailboxes from Exchange on-premise with recipient filter $strSelectUsers"
                    $colUsers += Invoke-Command -Session $sourceSession -ScriptBlock {param ($strSelectUsers,$sourceOU) Get-Recipient -Filter $strSelectUsers -RecipientType UserMailbox,MailUser -ResultSize Unlimited -OrganizationalUnit $sourceOU | Get-User -ResultSize Unlimited} -ArgumentList $strSelectUsers,$sourceOU -ErrorAction Stop
                }
            }
		Catch
			{Out-File -InputObject "$(Get-Date -Format MM.dd.yyyy-HH:mm:ss);$($WriteMode);ERROR;Source;$($sourceUser.Split('@')[1]);;Read;;;Error getting users: $($Error[0])" -FilePath $LogFilePath -Append}

		
        Write-Verbose "Found $($colUsers.Count) users in the source organization"
        If ($colUsers.Count -eq 0)
		{
			Write-Verbose "No users found in source organization!"
			Out-File -InputObject "$(Get-Date -Format MM.dd.yyyy-HH:mm:ss);$($WriteMode);WARNING;Source;$($sourceUser.Split('@')[1]);;Read;;;No users found in source tenant $($sourceUser.Split('@')[1])" -FilePath $LogFilePath -Append
			return
		}

        #Store values for later use
		ForEach ($User in $colUsers)
			{
                $arrUserMail += $User.WindowsEmailAddress
                $arrUPN += $User.UserPrincipalName
            
            }

		### ENUMERATE CONTACTS

		# Get all contacts on the target
		Try
			{
                Write-Verbose "Getting all contacts in the target organization"
                $colContacts = Invoke-Command -Session $targetSession -ScriptBlock {Get-Recipient -RecipientType MailContact -ResultSize Unlimited | Get-MailContact -ResultSize Unlimited} -ErrorAction Stop}
		Catch
			{Out-File -InputObject "$(Get-Date -Format MM.dd.yyyy-HH:mm:ss);$($WriteMode);ERROR;Target;$($targetUser.Split('@')[1]);;Read;;;Error getting contacts: $($Error[0])" -FilePath $LogFilePath -Append}

        Write-Verbose "Found $($colContacts.count) contacts in target organization"

        #Store the contact email address in an array for future use
		ForEach ($Contact in $colContacts)
			{$arrContactMail += $Contact.WindowsEmailAddress}

		### FIND CONTACTS TO ADD AND UPDATE

		ForEach ($User in $colUsers)
		{
			If ($arrContactMail -contains $User.WindowsEmailAddress)
			{
				Write-Verbose "Contact found for $($User.WindowsEmailAddress)"
				$colUpdContact += $User
			}
			Else
			{
				Write-Verbose "No contact found for $($User.WindowsEmailAddress)"
				$colAddContact += $User
			}
		}

		### FIND CONTACTS TO DELETE

		ForEach ($contact in $colContacts)
		{
            #If custom attribute has a value, and it is not found in the list of UPNs, and the email address of the contact is not found in the list of
            #user email addresses, then it okay to delete the contact
			If (($contact.customAttribute5) -and ($arrUPN -notcontains $contact.customAttribute5) -and ($arrUserMail -notcontains $contact.WindowsEmailAddress))
			{
				$colDelContact += $($contact.WindowsEmailAddress)
				Write-Verbose "Contact will be deleted for $($contact.WindowsEmailAddress)"
			}
		}
        
        Write-Verbose "Found $($colAddContact.count) contacts to add, $($colUpdContact.count) contacts to update, and $($colDelContact.count) contacts to delete"

		Write-Verbose "Updating $($targetUser.Split("@")[1])"

		### ADDS

		ForEach ($User in $colAddContact)
		{
			Write-Verbose "ADDING contact for $($User.WindowsEmailAddress)"

			$Alias = "c-" + $User.WindowsEmailAddress.Split("@")[0]

            $strAddContact = "Set-Contact $($User.WindowsEmailAddress)"
            
            ForEach ($Attrib in $arrAttribs)
			{
                if($user.$Attrib){
				    $strAddContact += " -$($Attrib) `"$($User.$Attrib)`""
                }
			}

            $addCmd = "Invoke-Command -Session `$targetSession -ScriptBlock {$($strAddContact)}"
							
			
			Out-File -InputObject "$(Get-Date -Format MM.dd.yyyy-HH:mm:ss);$($WriteMode);CHANGE;Target;$($targetUser.Split('@')[1]);$($User.WindowsEmailAddress);Add;;;" -FilePath $LogFilePath -Append
			If ($ReadOnly)
			{
                    #Create the contact object with a target OU if on-premise, otherwise create it without
                    if($targetExchangeOnline){
                        Add-Content -Path $readOnlyFilePath -Value  "New-MailContact -Name $($User.DisplayName) -ExternalEmailAddress $($User.WindowsEmailAddress) -Alias $($Alias)"
                    }
                    else{
					    Add-Content -Path $readOnlyFilePath -Value  "New-MailContact -Name $($User.DisplayName) -ExternalEmailAddress $($User.WindowsEmailAddress) -Alias $($Alias) -OrganizationalUnit $targetOU"
                    }
                    #Write the attributes to the contact, this should be a loop of the various attributes to be written
					Add-Content -Path $readOnlyFilePath -Value $strAddContact
                    #Write the user's UPN to customAttribute5 as a way of tracking that script created the contact, and also to track for deletion
                    Add-Content -Path $readOnlyFilePath -Value  "Set-MailContact $($User.WindowsEmailAddress) -customAttribute5 $($User.userPrincipalName)"

			}
			Else
			{
				Try
				{
                    #Create the contact object with a target OU if on-premise, otherwise create it without
                    if($targetExchangeOnline){
    					# Create Contact Object and set attributes
	    				Invoke-Command -Session $targetSession -ScriptBlock {param ($User,$Alias) New-MailContact -Name $User.DisplayName -ExternalEmailAddress $User.WindowsEmailAddress -Alias $Alias} -ArgumentList $User,$Alias -ErrorAction Stop
                    }
                    else{
                        Invoke-Command -Session $targetSession -ScriptBlock {param ($User,$Alias,$targetOU) New-MailContact -Name $User.DisplayName -ExternalEmailAddress $User.WindowsEmailAddress -Alias $Alias -OrganizationalUnit $targetOU} -ArgumentList $User,$Alias,$targetOU -ErrorAction Stop
		    			
                    }
                    #Write the attributes to the contact, this should be a loop of the various attributes to be written
                    Invoke-Expression $addCmd -ErrorAction Stop
                    #Invoke-Command -Session $targetSession -ScriptBlock {param ($User) Set-Contact $User.WindowsEmailAddress -DisplayName $User.DisplayName -Company $User.Company -FirstName $User.FirstName -MobilePhone $User.MobilePhone -PostalCode $User.PostalCode -LastName $User.LastName -StateOrProvince $User.StateOrProvince -StreetAddress $User.StreetAddress -Phone $User.Phone -Title $User.Title -CountryOrRegion $User.CountryOrRegion -City $User.City -Fax $User.Fax -Office $User.Office -Notes $User.Notes} -ArgumentList $User -ErrorAction Stop
                    #Write the user's UPN to customAttribute5 as a way of tracking that script created the contact, and also to track for deletion
                    Invoke-Command -Session $targetSession -ScriptBlock {param ($User) Set-MailContact $User.WindowsEmailAddress -customAttribute5 $User.userPrincipalName} -ArgumentList $User -ErrorAction Stop
				}
				Catch
					{Out-File -InputObject "$(Get-Date -Format MM.dd.yyyy-HH:mm:ss);$($WriteMode);ERROR;Target;$($targetUser.Split('@')[1]);$($User.WindowsEmailAddress);Add;;;Error creating contact: $($Error[0])" -FilePath $LogFilePath -Append}
			}
		}

		### UPDATES

		ForEach ($User in $colUpdContact)
		{
			Write-Verbose "VERIFYING contact for $($User.WindowsEmailAddress)"
            
            #Filter used to find the target contact object(s)
			$strFilter = "WindowsEmailAddress -eq `"$($User.WindowsEmailAddress)`""
			Try
				{$colContacts = Invoke-Command -Session $targetSession -ScriptBlock {param ($strFilter) Get-Contact -Filter $strFilter} -ArgumentList $strFilter -ErrorAction Stop}
			Catch
				{Out-File -InputObject "$(Get-Date -Format MM.dd.yyyy-HH:mm:ss);$($WriteMode);ERROR;Target;$($targetUser.Split('@')[1]);$($User.WindowsEmailAddress);Find;;;Error getting contact: $($Error[0])" -FilePath $LogFilePath -Append}
			ForEach ($Contact in $colContacts)
			{
                #initialize update string and cmd string
				$strUpdateContact = $null
				$updateCmd = $null
                $strWriteBack = $null
                $writeBackCmd = $null

                #Iterate through attributes and append to the strUpdateContact string if the attribute value has changed
				ForEach ($Attrib in $arrAttribs)
				{
					If ($User.$Attrib -ne $Contact.$Attrib)
					{
                        if($ReadOnly){
						    Add-Content -Path $readOnlyFilePath -Value  "	Changing $Attrib"
						    Add-Content -Path $readOnlyFilePath -Value  "		Before: $($Contact.$Attrib)"
						    Add-Content -Path $readOnlyFilePath -Value  "		After: $($User.$Attrib)"
                        }
						$strUpdateContact += " -$($Attrib) `"$($User.$Attrib)`""
						Out-File -InputObject "$(Get-Date -Format MM.dd.yyyy-HH:mm:ss);$($WriteMode);CHANGE;Target;$($targetUser.Split('@')[1]);$($User.WindowsEmailAddress);Update;$($Contact.$Attrib);$($User.$Attrib);" -FilePath $LogFilePath -Append
					}
				}

                #Check if LegacyExchangeDN has been written back to User object
                $mailContact = Invoke-Command -Session $targetSession -ScriptBlock {param ($contact) Get-MailContact $($contact.WindowsEmailAddress)} -ArgumentList $Contact -ErrorAction Stop
                $x500 = "X500:$($mailContact.LegacyExchangeDN)"
                $userRec = Invoke-Command -Session $sourceSession -ScriptBlock {param ($User) Get-Recipient $($User.WindowsEmailAddress)} -ArgumentList $User -ErrorAction Stop

                if($userRec.emailAddresses -notcontains $x500){
                    if($user.recipientType -eq "MailUser"){
                            if($user.RecipientTypeDetails -eq "RemoteUserMailbox"){
                                #Exchange Online owns the mail properties, so you have to manually add the address in AD
                                $strWriteBack = "Set-ADUser -Identity $($User.samAccountName) -Add @{ProxyAddresses=`"$x500`"} -Server $($user.originatingserver) -Credential `$sourceCredential"
                            }
                            else{
                                $strWriteBack = "Invoke-Command -Session `$sourceSession -ScriptBlock {Set-MailUser $($User.WindowsEmailAddress) -EmailAddresses @{Add=`"$x500`"}}"
                            }
                    }
                    elseif($user.recipientType -eq "UserMailbox"){
                        $strWriteBack = "Invoke-Command -Session `$sourceSession -ScriptBlock {Set-Mailbox $($User.WindowsEmailAddress) -EmailAddresses @{Add=`"$x500`"}}"
                    }
                    else{
                        Out-File -InputObject "$(Get-Date -Format MM.dd.yyyy-HH:mm:ss);$($WriteMode);ERROR;Target;$($targetUser.Split('@')[1]);$($User.WindowsEmailAddress);Update;;$x500;Recipient Type is not MailUser UserMailbox" -FilePath $LogFilePath -Append
                    }
                }

                #If there is anything to update
				If ($strUpdateContact.Length -gt 0)
				{
                    Write-Verbose "Updating attributes for $($User.WindowsEmailAddress)"
                    #Prepend the command for the contact being modified
					$strUpdateContact = "Set-Contact $($User.WindowsEmailAddress)" + $strUpdateContact
					If ($ReadOnly)
						{Add-Content -Path $readOnlyFilePath -Value  $strUpdateContact}
					Else
					{
						Try
						{
                            #Create the complete command and invoke it
							$updateCmd = "Invoke-Command -Session `$targetSession -ScriptBlock {$($strUpdateContact)}"
							Invoke-Expression $updateCmd -ErrorAction Stop
						}
						Catch
							{Out-File -InputObject "$(Get-Date -Format MM.dd.yyyy-HH:mm:ss);$($WriteMode);ERROR;Target;$($targetUser.Split('@')[1]);$($User.WindowsEmailAddress);Update;;;Error updating contact: $($Error[0])" -FilePath $LogFilePath -Append}
					}
				}
                If ($strWriteBack){
                    Write-Verbose "Updating X500 for $($User.WindowsEmailAddress)"
                    Out-File -InputObject "$(Get-Date -Format MM.dd.yyyy-HH:mm:ss);$($WriteMode);CHANGE;Target;$($targetUser.Split('@')[1]);$($User.WindowsEmailAddress);Update;;$x500;" -FilePath $LogFilePath -Append
                    If($ReadOnly){
                        Add-Content -Path $readOnlyFilePath -Value  $strWriteBack
                    }
                    else{
                        Try
						{
							Invoke-Expression $strWriteBack -ErrorAction Stop
						}
						Catch
							{Out-File -InputObject "$(Get-Date -Format MM.dd.yyyy-HH:mm:ss);$($WriteMode);ERROR;Target;$($targetUser.Split('@')[1]);$($User.WindowsEmailAddress);Update;;;Error updating user: $($Error[0])" -FilePath $LogFilePath -Append}
                    }
                }
			}
		}

		### DELETES

		ForEach ($Contact in $colDelContact)
		{
            #Notifiy the deletion is pending and create filter to find the contact
			Write-Verbose "DELETING contact for $($Contact)"
			$strFilter = "WindowsEmailAddress -eq `"$($Contact)`""
			Out-File -InputObject "$(Get-Date -Format MM.dd.yyyy-HH:mm:ss);$($WriteMode);CHANGE;Target;$($targetUser.Split('@')[1]);$($Contact);Delete;;;" -FilePath $LogFilePath -Append

			If ($ReadOnly)
				{Add-Content -Path $readOnlyFilePath -Value "Get-MailContact -Filter $($strFilter) | Remove-MailContact -Confirm:$false"}
			Else
			{
				Try
					{
                        $c = Invoke-Command -Session $targetSession -ScriptBlock {param ($strFilter) Get-MailContact -Filter $strFilter} -ArgumentList $strFilter -ErrorAction Stop
                        WriteObjectInfo -sourceObj $c -outFile $DeletesFilePath
                        Invoke-Command -Session $targetSession -ScriptBlock {param ($strFilter) Get-MailContact -Filter $strFilter | Remove-MailContact -Confirm:$false} -ArgumentList $strFilter -ErrorAction Stop

                    }
				Catch
					{Out-File -InputObject "$(Get-Date -Format MM.dd.yyyy-HH:mm:ss);$($WriteMode);ERROR;Target;$($targetUser.Split('@')[1]);$($Contact);Delete;;;Error deleting contact: $($Error[0])" -FilePath $LogFilePath -Append}
			}
		}
		Remove-PSSession $sourceSession
		Remove-PSSession $targetSession
		$sourceSession = $null
		$targetSession = $null
	}
}

### --- MAIN ---

# If running in Read mode, the default value as a safety feature, write a notification to the console and set the variable for writes to the log
If ($ReadOnly)
{
	Write-Warning "Running in Read Only Mode"
	$WriteMode = "Read-Only"
    Try
    {
	    $readOnlyFilePath = $LogFileDir + "\GalSync-ReadOnly-$(Get-Date -Format yyyy-MM-dd_hh-mm-ss).log"
	    # If the log file doesn't exist, create it and write the headers to it
	    If (-not (Test-Path $readOnlyFilePath))
	    {	
		    New-Item $readOnlyFilePath -Type File  -ErrorAction Stop | Out-Null # Create the log file
		    Out-File -InputObject "ReadOnly Log File Commands" -FilePath $readOnlyFilePath -Append -ErrorAction Stop
	    }
	    # Test to make sure we can access the log file
	    [IO.File]::OpenWrite($readOnlyFilePath).Close()
    }
    # If we get an error creating or accessing the log file, write a warning to the screen and exit the script
    Catch
    {
	    Write-Warning "Cannot access log file `'$($readOnlyFilePath)`'. Exiting script."
	    Exit
    }
}
# Else running in Write mode, set the variable for writes to the log
Else
	{$WriteMode = "WRITE"}

Try
{
	$LogFilePath = $LogFileDir + "\GalSync-$(Get-Date -Format yyyy-MM-dd_hh-mm-ss).log"
	# If the log file doesn't exist, create it and write the headers to it
	If (-not (Test-Path $LogFilePath))
	{	
		New-Item $LogFilePath -Type File  -ErrorAction Stop | Out-Null # Create the log file
		Out-File -InputObject "Timestamp;WriteMode;DebugLevel;SourceorTarget;Domain;User;Action;OldValue;NewValue;Details" -FilePath $LogFilePath -Append -ErrorAction Stop
	}
	# Test to make sure we can access the log file
	[IO.File]::OpenWrite($LogFilePath).Close()
}
# If we get an error creating or accessing the log file, write a warning to the screen and exit the script
Catch
{
	Write-Warning "Cannot access log file `'$($LogFilePath)`'. Exiting script."
	Exit
}

Try
{
	$DeletesFilePath = $LogFileDir + "\GalSync-Deleted-$(Get-Date -Format yyyy-MM-dd_hh-mm-ss).log"
	# If the log file doesn't exist, create it and write the headers to it
	If (-not (Test-Path $DeletesFilePath))
	{	
		New-Item $DeletesFilePath -Type File  -ErrorAction Stop | Out-Null # Create the log file
	}
	# Test to make sure we can access the log file
	[IO.File]::OpenWrite($DeletesFilePath).Close()
}
# If we get an error creating or accessing the log file, write a warning to the screen and exit the script
Catch
{
	Write-Warning "Cannot access log file `'$($DeletesFilePath)`'. Exiting script."
	Exit
}

Out-File -InputObject "$(Get-Date -Format MM.dd.yyyy-HH:mm:ss);$($WriteMode);INFO;;;;;;;Beginning script, running as $($env:USERDOMAIN)\$($env:USERNAME)" -FilePath $LogFilePath -Append

# Run the Function to Sync users from the first to the second tenant
Write-Verbose "$($FirstUser.Split('@')[1]) Users --> $($SecondUser.Split('@')[1]) Contacts"
SyncContacts -sourceUser $FirstUser -sourcePWFile $FirstPWFile -targetUser $SecondUser -targetPWFile $SecondPWFile `
 -sourceExchangeOnline $FirstExchangeOnline -sourceExchangeServer $FirstExchangeServer `
 -targetExchangeOnline $SecondExchangeOnline -targetExchangeServer $SecondExchangeServer `
 -sourceOU $FirstSourceOU -targetOU $SecondTargetOU

# Run the Function to Sync users from the second to the first tenant
Write-Verbose "`n$($SecondUser.Split('@')[1]) Users --> $($FirstUser.Split('@')[1]) Contacts"
SyncContacts -sourceUser $SecondUser -SourcePWFile $SecondPWFile -targetUser $FirstUser -targetPWFile $FirstPWFile `
 -sourceExchangeOnline $SecondExchangeOnline -sourceExchangeServer $SecondExchangeServer `
 -targetExchangeOnline $FirstExchangeOnline -targetExchangeServer $FirstExchangeServer `
 -sourceOU $SecondSourceOU -targetOU $FirstTargetOU

Out-File -InputObject "$(Get-Date -Format MM.dd.yyyy-HH:mm:ss);$($WriteMode);INFO;;;;;;;End of script." -FilePath $LogFilePath -Append
#---------------End of script GALSync.ps1---------------
