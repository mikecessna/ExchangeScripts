<# Script to change all email addresses by supplying a new and old domain to the script
#>
#set up the command line arguments
# require an Old and New Domain name
Param(
[Parameter(Mandatory=$true)][string]$OldDomain,
[Parameter(Mandatory=$true)][string]$NewDomain,
[switch]$MakeChanges
)

#set up a counter var
$counter=1
$output=@()
#Set up Mail variables
$smtpserver="10.10.10.10"
$from="emailadmiin@yourcompany.com"
$copylocation="e:\bin\copy.htm"

# if using the MakeChanges Switch the HTML Copy file is required
#Get the HTML Copy and Exit if you can't find the copy file
if ($MakeChanges) {
    If (Test-Path $copylocation) {
        $copy=Get-Content $copylocation | out-string
    } else {
        Write-Host "Can't find HTML Copy file at: $copylocation" -ForegroundColor "Red"
        Write-Host "Check your files and try again." -ForegroundColor "Red"
        Exit
    }
}

#Load the Exchange Module
If ("$env:ExchangeInstallPath\bin\RemoteExchange.ps1") {
    if (!(Get-PSSnapin | where {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.E2010"}))
        {
        	Write-Verbose "Loading the Exchange snapin"
        	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue
        	. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
        	Connect-ExchangeServer -auto -AllowClobber
        }
    } else {
            write-host "Exchange Management shell not found. You must have the shell installed."
}

#region Functions
#Functions go here
Function ChangeAddresses {
    param($mailbox)
    Write-Host "================================"
    Write-Host "Working on " $mailbox.DisplayName
    #get email addresses that match $OldDomain
    $addresses= $mailbox.EmailAddresses | ?{$_ -match $OldDomain}
    Write-Host $addresses.count " addresses to change"
    #$addresses
    $newPrimaryAddress=""
    $newAddresses=@()
    foreach ($address in $addresses) {
        #grab the primary address which starts with SMTP:
        #push it into a holding var for later use
        #strip all addresses of the smtp: prefix, swap the domain names, and put them in a var
        if ($address -cmatch "SMTP:") {
            $newPrimaryAddress=($address -replace('smtp:','') -replace($OldDomain,$NewDomain))
            $newAddresses+=($address -replace('smtp:','') -replace($OldDomain,$NewDomain))
        } else {
            $newAddresses+=($address -replace('smtp:','') -replace($OldDomain,$NewDomain))
        }
    }
    #Now add the new email addresses to the mailbox
    Write-Host "New Addresses: " $newAddresses
    if ($type -eq 1) {
        Set-Mailbox -identity $mailbox.Identity -EmailAddresses @{'Add'=@($newAddresses)}
    } elseif ($type -eq 2) {
        Set-DistributionGroup -identity $mailbox.Identity -EmailAddresses @{'Add'=@($newAddresses)}
    }
    #set the primary to the new address
    if ($newPrimaryAddress) {
        Write-Host "NewPrimary: " $newPrimaryAddress
        if ($type -eq 1) {
            Set-Mailbox -identity $mailbox.Identity -PrimarySmtpAddress $newPrimaryAddress
        } elseif ($type -eq 2) {
            Set-DistributionGroup -identity $mailbox.Identity -PrimarySmtpAddress $newPrimaryAddress
        }
        #Send a mail message to the user if a usermailbox
        #don't send for DL changes
        if ($type -eq 1) { SendMail($mailbox)}
    }
    Write-Host "================================"
}

Function SendMail{
    param($mailbox)
    ##########construct email Body
    #To get the first and last name split the displayname field and remove any spaces
    #The first element in the array is the lastname
    #and the second element is firstname
	#an example would be Doe, John
	#change this depending on your display names or use the givenname and surname fields
    $fullname= ($mailbox.displayname -split(',')) -replace(' ','')
    $firstname=$fullname[1]
    $lastname=$fullname[0]
    #Replace the holding text in the html code
    $body= $copy -replace 'varFirstName' ,$firstname
    $body= $body -replace 'varLastName',$lastname
    $body= $body -replace 'varNewAddress',$newPrimaryAddress
    $body= $body -replace 'varAllEmails',($newaddresses -join '<br>')
    #Set Subject
    $subject="Primary Email Address Change"
    #Send mail message
    Write-Host "Sending mail to $newPrimaryAddress"
    Send-MailMessage -from $from -To $newPrimaryAddress -subject $subject -SmtpServer $smtpserver -Body $body -BodyAsHtml
}
#endregion


#get all of the mailboxes with email addresses that match the OldDomain
$mailboxes=Get-Mailbox -ResultSize unlimited | ?{$_.emailaddresses -match $OldDomain}
$dls=Get-DistributionGroup -ResultSize unlimited | ?{$_.emailaddresses -match $OldDomain}

#Check if in Report Mode.
if (!$MakeChanges) { Write-Host "Running in Report mode. To make changes use the -MakeChanges switch"}

#loop through the mailboxes and replace the old domain with the new one
foreach ($mailbox in $mailboxes) {
    if ($MakeChanges) {
        $type=1
        ChangeAddresses($mailbox)
    } Else {
        #get email addresses that match $OldDomain
        $addresses= $mailbox.EmailAddresses | ?{$_ -match $OldDomain}
        foreach ($address in $addresses) {
            $obj=New-Object PSObject
            $newAddress=$address -replace($OldDomain,$NewDomain)
            $obj | Add-Member NoteProperty MailBox($mailbox)
            $obj | Add-Member NoteProperty OldAddress($address -replace("smtp:",""))
            $obj | Add-Member NoteProperty NewAddress($newAddress -replace("smtp:",""))
            $output+=$obj
        }
    }
}
#loop through the DLs and replace the old domain with the new one
foreach ($dl in $dls) {
    if ($MakeChanges) {
        $type=2
        ChangeAddresses($dl)
    } Else {
        #get email addresses that match $OldDomain
        $addresses= $dl.EmailAddresses | ?{$_ -match $OldDomain}
        foreach ($address in $addresses) {
            $obj=New-Object PSObject
            $newAddress=$address -replace($OldDomain,$NewDomain)
            $obj | Add-Member NoteProperty MailBox($dl)
            $obj | Add-Member NoteProperty OldAddress($address -replace("smtp:",""))
            $obj | Add-Member NoteProperty NewAddress($newAddress -replace("smtp:",""))
            $output+=$obj
        }
    }
}
$output