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

#load the AD tools
Import-Module Activedirectory

#start the transcript
Start-Transcript c:\bin\ControlOAbyGroup.log

#grab all the mailboxes
$mailboxes=Get-Mailbox -ResultSize unlimited
#Grab the memebers of the Allowed OA group
$OAusers=Get-ADGroupMember -Identity "OutlookAnywhereAllowed" | select samaccountname | out-string

#loop through the mailboxes
foreach($mailbox in $mailboxes){
    #If the user is a member of the group then enable OA
    if($OAusers.contains($mailbox.samaccountname)){
        Write-Host "Enalbling OA for" $mailbox.identity
        Set-CASMailbox $mailbox.identity -MAPIBlockOutlookRpcHttp:$False
    }else{
    #Otherwise disable OA for the user
        Write-Host "Disabling OA for" $mailbox.identity
        Set-CASMailbox $mailbox.identity -MAPIBlockOutlookRpcHttp:$True
    }
}

Stop-Transcript