$a = "<style>"
$a = $a + "BODY{background-color:White;}"
$a = $a + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
$a = $a + "TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:DarkBlue}"
$a = $a + "TD{border-width: 1px;padding: 10px;border-style: solid;border-color: black;background-color:White}"
$a = $a + "</style>"

Get-Mailbox | Get-MailboxStatistics | where {$_.LastLogonTime -eq $null} | select-object displayname,itemcount,lastlogontime | ConvertTo-Html -head $a -body "<H2>Mailboxes with empty LastLogon Fields</H2>" | out-file neverloggedonmailboxes.html
$mailbody= get-content neverloggedonmailboxes.html | out-string
Send-MailMessage -from Exchange@company.com -to me@company.com -subject "Mailboxes with null logon stamps" -SmtpServer myserverFQDN -Body $mailbody -bodyashtml
