<#
	Script to parse through the smtp protocol logs and list out the unique IPs and DNS names
#>
[CmdletBinding()]
Param(
	[Parameter(Mandatory=$true)][string] $dir)
#if you want to exclude IPs (like other Exchange Servers in your org) put them here.
$Exclude=@("10.12.34.10","10.12.34.11","10.10.253.196","10.10.254.139","10.10.254.143","10.13.34.10","10.10.254.20")

$files=Get-ChildItem $dir -Filter *.log
if($files -ne $null){
	$logs=@()
	$output=@()
	foreach ($file in $files) {
	    $logs+=get-content $file.FullName | ?{$_ -notmatch "^#"} | % {$_.Split(",")[5]} | %{$_.Split(":")[0]}
	}
	$logs= $logs | sort-object | get-unique

	foreach($log in $logs){
	    if($Exclude -notcontains $log){
	        $dns=$null
	        Write-Verbose "Working on $log"
	        $objlog = new-object system.object
	        $objlog | add-member -type NoteProperty -name IP -value $log
	        $dns=[System.Net.Dns]::GetHostEntry($log).HostName
	        if($dns -ne $log){
	            $objlog | add-member -type NoteProperty -name DNS-Name -value $dns
	        } else {
	            $dns="Unknown"
	            $objlog | add-member -type NoteProperty -name DNS-Name -value $dns
	        }
	        Write-Verbose "Got this from DNS $dns"
	        $output+=$objlog
	    }
	}
	$output
} else {
	Write-Host "You log Directory ($dir) appears to be empty of log files."
	Write-Host "Check your path and try again."
	Write-Host "Also ensure that your Exchange server Protocol Logging is enabled."
}