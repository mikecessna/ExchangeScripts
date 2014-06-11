<#
Script to copy all of the receive connectors from one Exchange server to another.
Sourcce and Destination servers are required commanline parameters.
V1.0
Tested on Exchange 2010
#>
Param(
[Parameter(Mandatory=$true)][string]$SourceServer,
[Parameter(Mandatory=$true)][string]$DestServer
)

$Connectors=ForEach-Object {Get-ReceiveConnector -Server $SourceServer}
ForEach ($Connector in $Connectors) {
    New-ReceiveConnector -Name $Connector.Name -Server $DestServer -Bindings $Connector.Bindings  -RemoteIPRanges $Connector.RemoteIPRanges -AuthMechanism $Connector.AuthMechanism -PermissionGroups $Connector.PermissionGroups
}