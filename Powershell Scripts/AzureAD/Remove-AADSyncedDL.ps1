Install-Module -Name MSOnline
Connect-MSOLService

Get-MSOLGroup -Objectid XXXX | Remove-MSOLGroup
