# Install-Module MSOnline
Connect-MsolService

Get-MsolUser  | Select-Object UserprincipalName
Set-MsolUser -UserPrincipalName [user@domain.com] -ImmutableId [valeur $immutableid récupéré avant]

