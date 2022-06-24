# Install-Module MSOnline
Install-Module AzureAD
Connect-AzureAD

Set-AzureADUser -ObjectId xyz@domain.com -ImmutableId [ImmutableID récupéré avant]

