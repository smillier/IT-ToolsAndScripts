# Import active directory module for running AD cmdlets
Import-Module ActiveDirectory
  
# Store the data from NewUsersFinal.csv in the $ADUsers variable
Get-ADUser -Filter 'enabled -eq $true' -SearchBase "OU=USERS,DC=gDOMAINDC=LOCAL" -Properties "*" | Select SamAccountName,UserPrincipalName,HomeDirectory,HomeDrive,scriptPath | Export-Csv -Path "C:\folder\exportHD_LS.csv"


Read-Host -Prompt "Export done - Press Enter to exit..."