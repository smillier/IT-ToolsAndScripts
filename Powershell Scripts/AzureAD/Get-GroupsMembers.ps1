[cmdletbinding()]
Param([string]$computer=$env:computername)

$query="Associators of {Win32_Group.Domain='$computer',Name='Administrators'} where Role=GroupComponent"

write-verbose "Querying $computer"
write-verbose $query

Get-CIMInstance -query $query -computer $computer |
Select @{Name="Member";Expression={$_.Caption}},Disabled,LocalAccount,
@{Name="Type";Expression={([regex]"User|Group").matches($_.Class)[0].Value}},
@{Name="Computername";Expression={$_.ComputerName.ToUpper()}}