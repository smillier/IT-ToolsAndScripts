$computerName = hostname
$query="Associators of {Win32_Group.Domain='$computerName',Name='Administrators'} where Role=GroupComponent"
$members = @()
$membersIntune = @()

write-verbose "Querying $computerName"
write-verbose $query

$members = Get-CIMInstance -query $query -computer $computerName |
Select @{Name="Member";Expression={$_.Caption}},Disabled,LocalAccount,
@{Name="Type";Expression={([regex]"User|Group").matches($_.Class)[0].Value}},
@{Name="Computername";Expression={$_.ComputerName.ToUpper()}}
$members
foreach ($member in $members)
{
  if($member.Disabled -eq $False)
  {
     $membersIntune += $member.Member
  }
}


Write-Output -InputObject ($membersIntune -join ', ')
    Exit 0

