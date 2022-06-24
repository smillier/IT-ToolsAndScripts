Import-Module ActiveDirectory
$users = Get-ADUser -Filter * -SearchBase "OU=Users,DC=domain,DC=local"
$exp_users = @()
$i=0
foreach ($user in $users)
{
    $users[$i].ObjectGUID
    $item = New-Object PSObject
    $item | Add-Member -MemberType NoteProperty -Name "userPrincipalName"  -Value $user.UserprincipalName
    $id = [System.Convert]::ToBase64String($user.ObjectGUID.tobytearray())
    $item | Add-Member -MemberType NoteProperty -Name "immutableID"  -Value $id.ToString()
    $exp_users += $item
    $i++
}
$exp_users | Export-Csv -Path c:\temp\usersimmutableIDs.csv
#$immutableid = [System.Convert]::ToBase64String($user.ObjectGUID.tobytearray())
#$immutableid #(Ca sera la valeur à définir pour le ImmutableID dans Azure)
