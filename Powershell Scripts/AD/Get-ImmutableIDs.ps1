Import-Module ActiveDirectory
$user = Get-ADUser -Identity [username]
$immutableid = [System.Convert]::ToBase64String($user.ObjectGUID.tobytearray())
$immutableid #(Ca sera la valeur à définir pour le ImmutableID dans Azure)
