# Import active directory module for running AD cmdlets
Import-Module ActiveDirectory
  
# Store the data from NewUsersFinal.csv in the $ADUsers variable
$ADUsers = Import-Csv C:\File\NewUsersFinal.csv -Delimiter ";"

# Define UPN
$UPN = "domain.local"

# Loop through each row containing user details in the CSV file
foreach ($User in $ADUsers) {
    $User    
    #Read user data from each field in each row and assign the data to a variable as below
    $username = $User.newusername
    $password = $User.password
    $firstname = $User.name
    $lastname = $User.surname
    $OU = "OU=ImportUsers,DC=domain,DC=local"
   

    # Check to see if the user already exists in AD
    if (Get-ADUser -F { SamAccountName -eq $username }) {
        
        # If user does exist, give a warning
        Write-Warning "A user account with username $username already exists in Active Directory."
    }
    else {

        # User does not exist then proceed to create the new user account
        # Account will be created in the OU provided by the $OU variable read from the CSV file
        New-ADUser `
           -SamAccountName $username `
           -UserPrincipalName "$username@$UPN" `
           -Name "$firstname $lastname" `
            -GivenName $firstname `
            -Enabled $True `
            -Path $OU `
           -DisplayName "$lastname, $firstname" `
           -AccountPassword (ConvertTo-secureString $password -AsPlainText -Force) -ChangePasswordAtLogon $True

        # If user is created, show message.
        Write-Host "The user account $username is created." -ForegroundColor Cyan
    }
}

Read-Host -Prompt "Press Enter to exit"