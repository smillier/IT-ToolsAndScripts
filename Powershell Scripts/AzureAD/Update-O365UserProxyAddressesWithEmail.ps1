
    <#
    .SYNOPSIS
        Function to copy email address in the user UPN

    .DESCRIPTION
        Function to copy email address in the user UPN

    .PARAMETER UsersOU
       Specify the OU containing the users to update. Example OU=Users,DC=CONTOSO,DC=COM

    .EXAMPLE
        Update-O365UserProxyAddressesWithEmail `
        -Verbose 
      
    .NOTES
        Sylvain Millier
        github.com/smi
    .LINK
        https://github.com/smillier/IT-ToolsAndScripts/tree/main/Powershell%20Scripts/AzureAD
#>

[CmdletBinding()]
[OutputType([pscustomobject])]
param
(
    [Parameter(Mandatory=$true)]
    [String]$UsersOU
)

    TRY {
        #Get all users having an non empty email
        $OUExists = Get-ADOrganizationalUnit -Filter {distinguishedName -eq $UsersOU}
        
        if ($OUExists -ne $null)
        {
            Write-Host "[BEGIN] Getting users" -ForegroundColor Green
            $users = Get-ADUser -Properties mail,ProxyAddresses -Filter {mail -like '*'} -SearchBase $UsersOU    
            foreach ($user in $users)
            {
                $mail = $user.mail
                $proxyAddresses = $user.ProxyAddresses
                $upn = $user.UserPrincipalName
            
                Write-Host  "[BEGIN] Current user: $mail Current Adresses: $proxyAddresses" -ForegroundColor Yellow
                
                if ($user.ProxyAddresses -notcontains "SMTP:$mail")
                {
                    Write-Host  "Adding proxy address." -ForegroundColor Yellow
                    Set-ADUser -Identity $user -add @{proxyAddresses="SMTP:$mail"}
                    Write-Host "Done." -ForegroundColor Green
                }
                else
                {
                    Write-Host  "User's proxy addresses already contains this address. Nothing to change." -ForegroundColor Green
                }

               
            }          
        } 
        else {
            Write-Host "[ERROR] This OU does not exists in current domain." -ForegroundColor Red
        }
    }

    CATCH {
        $PSCmdlet.ThrowTerminatingError($_)
    }

