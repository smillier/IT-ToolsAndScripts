function Update-O365UserProxyAddressesWithEmail {
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
            if ($OUExists -isnot  $null)
            {
                Write-Verbose -Message "[BEGIN] Getting users"
                $users = Get-ADUser -Properties mail,ProxyAddresses -Filter {mail -like '*'} -SearchBase $UsersOU    
                foreach ($user in $users)
                {
                    Write-Verbose -Message "[BEGIN] Current user:" $user.mail + " Current Adresses: " + $user.ProxyAddresses
                    Write-Verbose -Message "Adding proxy address."
                        Set-ADUser -Identity $user.UserPrincipalName -add @{proxyAddresses="SMTP:"+ $user.mail}


                    }
                    Write-Host "Done."
                }          
            } 
            else {
                Write-Error "[ERROR] This OU does not exists in current domain."
            }
        }

        CATCH {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
