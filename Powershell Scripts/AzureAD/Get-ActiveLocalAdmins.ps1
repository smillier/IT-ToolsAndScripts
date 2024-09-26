function Get-LocalAdmins {

    <#
        .SYNOPSIS
        Compensate for a known, widespread - but inexplicably unfixed - issue in Get-LocalGroupMember.
        Issue here: https://github.com/PowerShell/PowerShell/issues/2996

        .DESCRIPTION
        The script uses ADSI to fetch all members of the local Administrators group.
        MSFT are aware of this issue, but have closed it without a fix, citing no reason.
        It will output the SID of AzureAD objects such as roles, groups and users,
        and any others which cnanot be resolved.
        the AAD principals' SIDS need to be mapped to identities using MS Graph.

        Designed to run from the Intune MDM and thus accepts no parameters.

        .EXAMPLE
        $results = Get-localAdmins
        $results

        The above will store the output of the function in the $results variable, and
        output the results to console

        .OUTPUTS
        System.Management.Automation.PSCustomObject
        Name        MemberType   Definition
        ----        ----------   ----------
        Equals      Method       bool Equals(System.Object obj)
        GetHashCode Method       int GetHashCode()
        GetType     Method       type GetType()
        ToString    Method       string ToString()
        Computer    NoteProperty string Computer=Workstation1
        Domain      NoteProperty System.String Domain=Contoso
        User        NoteProperty System.String User=Administrator
    #>

    [CmdletBinding()]

    $group1 = [ADSI]"WinNT://$env:COMPUTERNAME/Administrateurs"
    $group2 = [ADSI]"WinNT://$env:COMPUTERNAME/Administrators"
    if ($null -ne $group1)
    { 
        $admins = $group1.Invoke('Members') | ForEach-Object {
            $path = ([adsi]$_).path
            [pscustomobject]@{
                Computer = $env:COMPUTERNAME
                Domain = $(Split-Path (Split-Path $path) -Leaf)
                User = $(Split-Path $path -Leaf)
            }
        }
    }
    else {
        $admins = $group2.Invoke('Members') | ForEach-Object {
            $path = ([adsi]$_).path
            [pscustomobject]@{
                Computer = $env:COMPUTERNAME
                Domain = $(Split-Path (Split-Path $path) -Leaf)
                User = $(Split-Path $path -Leaf)
            }
        }
    }
    return $admins
}

$members = Get-LocalAdmins
$membersIntune = @()

foreach ($member in $members)
{
  
     $membersIntune += $member.User
}


Write-Output -InputObject ($membersIntune -join '; ')
    Exit 0