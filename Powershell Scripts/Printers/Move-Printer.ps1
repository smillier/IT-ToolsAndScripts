<#PSScriptInfo

.VERSION 1.0

.GUID abed7be5-a430-48fb-b5fc-0f622f9e46b5

.AUTHOR Jakub Šindelář

.COMPANYNAME Houby Studio

.COPYRIGHT 2020 Jakub Šindelář

.TAGS Printer Port Move Rename Configure

.LICENSEURI https://opensource.org/licenses/MIT

.PROJECTURI https://gist.github.com/megastary

.ICONURI 

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS 

.EXTERNALSCRIPTDEPENDENCIES 

.RELEASENOTES
 v1.0 - [2020-09-24] - Creates new TCP/IP port (RAW/9100), renames printer and moves printer to new IP address.
#>

<#
.SYNOPSIS
 Creates new TCP/IP port (RAW/9100), renames printer and moves printer to new IP address.

.DESCRIPTION
 Creates new TCP/IP port (RAW/9100), renames printer and moves printer to new IP address. 
 You probably want to tweak it for your use case, for example:
 if you use different language than en-US or want to delete old port or change printer driver afterwards..

.PARAMETER OldName
 Name of the current printer name, which we want to move.

.PARAMETER NewName
 New name for the printer specified with OldName parameter.

.PARAMETER NewPortName
 New port name (IP address), which will be assigned to the printer we have renamed.

.PARAMETER PathToPrinterScripts
 Path to the Printing_Admin_Scripts folder. Use only if OS language other than en-US.

.EXAMPLE
 .\Move-Printer.ps1

.EXAMPLE
 .\Move-Printer.ps1 -OldName 0120PR -NewName 0101PR -NewPortName 192.168.1.101

.EXAMPLE
 .\Move-Printer.ps1 -OldName 0120PR -NewName 0101PR -NewPortName 192.168.1.101 -PathToPrinterScripts "C:\Windows\System32\Printing_Admin_Scripts\cs-CZ"
#>
[CmdletBinding()]
Param(
    [Parameter(Mandatory=$true,
    Position=0)]
    [ValidateNotNullOrEmpty()]
    [String]$OldName,
    [Parameter(Mandatory=$true,
    Position=1)]
    [ValidateNotNullOrEmpty()]
    [String]$NewName,
    [Parameter(Mandatory=$true,
    Position=2)]
    [ValidateNotNullOrEmpty()]
    [String]$NewPortName,
    [Parameter(Mandatory=$false,
    Position=3)]
    [ValidateNotNullOrEmpty()]
    [String]$PathToPrinterScripts = "$($env:SystemRoot)\System32\Printing_Admin_Scripts\en-US" # Change only if OS is in different language.
)

<#
# At our company, we can determine IP address by parsing printer name, so we did not have to manually input new port name.
# Leaving it here commented out as others may find themselves in similar situation and may reuse it.
# Example: Printer 0150PR is parsed by code below as follows:
# [01]: Location number - that is 1
# [15]: Printer number - that is 15
# [PR]: Means it is a printer, we ignore it in this case
# Resulting IP is 192.168.1.115

[int]$LocationNumber = $NewName.Substring(0,2)
[int]$PrinterNumber = $NewName.Substring(2,2)
$PrinterIP = "{0:d2}" -f $PrinterNumber # We want printer with name 0102PR to be on 192.168.1.102 address
$NewPortName = "192.168.$LocationNumber.1$PrinterIP"
#>

# Nearly identical to Start-Process, except we can capture output to variable.
function Start-ProccessWithOutput {
    param (
        $Folder,
        $Script,
        $Arguments
    )
    $ProcessInfo = New-Object System.Diagnostics.ProcessStartInfo
    $ProcessInfo.FileName = "$($env:SystemRoot)\System32\cscript.exe"
    $ProcessInfo.RedirectStandardError = $true
    $ProcessInfo.RedirectStandardOutput = $true
    $ProcessInfo.UseShellExecute = $false
    $ProcessInfo.Arguments = "$(Join-Path -Path $Folder -ChildPath $Script) $Arguments"
    $Process = New-Object System.Diagnostics.Process
    $Process.StartInfo = $ProcessInfo
    $Process.Start() | Out-Null
    $Process.WaitForExit()
    $Output = $Process.StandardOutput.ReadToEnd()
    $Output
}

# Create new port
$NewPort = Start-ProccessWithOutput -Folder $PathToPrinterScripts -Script 'prnport.vbs' -Arguments "-a -r $NewPortName -h $NewPortName -o raw -n 9100"
$Result = $NewPort.Trim().Split([Environment]::NewLine)[$NewPort.Trim().Split([Environment]::NewLine).Length - 1]
if ($Result -ne "Created/updated port $NewPortName") {
    Write-Error "Could not create new port $NewPortName"
}
$Result

# Change printer name
$PrinterName = Start-ProccessWithOutput -Folder $PathToPrinterScripts -Script 'prncnfg.vbs' -Arguments "-x -p $OldName -z $NewName"
$Result = $PrinterName.Trim().Split([Environment]::NewLine)[$PrinterName.Trim().Split([Environment]::NewLine).Length - 1]
if ($Result -ne "New printer name $NewName") {
    Write-Error "Could not rename printer $OldName to $NewName"
}
$Result

# Change printer port
$PrinterPort = Start-ProccessWithOutput -Folder $PathToPrinterScripts -Script 'prncnfg.vbs' -Arguments "-t -p $NewName -r $NewPortName"
$Result = $PrinterPort.Trim().Split([Environment]::NewLine)[$PrinterPort.Trim().Split([Environment]::NewLine).Length - 1]
if ($Result -ne "Configured printer $NewName") {
    Write-Error "Could not change port for printer $NewName"
}
$Result