$localprograms = choco list
$InstalledProducts = get-wmiobject Win32_Product | fT
if ($InstalledProducts -like "{B211A06A-7EDD-4FD3-AE2C-3B85245BB570}")
{
 Write-Host ("Devolution Launcher not found...")
 exit 0
}
else
{
 Write-Host ("Devolution Launcher already installed.")
 exit 1
}