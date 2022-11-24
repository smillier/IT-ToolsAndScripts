param(
	[string]$ADWSDC = "localhost",
	[string]$strCriticalGroup = "Domain Admins"
	)

Import-Module ActiveDirectory

$GroupMembers = Get-ADGroupMember $strCriticalGroup -Server $ADWSDC

Write-Host "<prtg>"
Write-Host "<result>" 
"<channel>Users</channel>" 
"<value>"+ $GroupMembers.count +"</value>" 
"</result>"
"<text>" + $GroupMembers.count + "members in " + $strCriticalGroup + "</text>"
Write-Host "</prtg>"