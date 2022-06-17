
param ($servername)


$filterDate=(Get-Date).AddDays(-1)
#Get-VM | Get-VMSnapshot | Where-Object {$_.CreationTime -lt $filterDate}
$i= Invoke-Command -ComputerName $servername -ScriptBlock{Get-VMSnapshot -VMName *}
if ($null -eq $i)
{
	$x=[string]"0:OK"
    write-host $x
    exit 0
}
	else
	{
		$snapsCount = {{$i |  Where-Object {$_.CreationTime -lt $filterDate}} | measure}.Count
		if ($snapsCount  -eq 0) {
			$x=[string]$snapsCount+":OK"
			write-host $x
			exit 0
		}  
		Else 
		{
			if ($snapsCount  -gt 0) {
				$x=[string]$snapsCount+":Snapshots"
				write-host $x
				exit 1
			}  
			else 
			{
				$x=[string]"2:Error checking snapshots"
				write-host $x
				exit 3
			}
		}
}
	
