$PrinterNames = Import-CSV .\printers.csv

foreach ($PrinterName in $PrinterNames ){
  $P = Get-Printer $PrinterNames.PrinterName
  Remove-PrinterPort -Name $P.PortName
  Add-PrinterPort -Name $PrinterName.NewPortName -PrinterHostAddress $PrinterName.NewPortAddress
  Set-Printer -name $p -PortName $PrinterName.NewPortName
}