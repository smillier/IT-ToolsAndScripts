$PrinterNames = Get-Printer

foreach ($PrinterName in $PrinterNames ){
  $P = Get-PrinterPort -Name $PrinterName.PortName
  $PrinterName.PortName
  $NewName = $P.Name.ToString().replace(".10.",".20.")
  Add-PrinterPort -Name $NewName -PrinterHostAddress $P.PrinterHostAddress
  Set-Printer -name $PrinterName.Name -PortName $NewName
  Remove-PrinterPort -Name $P.Name
}
