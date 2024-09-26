$adapters = Get-NetAdapterAdvancedProperty -DisplayName 'Idle Power Saving' | Where-Object RegistryValue -eq '1'
foreach ($adapter in $adapters) {
    Write-Host("Disabling powersaving for adapter " + $adapter.DisplayName)
    Set-NetAdapterAdvancedProperty -InterfaceDescription $adapter.InterfaceDescription -DisplayName 'Idle Power Saving' -RegistryValue '0'
}