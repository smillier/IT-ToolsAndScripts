<#   
Script pour activer le chiffrement BitLocker des disques Windows
ARC IT, NOR - 05.10.2022
#>

Import-Module BitLocker

if((Get-BitlockerVolume -MountPoint $env:SystemDrive).VolumeStatus -eq "FullyDecrypted")
{
    if ((Get-Tpm).TpmPresent -eq "True")
    {
        echo "Il y a une puce TPM"
        if ((Get-Tpm).TpmReady -eq "False")
        {
            echo "La puce TPM est bien présente mais elle doit être initialisée.."
            initialize-tpm
            echo "La puce TPM a bien été initialisée"
            Add-BitLockerKeyProtector -MountPoint $env:SystemDrive -TpmProtector
            Enable-BitLocker -MountPoint $env:SystemDrive -RecoveryPasswordProtector -SkipHardwareTest
        }else{
            echo « La puce TPM est présente et déjà intialisée ! »
            Add-BitLockerKeyProtector -MountPoint $env:SystemDrive -TpmProtector
            Enable-BitLocker -MountPoint $env:SystemDrive -RecoveryPasswordProtector -SkipHardwareTest
        }
    }else{
        echo "Il n’y a pas de puce TPM : "
        echo "La sécurisation BitLocker par mot de passe va donc être choisi"
        $BitlLockerPwd = ConverTo-SecureString « Arc1400 » -AsPlainText -Force
        Add-BitLockerKeyProtector -MountPoint $env:SystemDrive -PasswordProtector -Password $BitLockerPwd
        Enable-BitLocker -MountPoint $env:SystemDrive -RecoveryPasswordProtector -SkipHardwareTest
    }
}else{
    echo "Le PC est déjà chiffré ou en cours de chiffrement"
}