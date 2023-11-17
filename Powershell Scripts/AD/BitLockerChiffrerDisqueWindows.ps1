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
            echo "La puce TPM est bien pr�sente mais elle doit �tre initialis�e.."
            initialize-tpm
            echo "La puce TPM a bien �t� initialis�e"
            Add-BitLockerKeyProtector -MountPoint $env:SystemDrive -TpmProtector
            Enable-BitLocker -MountPoint $env:SystemDrive -RecoveryPasswordProtector -SkipHardwareTest
        }else{
            echo � La puce TPM est pr�sente et d�j� intialis�e ! �
            Add-BitLockerKeyProtector -MountPoint $env:SystemDrive -TpmProtector
            Enable-BitLocker -MountPoint $env:SystemDrive -RecoveryPasswordProtector -SkipHardwareTest
        }
    }else{
        echo "Il n�y a pas de puce TPM : "
        echo "La s�curisation BitLocker par mot de passe va donc �tre choisi"
        $BitlLockerPwd = ConverTo-SecureString � Arc1400 � -AsPlainText -Force
        Add-BitLockerKeyProtector -MountPoint $env:SystemDrive -PasswordProtector -Password $BitLockerPwd
        Enable-BitLocker -MountPoint $env:SystemDrive -RecoveryPasswordProtector -SkipHardwareTest
    }
}else{
    echo "Le PC est d�j� chiffr� ou en cours de chiffrement"
}