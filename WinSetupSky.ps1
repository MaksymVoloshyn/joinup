if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }
$ADUser = "skyup\pcadd"
$Password = 'Support2018'
$cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ADUser,(ConvertTo-SecureString -AsPlainText $Password -Force)
Write-Host "������� ����� ��� ��"
$name = Read-Host
Add-Computer -DomainName skyup.aero -Credential $cred -ComputerName $env:COMPUTERNAME -newname $name -Credential $cred
Enable-LocalUser -Name "�������������"
Remove-Localuser -Name $env:UserName
Clear-TPM
Start-Sleep -Seconds 5
Restart-Computer