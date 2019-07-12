if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }
$ADUser = "joinup\m.voloshyn"
$Password = 'Justmax96'
$cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ADUser,(ConvertTo-SecureString -AsPlainText $Password -Force)
Write-Host "Введите новое имя ПК"
$name = Read-Host
Add-Computer -DomainName joinup.ua -Credential $cred -ComputerName $env:COMPUTERNAME -newname $name -Credential $cred
Enable-LocalUser -Name "Администратор"
Remove-Localuser -Name $env:UserName
Clear-TPM
Start-Sleep -Seconds 5
Restart-Computer