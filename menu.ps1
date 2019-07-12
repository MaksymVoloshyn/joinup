#if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }
#Credentials
$ADUserS = "skyup\pcadd"
$PasswordS = 'Support2018'
$credS = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ADUserS,(ConvertTo-SecureString -AsPlainText $PasswordS -Force)
$ADUserJ = "joinup\m.voloshyn"
$PasswordJ = 'Justmax96'
$credJ = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ADUserJ,(ConvertTo-SecureString -AsPlainText $PasswordJ -Force)

#functions

function RenamePC
    {
        Write-Host Введите новое имя ПК
        $name = Read-Host
        Rename-Computer -NewName $name -Force
        }

function DomainAdd
{
cls
Write-Host Select domain:
Write-Host 1. JoinUP -ForegroundColor Green
Write-Host 2. SkyUP  -ForegroundColor Green
Write-Host
$choiced = Read-Host Select the menu item
Switch($choiced){
  1{Add-Computer -DomainName joinup.ua -ComputerName $env:COMPUTERNAME -Credential $credJ}
  2{Add-Computer -DomainName skyup.aero -ComputerName $env:COMPUTERNAME -Credential $credS}
    default {Write-Host ″Wrong choice, try again.″ -ForegroundColor Red}
}
Write-Host Sucsess!
Wait-Event -Timeout 5
Menu
}

function EnbAdmin
{
Enable-LocalUser -Name "Администратор"
Write-Host Sucsess!
Wait-Event -Timeout 5
Menu
}

function delcurrusr
{
Remove-Localuser -Name $env:UserName
Write-Host Sucsess!
Wait-Event -Timeout 5
Menu
}

function clrtpm
{
Clear-TPM
Write-Host Sucsess!
Wait-Event -Timeout 5
Menu
}

function Signature
{
Copy-Item \\10.1.1.252\Signatures\$env:UserName\* -Filter *.htm -destination C:\Users\$env:UserName\AppData\Roaming\Microsoft\Signatures\ -Recurse -Force
Write-Host Sucsess!
Wait-Event -Timeout 5
Menu
}

function printadd
{
cls
Write-Host "Select printer from list:"
#Путь к списку принетров
$PrintTable="\\10.1.1.252\Soft\Printers.xls"
#Имя листа (WorkSheet) рабочей книги Excel
$SheetName="Лист1"
#"Запускаем" Excel (создаем COM-объект Excel.Application)
$objExcel=New-Object -comobject Excel.Application
#выполняем открытие файла ("Рабочей книги") в Excel
$objWorkbook=$objExcel.Workbooks.Open($PrintTable)
$PrintN = 1
$PrintI = 2
#Константа для использования с методом SpecialCells
$xlCellTypeLastCell = 11
#Получаем номер последней используемой строки на листе
$TotalsRow=$objWorkbook.Worksheets.Item($SheetName).UsedRange.SpecialCells($xlCellTypeLastCell).Row
#Выполняем перебор строк в открытом файле Excel
    for ($Row=1;$Row -le $TotalsRow; $Row++) {
        $PrintList=$objWorkbook.ActiveSheet.Cells.Item($Row, $PrintN).Value()
        Write-Host $PrintList
        }
Write-Host "Select item:"; $RowSel = Read-Host
$PrintName=$objWorkbook.ActiveSheet.Cells.Item($RowSel, $PrintN).Value()
$PrintIP=$objWorkbook.ActiveSheet.Cells.Item($RowSel, $PrintI).Value()
$PrintName
#Закрываем книгу Excel
#$objWorkbook.Close()
$objExcel.Workbooks.Close()
#Выходим из Excel (вернее даем команду на выход из Excel)
$objExcel.Quit()
#обнуляем объект
$objExcel = $null
#запускаем принудительную сборку мусора для освобождения памяти и окончательного завершения процесса
[gc]::collect()
[gc]::WaitForPendingFinalizers()
Add-PrinterPort -Name "IP_$PrintIP" -PrinterHostAddress "$PrintIP" 
Add-Printer -Name $PrintName -DriverName "HP LaserJet Pro MFP M521 PCL 6" -PortName "IP_$PrintIP"
(New-Object -ComObject WScript.Network).SetDefaultPrinter("$PrintName")
Write-Host Sucsess!
Wait-Event -Timeout 5
Menu
}

#menu
function Menu 
{
cls
Write-Host Select action:
Write-Host 0. "Rename PC&Add to domain, Enable admin usr, Delete current usr, clear tpm module" -ForegroundColor Green
Write-Host 1. Rename PC -ForegroundColor Green
Write-Host 2. Add to domain -ForegroundColor Green
Write-Host 3. Enable admin user -ForegroundColor Green
Write-Host 4. Delete current user -ForegroundColor Green
Write-Host 5. Clear the TPM module -ForegroundColor Green
Write-Host 6. Add signature -ForegroundColor Green
Write-Host 7. Add default printer -ForegroundColor Green
Write-Host 8. PC reboot -ForegroundColor Green
Write-Host 9. Exit -ForegroundColor Green
Write-Host
$choice = Read-Host Select the menu item
Switch($choice){
  0{RenamePC; DomainAdd; EnbAdmin; delcurrusr; clrtpm;}
  1{RenamePC}
  2{DomainAdd}
  3{EnbAdmin}
  4{delcurrusr}
  5{clrtpm}
  6{Signature}
  7{printadd}
  8{Restart-Computer}
  9{Write-Host ″Exit″; exit}
    default {Write-Host ″Wrong choice, try again.″ -ForegroundColor Red}
}
}
Menu
