Write-Host "Выберете принтер"
$Row = Read-Host
cls
#Путь к справочнику сотрудников
$PrintList="\\10.1.1.252\Soft\Printers.xls"
#Имя листа (WorkSheet) рабочей книги Excel
$SheetName="Лист1"
#"Запускаем" Excel (создаем COM-объект Excel.Application)
$objExcel=New-Object -comobject Excel.Application
#выполняем открытие файла ("Рабочей книги") в Excel
$objWorkbook=$objExcel.Workbooks.Open($PrintList)
$PrintN = 1
$PrintI = 2
#Выполняем перебор строк в открытом файле Excel
$PrintName=$objWorkbook.ActiveSheet.Cells.Item($Row, $PrintN).Value()
$PrintIP=$objWorkbook.ActiveSheet.Cells.Item($Row, $PrintI).Value()
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
Write-Host "Готово"