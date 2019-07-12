Write-Host "�������� �������"
$Row = Read-Host
cls
#���� � ����������� �����������
$PrintList="\\10.1.1.252\Soft\Printers.xls"
#��� ����� (WorkSheet) ������� ����� Excel
$SheetName="����1"
#"���������" Excel (������� COM-������ Excel.Application)
$objExcel=New-Object -comobject Excel.Application
#��������� �������� ����� ("������� �����") � Excel
$objWorkbook=$objExcel.Workbooks.Open($PrintList)
$PrintN = 1
$PrintI = 2
#��������� ������� ����� � �������� ����� Excel
$PrintName=$objWorkbook.ActiveSheet.Cells.Item($Row, $PrintN).Value()
$PrintIP=$objWorkbook.ActiveSheet.Cells.Item($Row, $PrintI).Value()
$PrintName
#��������� ����� Excel
#$objWorkbook.Close()
$objExcel.Workbooks.Close()
#������� �� Excel (������ ���� ������� �� ����� �� Excel)
$objExcel.Quit()
#�������� ������
$objExcel = $null
#��������� �������������� ������ ������ ��� ������������ ������ � �������������� ���������� ��������
[gc]::collect()
[gc]::WaitForPendingFinalizers()
Add-PrinterPort -Name "IP_$PrintIP" -PrinterHostAddress "$PrintIP" 
Add-Printer -Name $PrintName -DriverName "HP LaserJet Pro MFP M521 PCL 6" -PortName "IP_$PrintIP"
(New-Object -ComObject WScript.Network).SetDefaultPrinter("$PrintName")
Write-Host "������"