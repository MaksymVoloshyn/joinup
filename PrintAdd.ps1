Add-PrinterPort -Name "IP_10.1.12.23" -PrinterHostAddress "10.1.12.23" 
Add-Printer -Name HRAdmin -DriverName "HP LaserJet Pro MFP M521 PCL 6" -PortName IP_10.1.12.23
(New-Object -ComObject WScript.Network).SetDefaultPrinter('HRAdmin')