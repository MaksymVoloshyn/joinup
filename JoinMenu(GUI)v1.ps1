if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe " -NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }

Add-Type -assembly System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#Credentials
$ADUserS = "skyup\pcadd"
$PasswordS = 'Support2018'
$credS = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ADUserS,(ConvertTo-SecureString -AsPlainText $PasswordS -Force)
$ADUserJ = "joinup\pcadd"
$PasswordJ = 'Support2018'
$credJ = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ADUserJ,(ConvertTo-SecureString -AsPlainText $PasswordJ -Force)

#GUI
#form
$main_form = New-Object System.Windows.Forms.Form
$main_form.Text ='JoinMenu'
$main_form.Width = 250
$main_form.Height = 400
$main_form.AutoSize = $false

#functions

function DomainAddJoin
	{Add-Computer -DomainName joinup.ua -ComputerName $env:COMPUTERNAME -newname $name -Force -Credential $credJ
	Write-Host Sucsess!
	}
	
function DomainAddSky
 	{Add-Computer -DomainName skyup.aero -ComputerName $env:COMPUTERNAME -newname $name -Force -Credential $credS
	Write-Host Sucsess!
	}

function EnbAdmin
{
Enable-LocalUser -Name "Администратор"
Write-Host Sucsess!
}

function delcurrusr
{
if ($env:UserName -eq "Администратор") {$wshell.Popup("Это учётная запись администратора!")}
else {
Remove-Localuser -Name $env:UserName
Write-Host Sucsess!
}
}

function clrtpm
{
Clear-TPM
Write-Host Sucsess!
}

function Signature
{
Copy-Item \\10.1.1.252\Signatures\$env:UserName\* -Filter *.htm -destination C:\Users\$env:UserName\AppData\Roaming\Microsoft\Signatures\ -Recurse -Force
Write-Host Sucsess!
}

#printaddcombo
$ComboBox = New-Object System.Windows.Forms.ComboBox
$ComboBox.Location  = New-Object System.Drawing.Point(10,280)
$main_form.Controls.Add($ComboBox)

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
		$PrintComp="$Row. $Printlist"
		$ComboBox.Items.add($PrintComp)
        }
		
function printadd
{
$RowSel = $ComboBox.SelectedIndex+1
$PrintName=$objWorkbook.ActiveSheet.Cells.Item($RowSel, $PrintN).Value()
$PrintIP=$objWorkbook.ActiveSheet.Cells.Item($RowSel, $PrintI).Value()
Add-PrinterPort -Name "IP_$PrintIP" -PrinterHostAddress "$PrintIP" 
Add-Printer -Name $PrintName -DriverName "HP LaserJet Pro MFP M521 PCL 6" -PortName "IP_$PrintIP"
(New-Object -ComObject WScript.Network).SetDefaultPrinter("$PrintName")
Write-Host $RowSel
Write-Host Sucsess!
}

#GoButtonFunc
function Go
{
$name = $PCName.Text
if ($AddtoDomRadioJoin.Checked) 		{$AddtoDomRadioJoin.ForeColor='Green'; DomainAddJoin}
if ($AddtoDomRadioSky.Checked)  		{$AddtoDomRadioSky.ForeColor='Green';DomainAddSky}
if ($EAdmUsr.Checked -eq $true) 		{$EAdmUsr.ForeColor='Green'; EnbAdmin}
if ($DCurUsr.Checked -eq $true) 		{$DCurUsr.ForeColor='Green'; delcurrusr}
if ($ClrTPM.Checked -eq $true) 			{$ClrTPM.ForeColor='Green'; clrtpm}
if ($SignatureAdd.Checked -eq $true) 	{$SignatureAdd.ForeColor='Green'; Signature}
if ($ComboBox.SelectedIndex -gt -1) 	{$ComboBox.ForeColor='Green'; printadd}
if ($PCreboot.Checked -eq $true)		{Restart-Computer -Force}

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
}

#GUI

#PCName
$PCName = New-Object System.Windows.Forms.TextBox
$PCName.Location  = New-Object System.Drawing.Point(10,10)
$PCName.Text = 'Enter new PC name'
$main_form.Controls.Add($PCName)

#Select Domine
$Label = New-Object System.Windows.Forms.Label
$Label.Text = "Select the domine:"
$Label.Location  = New-Object System.Drawing.Point(10,40)
$Label.AutoSize = $true
$main_form.Controls.Add($Label)

#Select Print to install
$Labe2 = New-Object System.Windows.Forms.Label
$Labe2.Text = "Select printer to install:"
$Labe2.Location  = New-Object System.Drawing.Point(10,260)
$Labe2.AutoSize = $true
$main_form.Controls.Add($Labe2)

#addtodomaineradiojoin
$AddtoDomRadioJoin = New-Object System.Windows.Forms.RadioButton
$AddtoDomRadioJoin.Location  = New-Object System.Drawing.Point(10,60)
$AddtoDomRadioJoin.Text = 'JoinUP'
$AddtoDomRadioJoin.AutoSize = $true
$main_form.Controls.Add($AddtoDomRadioJoin)

#addtodomaineradiosky
$AddtoDomRadioSky = New-Object System.Windows.Forms.RadioButton
$AddtoDomRadioSky.Location  = New-Object System.Drawing.Point(10,80)
$AddtoDomRadioSky.Text = 'SkyUp'
$AddtoDomRadioSky.AutoSize = $true
$main_form.Controls.Add($AddtoDomRadioSky)

#Select Options
$Labe2 = New-Object System.Windows.Forms.Label
$Labe2.Text = "Select options:"
$Labe2.Location  = New-Object System.Drawing.Point(10,120)
$Labe2.AutoSize = $true
$main_form.Controls.Add($Labe2)

#Enable admin user
$EAdmUsr = New-Object System.Windows.Forms.CheckBox
$EAdmUsr.Text = 'Enable admin user'
$EAdmUsr.AutoSize = $true
$EAdmUsr.Checked = $false
$EAdmUsr.Location  = New-Object System.Drawing.Point(10,140)
$main_form.Controls.Add($EAdmUsr)

#Delete current user
$DCurUsr = New-Object System.Windows.Forms.CheckBox
$DCurUsr.Text = 'Delete current user'
$DCurUsr.AutoSize = $true
$DCurUsr.Checked = $false
$DCurUsr.Location  = New-Object System.Drawing.Point(10,160)
$main_form.Controls.Add($DCurUsr)

#Clear the TPM module
$ClrTPM = New-Object System.Windows.Forms.CheckBox
$ClrTPM.Text = 'Clear the TPM module'
$ClrTPM.AutoSize = $true
$ClrTPM.Checked = $false
$ClrTPM.Location  = New-Object System.Drawing.Point(10,180)
$main_form.Controls.Add($ClrTPM)

#PC_reboot
$PCreboot = New-Object System.Windows.Forms.CheckBox
$PCreboot.Text = 'Reboot the PC'
$PCreboot.AutoSize = $true
$PCreboot.Checked = $false
$PCreboot.Location  = New-Object System.Drawing.Point(10,200)
$main_form.Controls.Add($PCreboot)

#Signature_Add
$SignatureAdd = New-Object System.Windows.Forms.CheckBox
$SignatureAdd.Text = 'Add signature (domine users only)'
$SignatureAdd.AutoSize = $true
$SignatureAdd.Checked = $false
$SignatureAdd.Location  = New-Object System.Drawing.Point(10,220)
$main_form.Controls.Add($SignatureAdd)

#GoButton
$GoButton = New-Object System.Windows.Forms.Button
$GoButton.AutoSize = $true
$GoButton.Text = 'Go'
$GoButton.Location = New-Object System.Drawing.Point(80,320)
$GoButton.Add_Click({Go})
$main_form.Controls.Add($GoButton)

$main_form.ShowDialog() #Показать форму
