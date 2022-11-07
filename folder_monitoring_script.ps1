$watcher = New-Object System.IO.FileSystemWatcher
$watcher.IncludeSubdirectories = $false
$watcher.Path = 'path_to_CSV_export_folder'
$watcher.EnableRaisingEvents = $true

$action =
{

while(Test-Path -Path \\path_to_CSV_export_folder\*.csv)
{
$excel = New-Object -ComObject Excel.Application
$workbook = $excel.workbooks.open("ODBC_connection_v3.xlsm",$null,$true)
$excel.Visible = $false
$excel.DisplayAlerts = $false
$worksheet = $workbook.worksheets.item(1)
$excel.Run("ThisWorkbook.Add_New_Transactions","")
$workbook.close()
$excel.quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
Remove-Variable excel
Stop-Process -processname EXCEL
}

}

Register-ObjectEvent $watcher 'Created' -Action $action
