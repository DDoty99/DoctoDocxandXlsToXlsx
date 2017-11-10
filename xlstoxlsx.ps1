
$excel = New-Object -ComObject excel.application 
$excel.visible = $false 
$folderpath = Read-Host 'Path?'
$folderpath2 = $folderpath + '\*'
$fileType = "*xls" 
Get-ChildItem -Force -recurse -path $folderpath2 -include $fileType -ErrorAction SilentlyContinue | 
foreach-object { 
$path = ($_.fullname).substring(0,($_.FullName).lastindexOf(".")) 
$xlsxpath =($_.fullname).substring(0,($_.FullName).lastindexOf(".")) + ".xlsx" 
  if (test-path -literal $xlsxpath) {
 }
else {
try {
write-output "Converting $path to $fileType ..."
 $excel = $excel.workbooks.open($_.fullname) 
$excel.saveas($xlsxpath, [Microsoft.Office.Interop.Excel.XLFileFormat]::xlWorkbookDefault) 
$excel.close() 
  }

  
  catch{
  Write-Output "$path couldnt be converted"
  }
}
}
$excel.Quit() 
$excel = $null 
[gc]::collect() 
[gc]::WaitForPendingFinalizers()