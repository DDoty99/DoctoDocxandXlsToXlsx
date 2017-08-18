
$word = New-Object -ComObject excel.application 
$word.visible = $false 
$folderpath = Read-Host 'Path?'
$folderpath2 = $folderpath + '\*'
$fileType = "*xls" 
Get-ChildItem -recurse -path $folderpath2 -include $fileType | 
foreach-object { 
$path = ($_.fullname).substring(0,($_.FullName).lastindexOf(".")) 
$docxpath =($_.fullname).substring(0,($_.FullName).lastindexOf(".")) + ".xlsx" 
Write-Host "$docxpath"

  if (test-path $docxpath) {Write-Host "Skip $docxpath"
 }
else {
"Converting $path to $fileType ..."
 $doc = $Word.workbooks.open($_.fullname) 
$doc.saveas($path, [Microsoft.Office.Interop.Excel.XLFileFormat]::xlWorkbookDefault) 
$doc.close() 
  }
}
$word.Quit() 
$word = $null 
[gc]::collect() 
[gc]::WaitForPendingFinalizers()