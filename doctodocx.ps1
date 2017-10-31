[ref]$SaveFormat = "microsoft.office.interop.word.WdSaveFormat" -as [type] 
$word = New-Object -ComObject word.application 
$word.visible = $false 
$folderpath = Read-Host 'Path?'
$folderpath = $folderpath + '\*'
$fileType = "*doc" 
Get-ChildItem -Force -recurse -path $folderpath -include $fileType -ErrorAction SilentlyContinue | 
foreach-object { 
$path = ($_.fullname).substring(0,($_.FullName).lastindexOf(".")) 
$docxpath =($_.fullname).substring(0,($_.FullName).lastindexOf(".")) + ".docx"
  if (test-path -literal $docxpath) {
}
else {
try {
   write-output "Converting $path to $fileType ..."  
    $doc = $word.documents.open($_.fullname) 
    $doc.saveas([ref] $docxpath, [ref]$SaveFormat::wdFormatDocumentDefault) 
    $doc.close() 

}
catch {
Write-Output "$path couldnt be converted"
} 

}
}
$word.Quit() 
$word = $null 
[gc]::collect() 
[gc]::WaitForPendingFinalizers()