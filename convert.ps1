$pdf=gci *.pdf
new-item -itemType Directory Result
$pdf|%{
$filePath = $_.FullName
$convertPath = $filePath.replace(".pdf",".docx")

write-host "Converting $filePath..."
$wd = New-Object -ComObject Word.Application 
$wd.Visible = $false
$txt = $wd.Documents.Open( $filePath, $false, $false, $false)

$wd.Documents[1].SaveAs($convertPath) 
$wd.Documents[1].Close()

write-host "$filePath converted to $convertPath."

move-item $filepath Result
move-item $convertPath Result
}
write-host "Completed."
pause