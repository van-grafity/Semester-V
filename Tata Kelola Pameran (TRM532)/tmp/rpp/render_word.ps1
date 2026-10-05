$ErrorActionPreference = 'Stop'
$rppRoot = 'D:\RPL\Semester V\Tata Kelola Pameran (TRM532)'
$rppWord = New-Object -ComObject Word.Application
$rppWord.Visible = $false
$rppWord.DisplayAlerts = 0
Write-Output 'Word instance ready'
try {
    Write-Output 'Opening RPP'
    $rppDoc = $rppWord.Documents.Open((Join-Path $rppRoot 'output\RPP Rebranding Senja Wedding Organizer.docx'), $false, $true)
    Write-Output 'RPP opened'
    $rppDoc.Repaginate()
    Write-Output 'Pagination done'
    $rppDoc.ExportAsFixedFormat((Join-Path $rppRoot 'tmp\rpp\render\rpp.pdf'),17)
    $rppDoc.Close(0)
    Write-Output 'Export complete'
} finally {
    $rppWord.Quit()
}
