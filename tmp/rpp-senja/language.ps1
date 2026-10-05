$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$path='D:\RPL\Semester V\Penjenamaan (TRM527)\output\RPP Rebranding Visual Senja Organizer - Revisi.docx'
$backup='D:\RPL\Semester V\tmp\rpp-senja\before-language.docx'
Copy-Item -LiteralPath $path -Destination $backup
$path='D:\RPL\Semester V\Penjenamaan (TRM527)\output\RPP Rebranding Visual Senja Organizer - Revisi Bahasa.docx'
$zip=[IO.Compression.ZipFile]::OpenRead($backup)
$reader=[IO.StreamReader]::new($zip.GetEntry('word/document.xml').Open())
$xml=[xml]$reader.ReadToEnd();$reader.Dispose()
$ns=[Xml.XmlNamespaceManager]::new($xml.NameTable)
$ns.AddNamespace('w','http://schemas.openxmlformats.org/wordprocessingml/2006/main')
$changes=[ordered]@{
'pemahaman terhadap manusia dan cerita di balik perayaan'='pemahaman terhadap karakter, kebutuhan, dan cerita klien'
'Senja memahami manusia dan cerita di balik perayaan, lalu menghubungkannya dengan pilihan konsep serta pengalaman'='Senja memahami karakter, kebutuhan, dan cerita klien sebagai dasar penyusunan konsep acara'
'Senja memahami orang dan cerita di balik perayaan, lalu menerjemahkannya menjadi pengalaman yang personal dan utuh'='Senja memahami karakter dan kebutuhan klien, lalu menerjemahkannya ke dalam konsep acara yang sesuai'
'Kesan yang diuji: manusiawi, empathetic, dan ekspresif secara terarah.'='Kesan yang diuji: personal, empatik, dan ekspresif sesuai karakter klien.'
'Kabel di jalur orang'='Kabel di jalur pengunjung'
}
$count=0
foreach($t in $xml.SelectNodes('//w:t',$ns)){
 foreach($key in $changes.Keys){if($t.InnerText.Contains($key)){$t.InnerText=$t.InnerText.Replace($key,$changes[$key]);$count++}}
}
if($count -ne 5){throw "Expected 5 edits, found $count"}
$stream=[IO.File]::Open($path,[IO.FileMode]::Create)
$out=[IO.Compression.ZipArchive]::new($stream,[IO.Compression.ZipArchiveMode]::Create)
foreach($entry in $zip.Entries){$target=$out.CreateEntry($entry.FullName);$s=$target.Open();if($entry.FullName -eq 'word/document.xml'){$settings=[Xml.XmlWriterSettings]::new();$settings.Encoding=[Text.UTF8Encoding]::new($false);$writer=[Xml.XmlWriter]::Create($s,$settings);$xml.Save($writer);$writer.Dispose()}else{$src=$entry.Open();$src.CopyTo($s);$src.Dispose()};$s.Dispose()}
$out.Dispose();$stream.Dispose();$zip.Dispose()
Write-Output "Updated $count phrases"
