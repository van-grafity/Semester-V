$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$src='D:\RPL\Semester V\Penjenamaan (TRM527)\output\RPP Rebranding Visual Senja Organizer - Revisi Bahasa.docx'
$dest='D:\RPL\Semester V\Penjenamaan (TRM527)\output\RPP Rebranding Visual Senja Organizer - Revisi Ruang Lingkup.docx'
$zip=[IO.Compression.ZipFile]::OpenRead($src)
$reader=[IO.StreamReader]::new($zip.GetEntry('word/document.xml').Open())
$xml=[xml]$reader.ReadToEnd();$reader.Dispose()
$ns=[Xml.XmlNamespaceManager]::new($xml.NameTable)
$ns.AddNamespace('w','http://schemas.openxmlformats.org/wordprocessingml/2006/main')
$count=0
foreach($t in $xml.SelectNodes('//w:t',$ns)){
 if($t.InnerText.StartsWith('Cakupan meliputi riset, project brief, creative brief,')){
 $t.InnerText='Ruang lingkup proyek mencakup riset merek dan audiens, penyusunan project brief dan creative brief, pengembangan identitas visual, penerapan pada media komunikasi, pembuatan prototipe, pengujian, serta penyusunan brand guideline. Prototipe UI/UX dan motion logo dikembangkan sebagai bagian dari luaran akademik. Hasil proyek disiapkan untuk pameran dan serah terima kepada mitra.'
 $count++
 }
}
if($count -ne 1){throw "Expected 1 paragraph, found $count"}
$stream=[IO.File]::Open($dest,[IO.FileMode]::Create)
$out=[IO.Compression.ZipArchive]::new($stream,[IO.Compression.ZipArchiveMode]::Create)
foreach($entry in $zip.Entries){$target=$out.CreateEntry($entry.FullName);$s=$target.Open();if($entry.FullName -eq 'word/document.xml'){$settings=[Xml.XmlWriterSettings]::new();$settings.Encoding=[Text.UTF8Encoding]::new($false);$writer=[Xml.XmlWriter]::Create($s,$settings);$xml.Save($writer);$writer.Dispose()}else{$inputStream=$entry.Open();$inputStream.CopyTo($s);$inputStream.Dispose()};$s.Dispose()}
$out.Dispose();$stream.Dispose();$zip.Dispose()
Write-Output $dest
