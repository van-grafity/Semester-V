$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$dest='D:\RPL\Semester V\Penjenamaan (TRM527)\output\RPP Rebranding Visual Senja Organizer.docx'
$backup='D:\RPL\Semester V\tmp\rpp-senja\before-compact.docx'
Copy-Item -LiteralPath $dest -Destination $backup
$dest='D:\RPL\Semester V\Penjenamaan (TRM527)\output\RPP Rebranding Visual Senja Organizer - Revisi.docx'
$zip=[IO.Compression.ZipFile]::OpenRead($backup)
$reader=[IO.StreamReader]::new($zip.GetEntry('word/document.xml').Open())
$xml=[xml]$reader.ReadToEnd();$reader.Dispose()
$w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
$ns=[Xml.XmlNamespaceManager]::new($xml.NameTable);$ns.AddNamespace('w',$w)
$body=$xml.SelectSingleNode('//w:body',$ns);$nodes=@($body.ChildNodes)
function SetPara($p,[string]$text){
 $r=$p.SelectSingleNode('w:r',$ns).CloneNode($true)
 foreach($ch in @($r.ChildNodes)){if($ch.LocalName -ne 'rPr'){[void]$r.RemoveChild($ch)}}
 foreach($ch in @($p.ChildNodes)){if($ch.LocalName -ne 'pPr'){[void]$p.RemoveChild($ch)}}
 $t=$xml.CreateElement('w','t',$w);$t.InnerText=$text;[void]$r.AppendChild($t);[void]$p.AppendChild($r)
}
function Cell($row,[string[]]$lines){
 $c=$nodes[1].SelectNodes('w:tr',$ns)[$row].SelectNodes('w:tc',$ns)[2]
 $base=$c.SelectSingleNode('w:p',$ns).CloneNode($true)
 foreach($p in @($c.SelectNodes('w:p',$ns))){[void]$c.RemoveChild($p)}
 foreach($line in $lines){$p=$base.CloneNode($true);SetPara $p $line;[void]$c.AppendChild($p)}
}
Cell 1 @('Happy Yugo Prasetya, S.Sn., M.Sn.')
Cell 2 @('Miftahul Husna Ghawa, S.Tr.Kom.')
Cell 3 @('Rebranding Visual Senja Organizer')
Cell 4 @(
'1. Laporan proyek dan slide presentasi.',
'2. Logo guideline A4 landscape, termasuk sistem identitas visual dan contoh penerapannya.',
'3. Video motion logo berdurasi minimal 10 detik.',
'4. Poster karya ilmiah berukuran A2.',
'5. Dokumen pengajuan HKI dan Berita Acara Serah Terima PBL.',
'6. Dokumen pendukung Branding Project: hasil pengujian logo beserta analisis; copywriting guideline dan tagline; prototipe produk, desain antarmuka, wireframe, dan use case; hasil pengujian usability beserta analisis; dokumen perencanaan, pelaksanaan, dan evaluasi pameran; visualisasi portofolio individu; serta identifikasi dan analisis mitigasi risiko pameran dalam ruangan dengan penerapan K3L.'
)
Cell 6 @('Estimasi Rp1.100.000; rincian pada bagian 7, belum disahkan.')
Cell 8 @('14 minggu')
# Move detailed strategy paragraphs intact to an appendix.
$appendix=$nodes[12].CloneNode($true)
SetPara $appendix 'Lampiran landasan strategi dan arah kreatif'
[void]$body.InsertBefore($appendix,$nodes[106])
foreach($i in 12..31){
 $p=$nodes[$i]
 $break=$p.SelectSingleNode('w:pPr/w:pageBreakBefore',$ns)
 if($break -and $i -eq 12){[void]$break.ParentNode.RemoveChild($break)}
 [void]$body.RemoveChild($p);[void]$body.InsertBefore($p,$nodes[106])
}
SetPara $nodes[9] 'Desain rebranding Senja Organizer berangkat dari gagasan Senja sebagai momen yang berkembang menjadi pengalaman, cerita, dan ingatan. Arah yang diusulkan menempatkan pemahaman terhadap karakter, cerita, kebutuhan, dan preferensi pemilik acara sebagai dasar komunikasi merek. Nilai Understanding, Personalization, Coherence, dan Collaboration diterjemahkan melalui kepribadian Empathetic, Thoughtful, Refined, dan Collaborative.'
# Replace long component table with concise implementation paragraphs.
[void]$body.RemoveChild($nodes[10])
foreach($text in @(
'Perancangan mencakup logo, warna, tipografi, elemen grafis, fotografi, tata letak, dan tone of voice. Sistem identitas diterapkan pada media sosial, katalog layanan, prototipe akademik, serta materi pameran. Identitas Senja tetap konsisten, sementara cerita yang ditampilkan memberi ruang bagi karakter masing-masing klien.',
'Tim mengembangkan minimal tiga alternatif logo melalui riset visual dan eksplorasi konsep, kemudian memilih serta menyempurnakannya berdasarkan hasil pengujian dan masukan mitra. Bentuk, warna, dan jenis huruf ditetapkan setelah audit; hasil akhir dirangkum dalam logo guideline A4 landscape dan aset penerapannya. Rincian visi, misi, nilai, arah kreatif, dan penulisan wara disajikan pada Lampiran landasan strategi dan arah kreatif.'
)){$p=$nodes[9].CloneNode($true);SetPara $p $text;[void]$body.InsertBefore($p,$nodes[11])}
# Log this specific editorial revision.
$table=$nodes[82];$rows=$table.SelectNodes('w:tr',$ns);$new=$rows[2].CloneNode($true)
$texts=@('03 / 5 Oktober 2026','Perapian identitas dan daftar luaran; penyelarasan biaya pada identitas; peringkasan Desain Umum dan pemindahan rincian strategi ke lampiran.','Tim PBL Senja')
for($i=0;$i -lt 3;$i++){SetPara $new.SelectNodes('w:tc',$ns)[$i].SelectSingleNode('w:p',$ns) $texts[$i]}
[void]$table.InsertBefore($new,$rows[3])
$stream=[IO.File]::Open($dest,[IO.FileMode]::Create)
$out=[IO.Compression.ZipArchive]::new($stream,[IO.Compression.ZipArchiveMode]::Create)
foreach($entry in $zip.Entries){$target=$out.CreateEntry($entry.FullName);$s=$target.Open();if($entry.FullName -eq 'word/document.xml'){$settings=[Xml.XmlWriterSettings]::new();$settings.Encoding=[Text.UTF8Encoding]::new($false);$writer=[Xml.XmlWriter]::Create($s,$settings);$xml.Save($writer);$writer.Dispose()}else{$src=$entry.Open();$src.CopyTo($s);$src.Dispose()};$s.Dispose()}
$out.Dispose();$stream.Dispose();$zip.Dispose()
Write-Output $dest
