$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$root = 'D:\RPL\Semester V'
$source = Join-Path $root 'Tata Kelola Pameran (TRM532)\output\RPP Rebranding Senja Wedding Organizer.docx'
$outdir = Join-Path $root 'Penjenamaan (TRM527)\output'
[void][IO.Directory]::CreateDirectory($outdir)
$dest = Join-Path $outdir 'RPP Rebranding Visual Senja Organizer.docx'
$before = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
$zip = [IO.Compression.ZipFile]::OpenRead($source)
$reader = [IO.StreamReader]::new($zip.GetEntry('word/document.xml').Open())
$xml = [xml]$reader.ReadToEnd()
$reader.Dispose()
$ns = [Xml.XmlNamespaceManager]::new($xml.NameTable)
$w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
$ns.AddNamespace('w',$w)
$body = $xml.SelectSingleNode('//w:body',$ns)
$nodes = @($body.ChildNodes)
function TextOf($n) { return (($n.SelectNodes('.//w:t',$ns) | ForEach-Object { $_.InnerText }) -join '') }
function SetPara($p,[string]$text) {
    $r = $p.SelectSingleNode('w:r',$ns)
    if ($null -eq $r) { $r = $xml.CreateElement('w','r',$w) }
    $newr = $r.CloneNode($true)
    foreach($ch in @($newr.ChildNodes)) { if($ch.LocalName -ne 'rPr') { [void]$newr.RemoveChild($ch) } }
    foreach($ch in @($p.ChildNodes)) { if($ch.LocalName -ne 'pPr') { [void]$p.RemoveChild($ch) } }
    $t = $xml.CreateElement('w','t',$w)
    $t.InnerText = $text
    [void]$newr.AppendChild($t)
    [void]$p.AppendChild($newr)
}
function SetCell($table,[int]$row,[int]$col,[string]$text) {
    $c = $table.SelectNodes('w:tr',$ns)[$row].SelectNodes('w:tc',$ns)[$col]
    $p = $c.SelectSingleNode('w:p',$ns)
    SetPara $p $text
    foreach($other in @($c.SelectNodes('w:p',$ns) | Select-Object -Skip 1)) { [void]$c.RemoveChild($other) }
}
function InsertText($anchor,[string]$text,[bool]$heading=$false,[bool]$page=$false) {
    $p = $(if($heading){$nodes[13].CloneNode($true)}else{$nodes[14].CloneNode($true)})
    SetPara $p $text
    if($page){$b=$xml.CreateElement('w','pageBreakBefore',$w);[void]$p.SelectSingleNode('w:pPr',$ns).AppendChild($b)}
    [void]$body.InsertBefore($p,$anchor)
}
SetPara $nodes[1] 'Rebranding Visual Senja Organizer'
SetCell $nodes[2] 3 2 'Rebranding Visual Senja Organizer'
SetCell $nodes[2] 7 2 'Senja Organizer; brief awal menyebut Senja Event & Wedding Planner Batam dan CV. Multi Art Project. Nama resmi, PIC, serta kontak dikonfirmasi kepada mitra.'
SetCell $nodes[2] 8 2 '14 minggu pembelajaran; evaluasi mengikuti kalender akademik 2026/2027.'
SetPara $nodes[5] 'Proyek Rebranding Visual Senja Organizer mengembangkan kembali identitas merek agar mampu merepresentasikan karakter, nilai, dan pengalaman yang ingin dibangun Senja. Perancangan berangkat dari gagasan bahwa setiap acara merupakan momen dengan cerita dan karakter pemiliknya. Identitas dikembangkan sebagai sistem visual dan verbal yang saling terhubung pada media komunikasi, prototipe akademik, dan pameran hasil PBL.'
SetPara $nodes[6] 'Masalah yang akan diteliti adalah kesesuaian identitas dan komunikasi Senja saat ini dengan pengalaman yang diharapkan klien. Tim mengaudit aset, mewawancarai mitra serta audiens, dan membandingkan kompetitor untuk mengidentifikasi kesenjangan persepsi. Klaim tentang ketidakkonsistenan visual, rendahnya interaksi, atau kelemahan layanan belum diperlakukan sebagai temuan tanpa bukti.'
SetPara $nodes[7] 'Pertanyaan proyek: bagaimana merancang sistem identitas visual Senja Organizer yang mengomunikasikan pemahaman terhadap manusia dan cerita di balik perayaan, serta dapat diterapkan secara konsisten pada berbagai media? Keberhasilan dinilai melalui keterbacaan, kesesuaian asosiasi dengan positioning dan personality, koherensi penerapan, serta pemahaman informasi oleh audiens.'
SetPara $nodes[10] 'Desain umum menggunakan gagasan SENJA = MOMENT dengan hubungan Moment, Experience, Story, dan Memory. Visi yang diusulkan ialah mengubah momen yang direncanakan hari ini menjadi cerita yang layak dikenang setelah perayaan berakhir. Senja diarahkan untuk memahami karakter, cerita, kebutuhan, dan preferensi pemilik acara, lalu menerjemahkannya menjadi pengalaman yang memiliki karakter tersendiri. Arah ini merupakan proposed rebranding yang harus divalidasi bersama mitra dan melalui riset.'
SetCell $nodes[11] 1 1 'Hipotesis audiens dari brief awal: calon pengantin 22–35 tahun di Batam–Kepri, kelas menengah, aktif di Instagram; keluarga ikut memberi pertimbangan. Segmen dan kebutuhan diuji melalui riset.'
SetCell $nodes[11] 2 1 'Understanding, Personalization, Coherence, Collaboration menjadi nilai usulan. Positioning menekankan proses memahami orang dan cerita sebelum mengembangkan pengalaman acara.'
SetCell $nodes[11] 2 2 'Validasi kesesuaian dengan praktik Senja dan kebutuhan audiens; analisis kompetitor sebelum menetapkan diferensiasi atau USP.'
SetCell $nodes[11] 3 1 'Minimal tiga alternatif logo dalam satu landasan strategi; sistem warna, tipografi, grafis, fotografi, dan layout. Pilihan simbol dan palet ditentukan setelah audit dan eksplorasi.'
SetCell $nodes[11] 4 1 'Prototipe katalog layanan: pengenalan Senja, pendekatan memahami klien, layanan, portofolio dengan konteks cerita, dan simulasi konsultasi kebutuhan.'
SetCell $nodes[11] 5 1 'Alur booth memperlihatkan hubungan temuan riset, strategi, alternatif, pengujian, hasil identitas, dan penerapan; pengunjung dapat mencoba prototipe serta memberi tanggapan.'
InsertText $nodes[13] 'Landasan strategi merek' $true $true
InsertText $nodes[13] 'Visi: Mengubah momen yang direncanakan hari ini menjadi cerita yang layak dikenang setelah perayaan berakhir.'
InsertText $nodes[13] 'Misi 1: Memahami karakter, kebutuhan, dan preferensi setiap klien sebagai dasar pengembangan konsep. Misi 2: Mengembangkan pengalaman acara yang relevan dengan cerita dan karakter pemilik acara. Misi 3: Menghindari pendekatan satu konsep untuk semua dalam perencanaan dan pelaksanaan.'
InsertText $nodes[13] 'Misi 4: Menyatukan konsep, visual, detail, dan pengalaman tamu menjadi satu karakter acara yang konsisten. Misi 5: Berkolaborasi dengan vendor untuk menjaga keutuhan konsep dari perencanaan hingga pelaksanaan.'
InsertText $nodes[13] 'Understanding diwujudkan melalui mendengar dan menggali kebutuhan. Personalization menjadikan pemahaman terhadap klien sebagai dasar keputusan. Coherence menghubungkan seluruh detail pada cerita utama. Collaboration menjaga kesamaan pemahaman antara klien, Senja, dan vendor.'
InsertText $nodes[13] 'Personality yang dituju adalah Empathetic, Thoughtful, Refined, dan Collaborative. Implikasinya: komunikasi yang mendengar, keputusan desain beralasan, hierarki visual yang jelas dan tidak berlebihan, serta hubungan kerja yang terbuka dengan arahan yang tetap tegas.'
InsertText $nodes[13] 'Proses usulan: Understand, Translate, Personalize, Unify, Deliver; Remember merupakan hasil yang diharapkan. Prinsip internal “This feels like us” dipakai untuk menilai relevansi pengalaman bagi klien. Memorability merupakan hasil, bukan core value atau USP yang telah terbukti. Pernyataan konsep dalam referensi belum menjadi tagline final.'
InsertText $nodes[13] 'Identitas Senja dan identitas acara klien' $true
InsertText $nodes[13] 'Logo, aturan warna, tipografi, dan gaya komunikasi Senja harus konsisten. Personalisasi diterapkan pada isi cerita dan ekspresi acara sesuai kebutuhan pemiliknya; tidak berarti mengganti identitas Senja untuk setiap klien. PBL menghasilkan sistem identitas dan contoh penerapan, bukan pelaksanaan acara klien atau perubahan operasional layanan yang belum disepakati.'
InsertText $nodes[13] 'Alternatif arah kreatif' $true $true
InsertText $nodes[13] 'Arah A — Moments into Memories. Fokus pada pengalaman yang meninggalkan ingatan. Eksplorasi visual menelaah jejak momen, hubungan antarbagian cerita, dan fotografi interaksi yang bermakna. Kesan yang diuji: intim, thoughtful, dan refined.'
InsertText $nodes[13] 'Arah B — Every Story Its Own Moment. Fokus pada personalisasi. Eksplorasi menelaah sistem bingkai atau elemen grafis adaptif yang memberi ruang bagi cerita berbeda, dengan identitas Senja sebagai pengikat. Kesan yang diuji: manusiawi, empathetic, dan ekspresif secara terarah.'
InsertText $nodes[13] 'Arah C — The Art of Transition. Fokus pada perjalanan dan peralihan momen. Eksplorasi menelaah ritme, urutan, ruang, dan perubahan bentuk tanpa kewajiban menggunakan simbol sunset. Kesan yang diuji: thoughtful, refined, dan berkembang secara utuh.'
InsertText $nodes[13] 'Ketiga arah merupakan alternatif eksplorasi kelompok yang sama, bukan pembagian anggota menjadi Tim A, B, dan C. Semuanya menggunakan visi, nilai, dan positioning yang sama. Warna, jenis huruf, dan bentuk logo belum dikunci. Logo lama dinilai berdasarkan ekuitas, keterbacaan, relevansi, dan temuan audit sebelum diputuskan untuk dipertahankan, disempurnakan, atau diubah.'
InsertText $nodes[13] 'Pemilihan arah menggunakan bukti riset, kesesuaian personality, pembeda terhadap kompetitor, keterbacaan, serta kelayakan penerapan. Setiap keputusan menjelaskan alasan, sumber, pembeda, penerjemahan ke visual/verbal, dan status buktinya.'
SetPara $nodes[14] 'Komunikasi menguraikan cara Senja memahami manusia dan cerita di balik perayaan, lalu menghubungkannya dengan pilihan konsep serta pengalaman. Pesan didukung contoh proses atau portofolio yang dapat diverifikasi dan diizinkan mitra. Usulan key message: Senja memahami orang dan cerita di balik perayaan, lalu menerjemahkannya menjadi pengalaman yang personal dan utuh.'
SetPara $nodes[16] 'Voice diusulkan empatik, penuh pertimbangan, jelas, dan kolaboratif. Tone lebih informatif pada layanan, lebih naratif pada cerita portofolio, dan lebih mengundang dialog pada konsultasi. Pesan menghindari sentimentalitas berlebihan, janji kesempurnaan, serta klaim keunggulan tanpa bukti. Naskah final disusun dan diuji sesuai materi Penulisan Wara.'
SetCell $nodes[22] 1 1 'Wawancara mitra tentang sejarah, praktik memahami klien dan koordinasi vendor; audit identitas; riset audiens dan kompetitor; inventaris aset serta izin. Keluaran: bukti riset dan project brief. Belum menentukan logo atau gaya final.'
SetCell $nodes[22] 2 1 'Analisis masalah, SWOT, audiens, persona, persepsi saat ini dan yang dituju; validasi visi, misi, values, personality, positioning serta proposed USP. Keluaran: creative brief, kriteria keberhasilan dan batas lingkup. Bedakan fakta, asumsi, dan usulan.'
SetCell $nodes[22] 3 1 'Eksplorasi kata kunci dan metafora dari strategi; moodboard/stylescape tiga arah; sketsa dan minimal tiga alternatif logo. Jelaskan alasan warna, tipografi, grafis, fotografi, layout dan voice. Rancang alur prototipe dan konsep booth; uji awal untuk memperoleh masukan.'
SetCell $nodes[22] 4 1 'Kembangkan logo terpilih beserta variasi, sistem identitas, aplikasi media sosial dan media acara, contoh komunikasi, serta prototipe interaktif. Siapkan guideline A4 landscape, motion logo minimal 10 detik, poster A2, dan materi booth.'
SetCell $nodes[22] 5 1 'Uji keterbacaan, pengenalan dan asosiasi identitas dengan positioning/personality; uji pemahaman pesan dan usability prototipe. Analisis hasil, revisi, uji ulang, finalisasi guideline serta paket aset. Dokumentasikan keterbatasan dan evaluasi pameran.'
SetPara $nodes[25] 'Logo diuji pada 16–24 px dan cetak 10–15 mm, monokrom, pengenalan singkat dan kemiripan dengan merek lain. Pengujian persepsi meminta peserta mendeskripsikan kesan sebelum melihat kata sasaran, lalu menilai kesesuaian dengan Empathetic, Thoughtful, Refined, dan Collaborative beserta alasannya. Bandingkan identitas awal dan usulan dengan urutan penyajian bervariasi. Penilaian tidak hanya menanyakan desain yang paling disukai.'
InsertText $nodes[27] 'Kriteria kerja usulan: minimal tiga alternatif beserta rasional tersedia; logo terpilih terbaca pada ukuran uji; semua aplikasi menaati guideline; mayoritas peserta mampu menjelaskan pesan pemahaman klien/personalitas acara dengan kata sendiri. Laporkan jumlah peserta dan jawaban aktual, bukan hanya persentase. Keberhasilan branding jangka panjang atau pertumbuhan penjualan tidak disimpulkan dari uji formatif ini.'
SetCell $nodes[36] 2 1 'Analisis audit, SWOT/kompetitor, audiens dan persona; validasi visi, misi, values, personality, proposed positioning/USP; creative brief serta kriteria penerimaan.'
SetCell $nodes[36] 3 1 'Eksplorasi tiga arah kreatif berbasis Moment; minimal tiga alternatif logo; formula wara dan voice, struktur prototipe dan wireframe; uji awal serta konsep pameran.'
SetCell $nodes[59] 2 1 'Apakah arah story-led sesuai praktik dan kemampuan Senja? Nilai, positioning, batas personalisasi, ruang lingkup, dan kriteria apa yang disepakati?'
SetCell $nodes[63] 1 1 'Keterbacaan, hubungan dengan values/personality dan positioning, koherensi aplikasi, pemahaman pesan serta bukti klaim.'
SetPara $nodes[84] 'Landasan strategi usulan mengikuti referensi.md. Profil dan segmentasi dari brief awal tetap menjadi bahan validasi; tidak otomatis merupakan hasil riset. Angka pengikut, engagement, konversi, harga dan hasil layanan memerlukan bukti bertanggal. Mandatori logo sunset, palet golden hour, dan tagline pada brief awal dibahas kembali bersama mitra berdasarkan audit dan strategi terbaru.'
SetCell $nodes[85] 4 1 'Senja diarahkan sebagai organizer yang memulai perancangan dari karakter, cerita, kebutuhan dan preferensi pemilik acara, lalu menerjemahkannya menjadi pengalaman yang personal dan utuh.'
SetCell $nodes[85] 4 2 'Validasi wawancara klien, contoh proses layanan, kebutuhan audiens dan kompetitor. Tetap sebagai proposed positioning sampai terbukti.'
SetCell $nodes[85] 5 1 'Eksplorasi Moment, Experience, Story, Memory melalui tiga arah kreatif. Bentuk, warna dan tipografi mengikuti strategi serta hasil pengujian; tidak dikunci pada sunset, emas/cokelat atau serif mewah.'
SetCell $nodes[85] 5 2 'Audit identitas lama, bandingkan alternatif, dan uji keterbacaan, asosiasi, orisinalitas serta konsistensi penerapan. Pertahankan aset lama jika ada alasan berbasis bukti.'
SetCell $nodes[85] 6 1 'Konten menggambarkan pemahaman kebutuhan, cerita di balik keputusan konsep, hubungan detail, koordinasi vendor, serta informasi layanan dan konsultasi. Tone mengikuti konteks, dengan voice empatik dan kolaboratif.'
SetCell $nodes[85] 6 2 'Minta peserta menjelaskan pendekatan Senja dan hubungannya dengan kebutuhan klien. Gunakan hanya kisah, foto dan testimoni yang terverifikasi serta berizin.'
SetPara $nodes[90] 'Acuan strategi: referensi.md, khususnya bagian 1–15 dan 19–24. Acuan identitas anggota dan profil awal: Creative-Brief-dan-Project-Brief-Senja-WO.docx. Ketentuan visual dan tagline brief awal diperbarui melalui validasi mitra; tidak dianggap telah disetujui dalam RPP ini. Harga layanan hanya dicantumkan setelah dikonfirmasi. Prototipe digunakan sebagai demonstrasi akademik.'
$revrow=$nodes[67].SelectNodes('w:tr',$ns)[1].CloneNode($true)
[void]$nodes[67].InsertAfter($revrow,$nodes[67].SelectNodes('w:tr',$ns)[1])
SetCell $nodes[67] 2 0 '02 / 5 Oktober 2026'
SetCell $nodes[67] 2 1 'Penyelarasan strategi dengan referensi.md: Senja = Moment, visi/misi, values/personality, tiga arah kreatif, Design Thinking dan pengujian persepsi. Mandatori visual lama menjadi bahan validasi.'
SetCell $nodes[67] 2 2 'Tim PBL Senja; diajukan untuk reviu'
# Apply a real title style without changing the source page furniture.
$style=$xml.CreateElement('w','pStyle',$w);$style.SetAttribute('val',$w,'Title')
$pp=$nodes[0].SelectSingleNode('w:pPr',$ns)
$old=$pp.SelectSingleNode('w:pStyle',$ns);if($old){[void]$pp.RemoveChild($old)}
[void]$pp.PrependChild($style)
# Keep all original package parts except the intended document body byte-identical.
$stream=[IO.File]::Open($dest,[IO.FileMode]::Create)
$newzip=[IO.Compression.ZipArchive]::new($stream,[IO.Compression.ZipArchiveMode]::Create)
foreach($entry in $zip.Entries){
    $target=$newzip.CreateEntry($entry.FullName)
    $ts=$target.Open()
    if($entry.FullName -eq 'word/document.xml'){
        $settings=[Xml.XmlWriterSettings]::new();$settings.Encoding=[Text.UTF8Encoding]::new($false)
        $writer=[Xml.XmlWriter]::Create($ts,$settings);$xml.Save($writer);$writer.Flush();$writer.Dispose()
    }else{$es=$entry.Open();$es.CopyTo($ts);$es.Dispose()}
    $ts.Dispose()
}
$newzip.Dispose();$stream.Dispose();$zip.Dispose()
if((Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash -ne $before){throw 'Source changed'}
$contract=@"
# RPP template contract
Reference: $source
SHA256: $before
Output: $dest
Edit document.xml only; preserve every other part byte-for-byte, including styles, header, footer, theme, image and section geometry.
Retain 13 numbered sections, six-member roster, 14-week plan, provisional budget and blank approval blocks.
Rewrite scope, general design, work stages, strategy-related schedule/validation/appendix cells. Insert strategy and three creative directions using cloned paragraph patterns. Add revision 02.
Source reference remains unchanged. Compare package contents, render through Word if packaged renderer unavailable, inspect all final pages.
"@
[IO.File]::WriteAllText((Join-Path $root 'tmp\rpp-senja\artifact.md'),$contract)
Write-Output $dest
