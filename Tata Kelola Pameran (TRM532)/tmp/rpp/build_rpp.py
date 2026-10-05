import sys, json, hashlib
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from copy import deepcopy
sys.path.insert(0, str(Path(__file__).parent/'deps'))
from lxml import etree as E

ROOT=Path(__file__).resolve().parents[2]
REF=Path(r'D:\RPL\Semester V\TRM528 RPP.docx')
OUT=ROOT/'output'/'RPP Rebranding Senja Wedding Organizer.docx'
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS={'w':W}
def el(tag,**attrs):
    n=E.Element('{'+W+'}'+tag)
    for k,v in attrs.items():n.set('{'+W+'}'+k,str(v))
    return n
z=ZipFile(REF); parts={n:z.read(n) for n in z.namelist()}
doc=E.fromstring(parts['word/document.xml']); body=doc.find('w:body',NS)
src=[deepcopy(n) for n in body]; tables=[n for n in src if n.tag==el('tbl').tag]
sect=deepcopy(src[-1])
def para(text='',bold=False,size=10,center=False,page=False):
    p=deepcopy(src[8]); p.clear(); pp=el('pPr');p.append(pp)
    pp.append(el('pStyle',val='BodyText'));pp.append(el('spacing',before=0,after=100,line=240,lineRule='auto'))
    pp.append(el('jc',val='center' if center else 'left'))
    if page:pp.append(el('pageBreakBefore'))
    r=el('r');rp=el('rPr');r.append(rp);rp.append(el('rFonts',ascii='Times New Roman',hAnsi='Times New Roman'));rp.append(el('sz',val=int(size*2)));rp.append(el('color',val='000000'))
    if bold:rp.append(el('b'))
    for i,line in enumerate(text.split('\n')):
        if i:r.append(el('br'))
        t=el('t');t.text=line;r.append(t)
    p.append(r);return p
def add(text='',**kw):body.append(para(text,**kw))
def heading(num,title,page=False):
    p=para(f'{num}.  {title}',True,11,page=page);p.find('w:pPr',NS).append(el('keepNext'));body.append(p)
def sub(title):
    p=para(title,True);p.find('w:pPr',NS).append(el('keepNext'));body.append(p)
def table(headers,rows,widths,source=7,size=10):
    t=deepcopy(tables[source]);pr=t.find('w:tblPr',NS);grid=t.find('w:tblGrid',NS)
    for n in list(t):
        if n not in [pr,grid]:t.remove(n)
    for n in list(grid):grid.remove(n)
    tw=[int(v*567) for v in widths]
    for width in tw:grid.append(el('gridCol',w=width))
    for tag in ['tblW','tblInd','tblLayout','tblBorders','tblCellMar','tblpPr']:
        for n in pr.findall('w:'+tag,NS):pr.remove(n)
    pr.append(el('tblW',w=sum(tw),type='dxa'));pr.append(el('tblInd',w=0,type='dxa'));pr.append(el('tblLayout',type='fixed'))
    borders=el('tblBorders')
    for side in ['top','left','bottom','right','insideH','insideV']:borders.append(el(side,val='single',sz=4,color='808080'))
    pr.append(borders);mar=el('tblCellMar')
    for side in ['top','bottom','left','right']:mar.append(el(side,w=65,type='dxa'))
    pr.append(mar)
    allrows=([headers] if headers else [])+rows
    for ri,values in enumerate(allrows):
        tr=el('tr');trpr=el('trPr');tr.append(trpr);trpr.append(el('cantSplit'))
        ishead=headers is not None and ri==0
        if ishead:trpr.append(el('tblHeader'))
        for ci,text in enumerate(values):
            tc=el('tc');tcp=el('tcPr');tc.append(tcp);tcp.append(el('tcW',w=tw[ci],type='dxa'));tcp.append(el('vAlign',val='center'))
            if ishead:tcp.append(el('shd',fill='D9D9D9'))
            p=para(str(text),ishead,size);p.find('w:pPr/w:spacing',NS).set('{'+W+'}after','0');tc.append(p);tr.append(tc)
        t.append(tr)
    body.append(t);add('',size=3)

# Record the reference contract and package inventory before authoring.
contract=f'''# Template execution contract
Reference: {REF}
SHA256: {hashlib.sha256(REF.read_bytes()).hexdigest()}
Reference render: tmp/rpp/reference/template.pdf (12 pages).
One A4 portrait section, 11920 x 16840 twips; margins top 2120, right 708, bottom 280, left 992; header 969, footer 0.
Authority: retained template No.FO.8.6.1-V1 Rencana Pelaksanaan Proyek, 3 November 2023.
Preserve header, logo, styles, numbering, theme, section geometry and all package parts except word/document.xml.
Typography: Times New Roman, black, source Body Text 10 pt; table headers gray; bordered tables.
Editable slots: body child 4 identity; 7-8 scope; 10-18 design; 22-32 construction; 39-43 equipment; 48-53 risk; 57-61 schedule; 66-67 budget; 73-76 team; 79-82 workspace; 86-93 curriculum; 96-97 communication; 100-101 monitoring; 105-106 revisions; 115-126 approval.
Replace all old E-Koperasi content, team rosters, unsupported officials and approval QR appearances. Keep approval fields unsigned. Retain design-thinking figure at body child 24. Other body drawings are obsolete layout/approval artifacts.
Reuse source table components with resized grids for new row counts and readable wrapping. Keep the 13 numbered sections in order. Append reference and validation notes after approvals.
Pagination may change as substantive content replaces the sample. All content is editable. Unknown administrative details remain explicitly pending.
Fidelity gates: unchanged section geometry and header; all non-document package bytes identical; retained reference hash unchanged; no legacy project text or old signature drawings; all final pages visually inspected.
'''
Path(__file__).with_name('artifact.md').write_text(contract,encoding='utf8')
Path(__file__).with_name('package-inventory.json').write_text(json.dumps({n:{'bytes':len(b),'sha256':hashlib.sha256(b).hexdigest(),'policy':'editable' if n=='word/document.xml' else 'preserve'} for n,b in parts.items()},indent=2),encoding='utf8')
for n in list(body):body.remove(n)

add('Rencana Pelaksanaan Proyek',bold=True,size=14,center=True)
add('Rebranding Visual Senja Wedding Organizer',bold=True,size=12,center=True)
table(None,[
['Nomor ID',':','Kode PBL resmi menunggu konfirmasi manajer proyek'],
['Pengusul proyek',':','Happy Yugo Prasetiya, S.Sn., M.Sn. (sesuai isian awal template; konfirmasi penugasan)'],
['Manajer proyek',':','Miftahul Husna Ghawa, S.Tr.Kom. (sesuai isian awal template; konfirmasi penugasan)'],
['Judul proyek',':','Rebranding Visual Senja Wedding Organizer'],
['Luaran',':','1. Laporan proyek dan lampiran pendukung.\n2. Logo guideline A4 landscape dan hasil pengujian logo.\n3. Video motion logo minimal 10 detik.\n4. Poster karya ilmiah A2.\n5. Copywriting guideline dan tagline untuk media sosial dan website.\n6. Prototipe produk interaktif, desain antarmuka, wireframe, use case, dan analisis pengujian usability.\n7. Dokumen perencanaan, pelaksanaan, evaluasi pameran serta mitigasi risiko K3L.\n8. Visualisasi portofolio individu.\n9. Draf dokumen HKI, BAST PBL, dan slide presentasi.'],
['Sponsor',':','Belum ditetapkan'],['Biaya',':','Estimasi awal Rp1.100.000; rincian pada bagian 7'],
['Klien atau pelanggan',':','Senja Wedding Organizer; nama pemilik/PIC dan kontak resmi dikonfirmasi'],
['Waktu',':','14 minggu pembelajaran; evaluasi mengikuti kalender akademik 2026/2027']
],[4.45,.6,12.35],source=0)
heading(1,'Ruang lingkup')
add('Proyek ini merancang ulang identitas visual dan komunikasi merek Senja Wedding Organizer, lalu menerapkannya pada media sosial, prototipe website layanan, dan pameran hasil PBL. Pekerjaan diawali dengan audit identitas yang digunakan saat ini serta wawancara mitra dan calon pengguna untuk menetapkan masalah, kebutuhan, target audiens, dan kriteria keberhasilan.')
add('Hipotesis masalah yang diuji: calon pasangan kesulitan menilai layanan Senja ketika tampilan promosi belum konsisten, batas layanan belum mudah dipahami, dan alur dari portofolio menuju konsultasi belum jelas. Audit media dan wawancara digunakan untuk menerima, memperbaiki, atau menolak hipotesis ini.')
add('Pertanyaan proyek: bagaimana identitas visual, penulisan wara, dan pengalaman media digital Senja dapat dirancang secara konsisten agar calon pelanggan memahami karakter merek, informasi layanan, serta langkah menghubungi tim Senja?')
add('Cakupan meliputi riset, project brief, creative brief, eksplorasi desain, produksi prototipe, pengujian, revisi, pameran, dan serah terima aset. Pengembangan sistem transaksi, pembayaran, dan pemesanan operasional tidak termasuk cakupan awal. Penambahan fitur harus dicatat sebagai perubahan lingkup.')

heading(2,'Desain umum',page=True)
add('Arah awal yang diajukan adalah identitas yang hangat, tenang, dan profesional. Kesan tersebut akan diuji terhadap nilai merek dan preferensi audiens; keputusan logo, warna, tipografi, serta tagline ditetapkan melalui creative brief dan persetujuan mitra. Rebranding tidak otomatis berarti mengganti seluruh elemen lama: elemen yang masih relevan dapat dipertahankan berdasarkan hasil audit.')
table(['Komponen','Rancangan untuk Senja','Dasar keputusan'],[
['Audiens dan kebutuhan','Asumsi: pasangan usia 24–32 tahun di Batam, bekerja, merencanakan pernikahan dalam 6–12 bulan; keluarga ikut memberi pertimbangan.','Kebutuhan utama: rincian cakupan kerja, alur koordinasi, portofolio relevan dan jalur konsultasi yang jelas. Validasi melalui riset.'],
['Strategi merek','Rumuskan brand value, positioning, diferensiasi/USP dan pesan utama.','Wawancara mitra, audit merek, analisis SWOT dan kompetitor; jangan menyatakan keunggulan tanpa bukti.'],
['Identitas visual','Minimal tiga alternatif logo; palet warna, tipografi, elemen grafis, dan penerapan lintas media.','Moodboard/stylescape dan prinsip sederhana, mudah diingat, fleksibel, sesuai konteks, orisinal, serta seimbang.'],
['Media digital','Usulan struktur: Beranda, Tentang Senja, Layanan, Portofolio dan Kontak. Alur utama: kenali merek, pahami layanan, lihat karya, hubungi Senja.','Sitemap dan wireframe mengikuti temuan UX. Fitur konsultasi disimulasikan dalam prototipe.'],
['Pameran','Alur booth: identitas awal dan temuan riset, konsep rebranding, hasil akhir, prototipe interaktif, umpan balik.','Kurasi relevansi, kualitas, kelayakan teknis, kredit aset, aksesibilitas, dan keselamatan.']
],[3.2,7.2,7.0],source=1)
sub('Arahan penulisan wara')
add('Layanan wedding organizer terutama bersifat intangible. Komunikasi diarahkan pada pemahaman proses dan kepercayaan, didukung bukti yang dapat ditinjau seperti portofolio berizin dan penjelasan layanan yang sudah diverifikasi. Sasaran pesan adalah pengalaman dan branding, bukan klaim hasil yang belum terbukti.')
add('Enam unsur yang disusun tim adalah identitas merek, headline, subheadline, body copy, tagline, dan call to action. AIDA digunakan sebagai formula utama pada pengenalan merek; PAS dapat digunakan pada konten masalah dan solusi. FAB hanya dipakai setelah fitur dan manfaat layanan terverifikasi. Seluruh naskah diperiksa dengan 4C: clear, concise, compelling, credible.')
add('Pola storytelling dirancang dengan hook, conflict, turning point, solution, dan CTA. Persona penulis diarahkan persuasif-naratif pada pengenalan merek serta informatif pada rincian layanan. Ego audiens yang diuji ialah kebutuhan fungsional dan emosional. Voice konsisten hangat dan jelas, sedangkan tone disesuaikan dengan media dan konteks.')
add('Tim merumuskan premis masalah, fakta dasar, dan arah solusi; memilih retorika yang relevan; serta menguji curiosity gap yang terjawab oleh isi konten. Pesan inti diadaptasi untuk cetak, audio/video, dan digital. Naskah serta tagline final disusun mahasiswa dan direviu sesuai ketentuan tugas Penulisan Wara.')

heading(3,'Konstruksi produk',page=True)
figure=deepcopy(src[24]);body.append(figure)
add('Gambar 1. Metode Design Thinking',center=True,size=9)
add('Lima fase Design Thinking digunakan secara iteratif. Kerangka CDIO menghubungkan riset pada Conceive, pengembangan konsep pada Design, produksi pada Implement, dan pengujian penggunaan serta pameran pada Operate. Pengujian dilakukan sejak alternatif desain tersedia.')
table(['Fase','Uraian pekerjaan dan bukti keluaran'],[
['Empathize','Wawancara mitra dan calon pelanggan; audit logo, media promosi dan pengalaman informasi; identifikasi aset serta izin penggunaannya. Keluaran: catatan riset, inventaris aset dan project brief.'],
['Define','Rumusan masalah berbasis temuan; target audiens dan persona pengguna; kebutuhan fungsional/emosional; positioning, USP dan tujuan komunikasi. Keluaran: creative brief, kriteria penerimaan dan batas lingkup.'],
['Ideate','Susun kata kunci, moodboard/stylescape, sketsa dan minimal tiga alternatif logo. Rancang formula wara, premis, voice, sitemap, use case, wireframe serta konsep booth. Keluaran: alternatif dengan alasan pemilihan dan umpan balik awal.'],
['Prototype','Kembangkan logo terpilih, guideline A4 landscape, motion logo minimal 10 detik, rancangan konten, antarmuka dan prototipe interaktif. Siapkan poster A2, label karya, technical rider, denah booth dan SOP.'],
['Test','Uji keterbacaan dan asosiasi logo; pemahaman pesan/CTA; serta tugas pengguna pada prototipe. Dokumentasikan peserta, instrumen, hasil, keterbatasan dan revisi. Lakukan uji teknis booth, evaluasi pengunjung dan serah terima aset.']
],[3.2,14.2],source=3)
sub('Rancangan pengujian')
add('Logo diuji pada ukuran kecil 16–24 px dan cetak 10–15 mm, versi monokrom, pengenalan singkat, serta kesesuaian asosiasi dengan positioning. Wara diuji dengan meminta peserta menjelaskan pesan utama, layanan, dan tindakan yang diminta. Usability diuji melalui tugas menemukan layanan, membuka portofolio dan menemukan kontak pada prototipe.')
add('Usulan uji formatif melibatkan 5–8 calon pengguna yang sesuai kriteria rekrutmen. Catat keberhasilan tugas, kesalahan, waktu penyelesaian, komentar dan temuan revisi. Jumlah tersebut merupakan rencana kerja, bukan sampel representatif pasar. Target awal: minimal 80% peserta menuntaskan setiap tugas inti tanpa bantuan; seluruh masalah kritis diperbaiki dan diuji ulang.')

heading(4,'Kebutuhan peralatan dan bahan',page=True)
table(['Fase','Peralatan atau perangkat','Jumlah','Bahan atau komponen','Catatan'],[
['Seluruh fase','Laptop/PC dan internet','6 unit; 1 akses tim','Folder bersama, logbook dan backlog','Utamakan perangkat yang tersedia; cek akses dan cadangan.'],
['Riset','Ponsel/perekam; formulir survei','1–2 unit; 1 formulir','Panduan wawancara, persetujuan partisipan, catatan','Dokumentasi dilakukan dengan izin.'],
['Desain','Perangkat lunak vektor; Figma/Canva sesuai kebutuhan','Akses sesuai PIC','Sketsa, moodboard, font dan aset berlisensi','Utamakan lisensi kampus atau opsi yang dapat digunakan tim.'],
['Produksi','Perangkat lunak motion/video dan pengolah dokumen','Akses sesuai PIC','Master logo, naskah mahasiswa, storyboard','Simpan file sumber dan hasil ekspor.'],
['Pengujian','Laptop dan ponsel; formulir uji','1 set uji','Skenario tugas, data anonim, daftar revisi','Uji pada ukuran layar yang relevan.'],
['Pameran','Layar/laptop, meja, dudukan, kabel dan pengaman','Usulan 1 set booth','1 poster A2, label, QR dan bahan pemasangan','Jumlah final mengikuti denah dan izin peminjaman.']
],[2.0,4.5,2.2,4.3,4.4],source=4,size=9)
heading(5,'Tantangan dan isu')
add('Register awal berikut menggunakan skala kerja sementara: kemungkinan (a) dan keparahan (b) masing-masing 1–3; total a × b; rendah 1–2, sedang 3–4, tinggi 6–9. Penilaian harus divalidasi melalui borang No.FO.17.1.1-V0 dan survei lokasi. Angka berikut merupakan estimasi awal tim, bukan hasil asesmen resmi.',size=9)
table(['No','Fase','Tantangan','Risiko','a','b','a×b','Tingkat','Pengendalian dan PIC'],[
['1','Riset','Brief belum jelas','Revisi lingkup','3','2','6','Tinggi','Sahkan brief dan batas pekerjaan; ketua tim.'],
['2','Produksi','Versi file tertukar','Aset hilang/salah','2','2','4','Sedang','Penamaan versi, backup dan log perubahan; PIC aset.'],
['3','Desain','Aset atau klaim tanpa bukti','Keberatan mitra','2','3','6','Tinggi','Periksa izin aset dan bukti klaim sebelum publikasi; PIC konten.'],
['4','Test','Peserta tidak sesuai','Kesimpulan bias','2','2','4','Sedang','Kriteria rekrutmen, instrumen konsisten, catat batas studi; PIC UX.'],
['5','Pameran','Kabel di jalur orang','Tersandung','2','3','6','Tinggi','Pelindung kabel dan jalur bersih; PIC teknis.'],
['6','Pameran','Layar/internet gagal','Demo berhenti','2','2','4','Sedang','Demo offline, uji perangkat dan cadangan file; PIC teknis.'],
['7','Pameran','Instalasi/listrik tidak siap','Cedera/kerusakan','2','3','6','Tinggi','Cek pemasangan bersama teknisi venue; jangan operasikan sebelum aman; ketua booth.']
],[.6,1.5,2.7,2.3,.5,.5,.7,1.4,7.2],source=6,size=9)

heading(6,'Estimasi waktu pekerjaan',page=True)
add('Linimasa 14 minggu bersifat iteratif. Kegiatan antarmata kuliah berjalan paralel; persiapan pameran dimulai minggu 4–7. Tanggal evaluasi mengikuti panduan dan pengumuman terbaru dari pengajar.')
table(['Fase atau minggu','Uraian pekerjaan','Estimasi','Bukti dan titik pemeriksaan'],[
['Empathize\nM1–M2','Wawancara, audit identitas dan media, riset calon pelanggan, inventaris aset dan kebutuhan mitra.','2 minggu','Project brief dan data riset; masalah belum dianggap terbukti sebelum data ditelaah.'],
['Define\nM3','Analisis SWOT/kompetitor, positioning, USP, persona pengguna, kebutuhan audiens, creative brief, kriteria penerimaan.','1 minggu','Brief yang dibahas dengan mitra dan manajer proyek.'],
['Ideate dan uji awal\nM4–M5','Moodboard, minimal tiga alternatif logo, formula wara, tone of voice, sitemap, use case, wireframe; konsep dan pembagian peran pameran.','2 minggu','Alternatif beserta hasil uji awal; draf naskah mahasiswa dan daftar revisi.'],
['Prototype awal\nM6–M7','Aplikasi identitas pada media, draf guideline, prototipe awal, konsep booth, RAB, alur acara dan risiko; persiapan pitching.','2 minggu','Cakupan draf lintas mata kuliah untuk evaluasi tengah semester.'],
['Prototype lanjutan\nM8–M10','Perbaikan hasil evaluasi; desain antarmuka lengkap, prototipe interaktif, konten lintas media, motion logo dan poster.','3 minggu','Aset produksi dan prototipe siap uji; pemeriksaan kesesuaian brief.'],
['Test dan revisi\nM11–M12','Uji logo, pemahaman wara dan usability; analisis data, perbaikan dan pengujian ulang; gladi teknis pameran.','2 minggu','Laporan pengujian, revisi, technical rider, SOP dan daftar kesiapan.'],
['Finalisasi dan Operate\nM13–M14','Finalisasi guideline, laporan, HKI/BAST, portofolio, aset dan materi pameran; pelaksanaan/evaluasi pameran mengikuti jadwal resmi.','2 minggu','Paket serah terima, dokumentasi dan laporan evaluasi.'],
['Total','Durasi fase utama; pengujian dan koordinasi berlangsung berulang.','14 minggu','Evaluasi akademik mengikuti kalender kampus.']
],[3.0,7.0,2.0,5.4],source=7)
sub('Target evaluasi berdasarkan Panduan Branding Project 2026')
add('Tengah semester: progres minimal 50% dengan bukti logbook dan backlog SIAP PBL. Siapkan draf strategi, minimal tiga alternatif logo dan hasil uji, formula wara/voice/tagline, konsep media interaktif, UI, wireframe, pameran, portofolio dan risiko. Panduan mencantumkan batas unggah slide 23 Oktober 2026 serta pitching berbahasa Inggris maksimal 10 menit dan tanya jawab 10 menit.')
add('Akhir semester: progres minimal 80% untuk kelayakan evaluasi; seluruh luaran akhir tetap ditargetkan selesai. Panduan mencantumkan pameran minggu kedua UAS dan batas pengumpulan dokumen 5 Januari 2027. Ketua tim memeriksa perubahan jadwal melalui pengajar dan SIAP PBL. Angka progres evaluasi tidak menggantikan kewajiban menyelesaikan paket serah terima.')

heading(7,'Biaya proyek',page=True)
add('RAB ini adalah estimasi perencanaan internal, bukan penawaran penyedia. Perangkat milik anggota dan peminjaman kampus diasumsikan tanpa biaya tambahan. Harga, sumber dana, serta kontribusi pameran bersama ditetapkan setelah survei dan persetujuan tim/mitra.')
table(['Fase','Uraian dan perhitungan','Perkiraan biaya','Catatan'],[
['Riset','Transportasi 2 kunjungan × Rp75.000','Rp150.000','Asumsi kunjungan lokal'],
['Koordinasi','Internet tambahan 1 paket tim × Rp100.000','Rp100.000','Di luar akses yang tersedia'],
['Produksi','Cetak uji logo dan media 1 paket × Rp100.000','Rp100.000','Uji keterbacaan dan warna'],
['Pameran','Poster A2 1 lembar × Rp75.000','Rp75.000','Spesifikasi cetak dikonfirmasi'],
['Pameran','Label dan materi pendukung 1 paket × Rp125.000','Rp125.000','Termasuk informasi/QR'],
['Pameran','Bahan booth dan pengamanan kabel 1 paket × Rp300.000','Rp300.000','Sesuaikan kontribusi kelas'],
['Pelaporan','Dokumentasi dan serah terima 1 paket × Rp150.000','Rp150.000','Cetak/penyimpanan seperlunya'],
['Subtotal','Jumlah estimasi biaya langsung','Rp1.000.000','Belum termasuk biaya baru di luar lingkup'],
['Cadangan','10% × Rp1.000.000','Rp100.000','Dicatat penggunaannya'],
['Total','Subtotal + cadangan','Rp1.100.000','Estimasi, belum disahkan']
],[2.3,7.3,3.0,4.8],source=9,size=9)
heading(8,'Tim proyek')
add('Tim mahasiswa berikut sesuai daftar kelompok. Penugasan dosen/laboran serta NIK mengikuti penetapan resmi manajer proyek. Pembagian PIC di bawah merupakan usulan kerja dan dapat disesuaikan menurut kompetensi anggota.',size=9)
table(['No','Nama','NIM','Program studi dan usulan tanggung jawab'],[
['1','Ashsyura Az Zahra','4312411071','TRM; ketua tim, koordinasi dan pengendalian jadwal'],
['2','Meliska Grenova Sinaga','4312411052','TRM; riset merek dan creative brief'],
['3','Olyn Novita Dewi','4312411058','TRM; identitas visual dan guideline'],
['4','Erisa Junita','4312411064','TRM; penulisan wara dan integrasi media'],
['5','Metia Azzahra','4312432003','TRM; UX, instrumen dan analisis pengujian'],
['6','Ivan Suhendra S','4312431005','TRM; antarmuka, prototipe dan teknis booth']
],[.8,5.3,3.0,8.3],source=10,size=9)
add('Seluruh anggota berkontribusi dalam riset, produksi, presentasi, pameran dan laporan. Motion logo, poster, keuangan dan dokumentasi dibagi melalui backlog; setiap anggota menyusun portofolio pribadinya.',size=9)
heading(9,'Ruang kerja dan laboratorium')
add('Workspace kampus mengikuti persetujuan peminjaman; nomor ruang final dikonfirmasi. Pertemuan mitra dilakukan di lokasi yang disepakati atau secara daring. Acuan venue pameran dalam panduan adalah Auditorium GU lantai 2 pada minggu kedua UAS; alokasi booth dan jadwal teknis mengikuti panitia.',size=9)

heading(10,'Mata kuliah dan capaian pembelajaran yang terlibat',page=True)
add('Mata kuliah mengikuti tujuh mata kuliah Branding Project 2026. Kolom capaian memetakan pekerjaan Senja; rumusan CLO resmi digunakan untuk TRM527 dan TRM532 yang lesson plan-nya tersedia. Untuk mata kuliah lain, indikator kerja berikut perlu dicocokkan pengajar dengan CPMK resmi.')
table(['No','Mata kuliah','Capaian dalam proyek','CPMK atau indikator dan bukti'],[
['1','TRM527\nPenjenamaan\n4 SKS','Menganalisis kebutuhan identitas Senja dan mengembangkan sistem visual yang konsisten.','CLO 1–6: identifikasi kebutuhan; analisis strategi; pengembangan identitas; aplikasi ke media; publikasi panduan; teamwork. Bukti: project/creative brief, alternatif dan pengujian logo, guideline A4 landscape.'],
['2','TRM528\nPenulisan Wara\n3 SKS','Menyusun komunikasi persuasif berdasarkan karakter merek dan kebutuhan audiens.','Indikator: enam unsur wara, formula AIDA/PAS/FAB/4C yang relevan, persona/ego audiens, premis, retorika, curiosity gap, voice/tone dan CTA lintas media. Bukti: copywriting guideline dan tagline untuk media sosial serta website, draf dan revisi.'],
['3','TRM529\nMedia Digital Interaktif\n3 SKS','Mengembangkan prototipe dengan interaksi dan respons yang jelas.','Indikator: merancang konsep, narasi dan logika interaktif; membuat serta mengevaluasi prototipe. Bukti: prototipe layanan Senja dengan alur informasi menuju kontak/konsultasi.'],
['4','TRM530\nDesain Antar Muka\n3 SKS','Menerjemahkan kebutuhan pengguna menjadi antarmuka yang konsisten dengan identitas Senja.','Indikator: hierarki informasi, tata letak, navigasi, konsistensi komponen dan adaptasi layar. Bukti: dokumen desain/implementasi antarmuka website atau aplikasi sesuai luaran mata kuliah.'],
['5','TRM531\nDesain Pengalaman Pengguna\n3 SKS','Merancang pengalaman berdasarkan riset pengguna serta mengevaluasi usability.','Indikator: kebutuhan pengguna, persona, use case, wireframe, skenario uji dan analisis perbaikan. Bukti: riset pengguna, hasil pengujian usability dan riwayat revisi.'],
['6','TRM532\nTata Kelola Pameran\n2 SKS','Mengemas hasil rebranding menjadi pengalaman pameran yang terkurasi dan layak dilaksanakan.','CLO 1–6: diversifikasi pameran; konsep; tim dan alur acara; kuratorial; pelaksanaan; teamwork. Bukti: rancangan booth, manajemen kru, alur pengunjung, RAB, risiko K3L dan evaluasi.'],
['7','TRM533\nPengembangan Portofolio\n3 SKS','Mengkurasi dan menyajikan bukti kontribusi serta kompetensi individu.','Indikator: pemilihan karya, narasi proses, visualisasi, revisi dan publikasi portofolio. Bukti: portofolio tiap anggota dan QR/tautan untuk pameran.']
],[.7,3.2,5.0,8.5],source=11,size=9)
add('TRM527 merupakan mata kuliah proyek; TRM528–TRM533 merupakan mata kuliah pendukung menurut Panduan Branding Project 2026. Pelaksanaan dan penilaian tetap mengikuti pengajar masing-masing.',size=9)

heading(11,'Komunikasi antara manajer proyek dan klien',page=True)
add('Tabel ini merupakan agenda komunikasi yang harus dijawab melalui pertemuan resmi. Keputusan, tanggal, penanggung jawab dan tautan bukti dicatat setelah pertemuan; tidak dianggap sebagai persetujuan sebelum dikonfirmasi.')
table(['Fase','Pertanyaan atau keputusan','Jawaban atau status','Tindak lanjut'],[
['Empathize','Siapa PIC Senja, apa layanan utama, target pelanggan dan masalah identitas saat ini?','Menunggu wawancara mitra.','Ketua menjadwalkan pertemuan; riset mencatat bukti.'],
['Define','Nilai merek, positioning, ruang lingkup, aset yang boleh dipakai dan kriteria keberhasilan?','Menunggu validasi project/creative brief.','Ketua dan manajer proyek mengesahkan acuan kerja.'],
['Ideate','Alternatif logo dan arah komunikasi mana yang paling sesuai?','Diputuskan setelah presentasi alternatif dan umpan balik.','Catat alasan pemilihan dan revisi.'],
['Prototype','Apakah informasi layanan, konten dan alur kontak sudah sesuai?','Menunggu reviu prototipe.','PIC konten/UI menindaklanjuti daftar masalah.'],
['Test/serah terima','Apakah luaran memenuhi kriteria; apa aset dan hak penggunaan yang diserahkan?','Menunggu hasil uji dan pemeriksaan akhir.','Manajer proyek/PIC mitra memeriksa daftar aset dan BAST.']
],[2.5,6.2,4.3,4.4],source=13)
heading(12,'Monitoring dan evaluasi')
add('Pertemuan tim dan manajer proyek direncanakan minimal sekali setiap minggu. Ketua memelihara backlog berisi pekerjaan, PIC, tenggat, status, bukti dan kendala; setiap anggota memperbarui logbook SIAP PBL. Perubahan lingkup, biaya atau tenggat harus disertai dampak dan keputusan tertulis.')
table(['Objek evaluasi','Kriteria pemeriksaan','Bukti yang disimpan'],[
['Identitas dan wara','Konsistensi terhadap brief, keterbacaan, pemahaman pesan, kejelasan CTA dan bukti klaim.','Instrumen, hasil pengujian, komentar mitra dan revisi.'],
['Prototipe/UI/UX','Tugas inti dapat diselesaikan; navigasi, respons dan konten dipahami.','Keberhasilan tugas, kesalahan, waktu dan perbaikan.'],
['Pameran','Karya terkurasi, alur jelas, perangkat berfungsi, kru siap, risiko ditangani.','Denah, technical rider, uji teknis, absensi, dokumentasi, survei dan laporan biaya.'],
['Kontribusi dan serah terima','Kontribusi individu dapat ditelusuri; seluruh aset dan dokumen diperiksa.','Logbook, backlog, portofolio, daftar aset dan BAST.']
],[3.1,7.0,7.3],source=7,size=9)
add('Penilaian mengikuti panduan: kontribusi manajer proyek 30% dan pengajar mata kuliah 70%, dengan rubrik masing-masing. Evaluasi pameran membandingkan tujuan, realisasi kegiatan/biaya, umpan balik pengunjung, kendala serta tindakan perbaikan. Hasil pengujian dilaporkan apa adanya, termasuk keterbatasannya.')

heading(13,'Riwayat perubahan proyek',page=True)
table(['Revisi dan tanggal','Deskripsi perubahan','Pengusul atau penanggung jawab'],[
['01 / 2 Oktober 2026','Penyesuaian contoh E-Koperasi menjadi Rebranding Visual Senja Wedding Organizer; penyelarasan tujuh mata kuliah, luaran, jadwal, tim, risiko dan evaluasi dengan materi 2026.','Tim PBL Senja; menunggu reviu manajer proyek'],
['Revisi berikutnya','Diisi setelah terdapat keputusan perubahan lingkup, rancangan, jadwal atau biaya.','Dicatat oleh ketua tim']
],[3.3,9.0,5.1],source=14)
sub('Tanda tangan persetujuan')
add('Batam, ........ / ........ / 2026')
add('Persetujuan diberikan setelah isi RPP diverifikasi oleh pihak berwenang. Nama jabatan dan penandatangan mengikuti penetapan resmi. Ruang di bawah belum merupakan pengesahan.')
table(['Manajer proyek','PIC Senja Wedding Organizer'],[
['\n\n\nNama: ........................................\nNIK: ............................................','\n\n\nNama: ........................................\nJabatan: .......................................']
],[8.7,8.7],source=15)
table(['P3M','SHILAU','Pihak akademik terkait'],[
['\n\n\nNama: ..............................\nNIK: ................................','\n\n\nNama: ..............................\nNIK: ................................','\n\n\nNama: ..............................\nNIK: ................................']
],[5.8,5.8,5.8],source=16)
sub('Hal yang diselesaikan sebelum pengesahan')
add('Konfirmasi kode PBL, nama resmi merek, kelas, pengusul/manajer proyek dan PIC mitra; validasi temuan riset serta creative brief; sepakati PIC anggota, fitur prototipe, RAB dan sumber dana; cocokkan CPMK dengan lesson plan pengajar; tetapkan jadwal serta alokasi booth. Lampirkan project brief dan borang risiko yang telah divalidasi.')

add('Lampiran acuan materi perkuliahan',bold=True,size=11,page=True)
add('Acuan berikut menjadi dasar perencanaan dan memudahkan penelusuran hubungan antara materi, pekerjaan dan luaran. Nomor halaman merujuk urutan halaman PDF, termasuk sampul.')
table(['Acuan lokal','Pokok yang diterapkan','Bagian RPP'],[
['Panduan Branding Project 2026.pdf, hlm. 3–6','Tujuh mata kuliah, peran manajer, spesifikasi luaran, dokumen pendukung, HKI eksternal dan BAST.','Identitas; 1; 3; 10'],
['Panduan Branding Project 2026.pdf, hlm. 7–19','Linimasa, evaluasi tengah/akhir, logbook, backlog, tenggat dan penilaian.','6; 11; 12'],
['Pertemuan2_Penjenamaan Visual.pdf, hlm. 20–21; Pertemuan3_Creative Brief.pdf, hlm. 2–7','Analisis brand value, positioning, pasar sasaran, tujuan komunikasi dan arahan kreatif.','1; 2; 3'],
['Pertemuan4_Desain Identitas Visual.pdf, hlm. 14–27','Prinsip logo, eksplorasi minimal tiga alternatif, validasi, file akhir dan guideline.','2; 3; 6; 12'],
['Lesson Plan TRM527 Penjenamaan 2026, bagian III, V dan VII','Tujuan dan CLO Penjenamaan serta tahap CDIO.','3; 6; 10'],
['pengantar_penulisan wara.txt; tipe_dan_pola_penulisan_wara.txt','Unsur wara, tangible/intangible, AIDA, PAS, FAB, 4C serta pola storytelling.','2; 3; 10'],
['Persona & Target Audiens.pdf, hlm. 2–17','Persona target dan gaya penulis; kebutuhan, keinginan, citra diri serta ego audiens.','2; 3; 10'],
['Premis Corius.pdf, hlm. 3–19; INTEGRASI MEDIA DIGITA.pdf, hlm. 3–19','Premis, retorika, curiosity gap, pesan lintas media, voice, tone dan CTA.','2; 3; 10'],
['Template penulisan Integrasi Media.pdf, hlm. 1','Tugas individu mengisi premis, persona, ego, curiosity gap dan integrasi media. Kalimat wara dikerjakan mahasiswa sesuai ketentuan tugas.','2; 10'],
['Lesson Plan TRM532 Tata Kelola Pameran 2026, hlm. 1–3; tugas minggu 4–7','CLO pameran, peran kru, sepuluh komponen perencanaan, absensi dan evaluasi tim.','2; 4–6; 10; 12'],
['Kuratorial_Karya_Multimedia.pdf, hlm. 2–6 dan 9; Manajemen_Produksi_Pameran_Multimedia.pdf, hlm. 3–5 dan 8','Seleksi/narasi/label, alur pengunjung, kelayakan teknis, RAB, jadwal, risiko, SOP serta evaluasi.','2–6; 12']
],[6.5,8.1,2.8],source=7,size=9)
add('Lampiran pelaksanaan yang dikembangkan selama proyek: project brief, creative brief, catatan riset, hasil pengujian logo/wara/usability, guideline, prototipe, denah dan technical rider, Gantt, RAB final, borang risiko No.FO.17.1.1-V0, dokumentasi pameran, portofolio, draf HKI dan BAST.',size=9)
add('Lampiran asumsi kerja khusus Senja',bold=True,size=11,page=True)
add('Asumsi berikut merupakan arah perancangan yang akan diuji pada fase Empathize dan Define. Asumsi tidak menjadi klaim tentang kondisi bisnis atau hasil penelitian Senja. Jika temuan lapangan berbeda, tim memperbarui brief dan riwayat perubahan.')
table(['Aspek','Asumsi dan keputusan rancangan','Cara memeriksa'],[
['Segmen utama','Pasangan bekerja usia 24–32 tahun di Batam, rencana menikah 6–12 bulan lagi, waktu persiapan terbatas, membandingkan beberapa WO sebelum menghubungi.','Tanyakan jadwal persiapan, pihak pengambil keputusan, sumber informasi, dan alasan memilih WO kepada calon pelanggan.'],
['Persona kerja','Persona ilustratif: Rani, 27 tahun, karyawan di Batam, menyiapkan pernikahan bersama pasangan. Ia ingin pembagian tugas jelas dan waktu bersama keluarga, tetapi belum paham batas pekerjaan WO dibanding vendor.','Persona ini bukan responden nyata. Cocokkan kebutuhan dan hambatannya dengan wawancara pengguna.'],
['Layanan yang diprioritaskan','Asumsi fokus Senja: koordinasi persiapan dan pelaksanaan hari pernikahan. Prototipe menjelaskan alur diskusi kebutuhan, penentuan cakupan, koordinasi dan pelaksanaan.','Minta daftar layanan, tugas yang termasuk/tidak termasuk, proses konsultasi dan wilayah layanan dari mitra. Jangan menyamakan WO dengan paket dekorasi/katering.'],
['Positioning awal','Pendamping koordinasi pernikahan yang hangat dan terorganisasi, untuk pasangan yang ingin memahami proses serta pembagian tanggung jawab sejak awal.','Bandingkan janji tersebut dengan praktik layanan dan positioning kompetitor lokal. USP final harus didukung bukti.'],
['Arah visual','Elegan dan hangat dengan eksplorasi warna terakota, krem dan cokelat gelap; serif yang terbaca untuk aksen dan sans serif untuk isi. Eksplorasi wordmark, monogram S, dan kombinasi simbol tanpa ornamen berlebihan.','Uji tiga alternatif terhadap keterbacaan, kemiripan, asosiasi dan preferensi audiens. Warna final harus lolos uji kontras dan cetak.'],
['Konten prioritas','Informasi layanan/cakupan, alur kerja, portofolio yang berizin, pertanyaan umum dan kontak resmi. Voice hangat; tone informatif ketika menjelaskan layanan.','Minta peserta menjelaskan apa yang ditangani Senja dan apa yang harus ditanyakan sebelum memilih layanan.'],
['Interaksi utama','Prototipe memungkinkan pengguna menjelajah layanan dan portofolio, kemudian membuka simulasi konsultasi. Input yang diusulkan: nama, tanggal rencana, lokasi dan kebutuhan.','Uji tugas tanpa bantuan. Data yang dimasukkan pada demo adalah data contoh; pengiriman ke kanal nyata di luar lingkup prototipe.'],
['Bukti kepercayaan','Tampilkan portofolio, penjelasan proses dan identitas kontak yang diizinkan mitra. Testimoni hanya digunakan jika autentik serta mendapat izin.','Periksa sumber dan izin setiap foto, ulasan, angka pengalaman, harga dan klaim layanan sebelum dipublikasikan.']
],[2.7,9.0,5.7],source=7,size=9)
sub('Skenario uji yang sesuai layanan')
add('Skenario 1: Anda akan menikah enam bulan lagi di Batam. Temukan layanan yang relevan dan jelaskan apa yang ditangani Senja. Skenario 2: Cari contoh acara yang membantu Anda menilai kecocokan layanan. Skenario 3: Temukan cara berkonsultasi dan sebutkan informasi yang perlu Anda siapkan. Tim mencatat keberhasilan, titik bingung, waktu, dan pertanyaan yang belum terjawab.',size=9)
# Reconcile the located project brief with the course-wide academic requirements.
replacements={
'Kode PBL resmi menunggu konfirmasi manajer proyek':'TRM528 (mengikuti project brief; validasi kode pada SIAP PBL)',
'Senja Wedding Organizer; nama pemilik/PIC dan kontak resmi dikonfirmasi':'Senja Event & Wedding Planner Batam (CV. Multi Art Project, menurut project brief); PIC dan kontak resmi dikonfirmasi',
'prototipe website layanan':'prototipe akademik katalog layanan digital',
'Pengembangan sistem transaksi, pembayaran, dan pemesanan operasional tidak termasuk cakupan awal.':'Sesuai project brief, pembuatan website operasional, iklan berbayar dan produksi video event tidak termasuk lingkup mitra. Prototipe UI/UX dan motion logo tetap direncanakan untuk memenuhi luaran akademik; batas penyerahannya disepakati bersama pengajar dan mitra. Sistem transaksi, pembayaran dan pemesanan operasional tidak termasuk cakupan awal.',
'Arah awal yang diajukan adalah identitas yang hangat, tenang, dan profesional.':'Mengikuti creative brief Senja, arah identitas adalah romantis, hangat, rapi, elegan dan timeless dengan suasana golden hour.',
'Kesan tersebut akan diuji terhadap nilai merek dan preferensi audiens; keputusan logo, warna, tipografi, serta tagline ditetapkan melalui creative brief dan persetujuan mitra.':'Logo matahari terbenam SENJA dipertahankan atau disempurnakan; palet emas, cokelat hangat, krem dan aksen gelap dikembangkan bersama serif elegan dan sans serif yang terbaca. Tagline dari brief adalah “Timeless Moments, Perfectly Planned”. Validasi mitra tetap dilakukan sebelum finalisasi.',
'Asumsi: pasangan usia 24–32 tahun di Batam, bekerja, merencanakan pernikahan dalam 6–12 bulan; keluarga ikut memberi pertimbangan.':'Sesuai brief: calon pengantin 22–35 tahun di Batam–Kepri, kelas menengah dan aktif di Instagram; keluarga ikut mengambil keputusan. Asumsi skenario: pasangan bekerja, rencana menikah 6–12 bulan lagi.',
'Ketua dan manajer proyek mengesahkan acuan kerja.':'Ketua, manajer proyek dan PIC mitra memvalidasi acuan kerja.',
'TRM; riset merek dan creative brief':'TRM; riset merek dan copywriter',
'TRM; penulisan wara dan integrasi media':'TRM; desain grafis dan layout',
'TRM; UX, instrumen dan analisis pengujian':'TRM; media sosial dan konten',
'TRM; antarmuka, prototipe dan teknis booth':'TRM; presentasi dan dokumentasi',
'Seluruh anggota berkontribusi dalam riset, produksi, presentasi, pameran dan laporan. Motion logo, poster, keuangan dan dokumentasi dibagi melalui backlog; setiap anggota menyusun portofolio pribadinya.':'Peran dasar mengikuti usulan pada project brief. Usulan tugas tambahan lintas mata kuliah: Ivan pada prototipe/teknis, Metia pada koordinasi uji pengguna, serta Olyn dan Erisa pada motion/logo/poster. Seluruh anggota ikut pengujian, laporan dan pameran; setiap anggota menyusun portofolio pribadinya.',
'Pasangan bekerja usia 24–32 tahun di Batam, rencana menikah 6–12 bulan lagi, waktu persiapan terbatas, membandingkan beberapa WO sebelum menghubungi.':'Brief menetapkan calon pengantin 22–35 tahun di Batam–Kepri, kelas menengah, aktif di Instagram. Asumsi kerja yang diuji: pasangan bekerja, rencana menikah 6–12 bulan lagi dan membandingkan beberapa WO.',
'Elegan dan hangat dengan eksplorasi warna terakota, krem dan cokelat gelap; serif yang terbaca untuk aksen dan sans serif untuk isi. Eksplorasi wordmark, monogram S, dan kombinasi simbol tanpa ornamen berlebihan.':'Sesuai brief: emas, cokelat hangat, krem, aksen gelap; serif elegan dan sans serif bersih. Tiga alternatif merupakan penyempurnaan simbol matahari terbenam dan logotype SENJA, bukan penggantian identitas tanpa dasar.',
'Informasi layanan/cakupan, alur kerja, portofolio yang berizin, pertanyaan umum dan kontak resmi. Voice hangat; tone informatif ketika menjelaskan layanan.':'Prioritas brief: template feed/story, highlight cover, katalog paket, konten edukasi, proses kerja kru dan CTA WA/DM yang jelas. Voice hangat, meyakinkan dan personal; tone informatif pada layanan.',
'Asumsi berikut merupakan arah perancangan yang akan diuji pada fase Empathize dan Define. Asumsi tidak menjadi klaim tentang kondisi bisnis atau hasil penelitian Senja. Jika temuan lapangan berbeda, tim memperbarui brief dan riwayat perubahan.':'Arah yang sudah tercantum dalam creative/project brief dibedakan dari asumsi tambahan untuk persona dan skenario pengujian. Data jumlah pengikut, engagement, konversi, harga dan hasil layanan harus diaudit sebelum menjadi dasar klaim. Jika temuan lapangan berbeda, tim memperbarui brief dan riwayat perubahan.'
}
for node in body.xpath('.//w:t',namespaces=NS):
    if node.text:
        for a,b in replacements.items():node.text=node.text.replace(a,b)
add('Penyelarasan brief dengan rencana semester: brief memuat timeline usulan enam minggu dan tahun akademik 2025/2026. RPP menggunakan 14 minggu serta tahun 2026/2027 sesuai panduan kuliah yang berlaku pada berkas semester ini. Pekerjaan enam minggu diperlakukan sebagai paket awal pengembangan identitas, dilanjutkan integrasi luaran, pengujian dan pameran. Cakupan sekunder event korporat/komunitas tidak menjadi fokus uji utama calon pengantin.',size=9)
add('Acuan proyek tambahan: Creative-Brief-dan-Project-Brief-Senja-WO.docx, bagian Project Brief dan Creative Brief. Katalog paket tidak memuat harga rekaan; harga hanya dicantumkan setelah dikonfirmasi mitra. Uji prototipe memakai salinan akademik, bukan website yang dipublikasikan.',size=9)
body.append(sect)
# Normalize property order for strict Word parsing.
orders={
'pPr':['pStyle','keepNext','keepLines','pageBreakBefore','widowControl','numPr','pBdr','shd','tabs','spacing','ind','contextualSpacing','jc','rPr','sectPr'],
'tblPr':['tblStyle','tblpPr','tblOverlap','bidiVisual','tblStyleRowBandSize','tblStyleColBandSize','tblW','jc','tblCellSpacing','tblInd','tblBorders','shd','tblLayout','tblCellMar','tblLook','tblCaption','tblDescription'],
}
for tag,order in orders.items():
    for pr in body.findall('.//w:'+tag,NS):
        items=list(pr)
        for item in items:pr.remove(item)
        for item in sorted(items,key=lambda n:order.index(E.QName(n).localname) if E.QName(n).localname in order else 99):pr.append(item)
OUT.parent.mkdir(exist_ok=True)
parts['word/document.xml']=E.tostring(doc,xml_declaration=True,encoding='UTF-8',standalone=True)
with ZipFile(OUT,'w',ZIP_DEFLATED) as zz:
    for n,b in parts.items():zz.writestr(n,b)
with ZipFile(OUT) as zz:
    assert all(zz.read(n)==z.read(n) for n in parts if n!='word/document.xml')
text=' '.join(doc.xpath('//w:t/text()',namespaces=NS))
assert 'koperasi online' not in text.lower()
assert sum([2,1,2,2,3,2,2])==14
print(OUT)
print('Preserved all package parts except document.xml; source SHA256',hashlib.sha256(REF.read_bytes()).hexdigest())
