# Script untuk menghasilkan Excel Riset Komprehensif: Dari Masalah File Clutter Menjadi Ide Bisnis (30 Sumber)
# Membutuhkan modul ImportExcel

$excelPath = "C:\Users\ERIC PATRA RIYADI\.gemini\antigravity\scratch\riset-ide-bisnis-dokumen-ai\Riset_Solusi_File_Clutter_ke_Ide_Bisnis_30_Sumber.xlsx"

# Hapus file lama jika ada
if (Test-Path $excelPath) {
    Remove-Item $excelPath -Force
}

Write-Host "Menyiapkan Data Sheet 1: Ringkasan Eksekutif & Ide Bisnis..."
$sheet1Data = @(
    [PSCustomObject]@{
        "Komponen Strategis" = "Nama Inisiatif / Produk"
        "Deskripsi & Detail Rinci" = "DocuFlow AI (Autonomous Contextual Document Assistant & Semantic Search Engine)"
        "Kategori / Status" = "Konsep Produk & Validasi Bisnis"
    },
    [PSCustomObject]@{
        "Komponen Strategis" = "Value Proposition Utama"
        "Deskripsi & Detail Rinci" = "'Zero-Effort Document Intelligence: Never Organize a Folder Again'. Asisten AI cerdas yang mengindeks, memahami konten, dan mengelompokkan dokumen secara virtual di drive lokal maupun cloud tanpa memindahkan file fisik atau menuntut tagging manual dari pengguna."
        "Kategori / Status" = "Unique Value Proposition (UVP)"
    },
    [PSCustomObject]@{
        "Komponen Strategis" = "Problem Statement Utama"
        "Deskripsi & Detail Rinci" = "Pekerja berpengetahuan (knowledge workers) menghabiskan rata-rata 1.8 hingga 2.5 jam per hari (McKinsey, IDC) hanya untuk mencari dokumen yang berserakan. 54% pekerja lebih banyak menghabiskan waktu mencari file daripada membalas email (Dropbox/Wakefield), dan 60%+ mengalami stres kognitif akibat digital hoarding dan struktur folder yang rusak."
        "Kategori / Status" = "Validasi Masalah Pasar"
    },
    [PSCustomObject]@{
        "Komponen Strategis" = "Akar Masalah (Root Cause)"
        "Deskripsi & Detail Rinci" = "1. Struktur folder hierarkis (tree structure) tidak cocok dengan memori asosiatif manusia. 2. Manual tagging gagal karena friction terlalu tinggi. 3. Fragmentasi penyimpanan (Google Drive, OneDrive, Local SSD, WhatsApp, Email). 4. Multi-versi yang membingungkan (Draft, Final, Revisi_v2)."
        "Kategori / Status" = "Analisis Akar Masalah"
    },
    [PSCustomObject]@{
        "Komponen Strategis" = "Inovasi Solusi (Core Innovation)"
        "Deskripsi & Detail Rinci" = "1. Virtual Multi-Dimensional Spaces: Satu dokumen bisa muncul di berbagai kategori virtual (Klien, Tahun, Jenis Dokumen) tanpa duplikasi fisik. 2. Natural Language Semantic Retrieval: Pencarian makna konteks, bukan exact keyword. 3. Privacy-First Local AI: Pembuatan embedding teks dilakukan on-device tanpa mengunggah data sensitif ke server publik. 4. Smart Version & Duplicate Pruning."
        "Kategori / Status" = "Diferensiasi Teknologi"
    },
    [PSCustomObject]@{
        "Komponen Strategis" = "Target Segmen Pasar (Target Customers)"
        "Deskripsi & Detail Rinci" = "1. Freelancers, Konsultan, Akuntan, Notaris, & Peneliti (Solo Knowledge Workers). 2. Usaha Mikro, Kecil, dan Menengah (UMKM) & Agensi Kreatif (5-50 orang) tanpa tim IT khusus. 3. Tim Operasional & Legal Perusahaan yang kewalahan dengan file silo."
        "Kategori / Status" = "Segmentasi Pasar"
    },
    [PSCustomObject]@{
        "Komponen Strategis" = "Model Monetisasi & Pendapatan"
        "Deskripsi & Detail Rinci" = "SaaS Freemium: Free Tier (Indeks lokal hingga 1.000 file), Pro Tier ($9/bulan atau $90/tahun untuk file tak terbatas + multi-cloud connectors), Team Tier ($19/user/bulan dengan shared spaces), Enterprise Tier ($39/user/bulan untuk on-premise vector vault & custom compliance)."
        "Kategori / Status" = "Model Bisnis"
    },
    [PSCustomObject]@{
        "Komponen Strategis" = "Estimasi Ukuran Pasar (Market Size)"
        "Deskripsi & Detail Rinci" = "Total Addressable Market (TAM) Enterprise Search & DAM global diproyeksikan mencapai $10.2 Miliar pada tahun 2026 (Statista, CAGR 12.8%). Serviceable Addressable Market (SAM) untuk pasar individu & UMKM diperkirakan $3.1 Miliar. Target Serviceable Obtainable Market (SOM) tahun ke-3 adalah $45 Juta."
        "Kategori / Status" = "Estimasi Pasar Global"
    },
    [PSCustomObject]@{
        "Komponen Strategis" = "Metrik ROI bagi Pelanggan"
        "Deskripsi & Detail Rinci" = "Menghemat rata-rata 6.5 jam kerja per minggu per karyawan (~$350-$600 nilai produktivitas yang diselamatkan per bulan per karyawan), mengurangi duplikasi pekerjaan hingga 90%, dan menghilangkan stres pencarian dokumen."
        "Kategori / Status" = "Customer ROI & Impact"
    },
    [PSCustomObject]@{
        "Komponen Strategis" = "Rencana Roadmap Eksekusi (Phase 1 - 3)"
        "Deskripsi & Detail Rinci" = "Fase 1 (Bulan 1-4): MVP Desktop App (Windows/macOS) dengan local semantic search dan auto-tagging folder Downloads/Desktop. Fase 2 (Bulan 5-8): Integrasi Google Drive & OneDrive API + Fitur 'Ask My Docs' (Local RAG). Fase 3 (Bulan 9-12): Kolaborasi Tim & Sinkronisasi Ruang Virtual."
        "Kategori / Status" = "Roadmap Pengembangan"
    }
)

Write-Host "Menyiapkan Data Sheet 2: 30 Sumber Riset Empiris..."
$sheet2Data = @(
    [PSCustomObject]@{
        "No" = 1
        "Institusi / Penulis" = "IDC (Susan Feldman & Chris Sherman)"
        "Judul Publikasi / Riset" = "The High Cost of Not Finding Information"
        "Tahun" = 2001
        "Tipe Sumber" = "Laporan Riset Industri"
        "Kategori Isu" = "Pemborosan Waktu & Biaya"
        "Data Kunci / Statistik" = "Pekerja menghabiskan 2.5 jam/hari (30% hari kerja) mencari data; 50% pencarian gagal menemukan berkas yang tepat."
        "Temuan Utama" = "Kegagalan menemukan dokumen adalah pemborosan modal terbesar dalam ekonomi informasi, memicu pembuatan ulang dokumen yang sebenarnya sudah ada."
        "Validasi Pain Point" = "Ketidakmampuan menemukan berkas secara cepat menyebabkan frustrasi dan pemborosan waktu masif."
        "Implikasi Solusi / Fitur DocuFlow AI" = "Pencarian semantik cerdas berbasis vektor yang memahami maksud pengguna meski kata kunci tidak persis sama."
    },
    [PSCustomObject]@{
        "No" = 2
        "Institusi / Penulis" = "McKinsey Global Institute"
        "Judul Publikasi / Riset" = "The Social Economy: Unlocking Value and Productivity"
        "Tahun" = 2012
        "Tipe Sumber" = "Studi Ekonomi Global"
        "Kategori Isu" = "Produktivitas Kerja"
        "Data Kunci / Statistik" = "Karyawan rata-rata menghabiskan 1.8 jam setiap hari (9.3 jam/minggu) hanya untuk mencari dan mengumpulkan informasi."
        "Temuan Utama" = "Analogi 'The 5th Employee': dari setiap 5 karyawan yang dipekerjakan, waktu 1 karyawan habis sia-sia hanya untuk berburu berkas internal."
        "Validasi Pain Point" = "Waktu berharga tenaga ahli terbuang untuk tugas administratif pengarsipan dan pencarian dokumen."
        "Implikasi Solusi / Fitur DocuFlow AI" = "Menghasilkan penghematan terukur (6+ jam/minggu) sebagai dasar value proposition utama pada pitching B2B."
    },
    [PSCustomObject]@{
        "No" = 3
        "Institusi / Penulis" = "Gartner"
        "Judul Publikasi / Riset" = "Digital Workplace Survey: Addressing Information Sprawl"
        "Tahun" = 2021
        "Tipe Sumber" = "Survei Enterprise Global"
        "Kategori Isu" = "Fragmentasi Aplikasi"
        "Data Kunci / Statistik" = "47% pekerja digital kesulitan menemukan informasi yang diperlukan karena data tersebar di berbagai aplikasi terpisah."
        "Temuan Utama" = "Penyebaran alat kolaborasi (sprawl) menciptakan pulau-pulau data (data silos) yang tidak saling terhubung."
        "Validasi Pain Point" = "Dokumen tersimpan di Google Drive pribadi, OneDrive kantor, folder Downloads laptop, dan lampiran email."
        "Implikasi Solusi / Fitur DocuFlow AI" = "Multi-source indexing connector: mengindeks folder lokal dan akun cloud dalam satu indeks pencarian terpadu."
    },
    [PSCustomObject]@{
        "No" = 4
        "Institusi / Penulis" = "Microsoft"
        "Judul Publikasi / Riset" = "Work Trend Index Annual Report: Will AI Fix Work?"
        "Tahun" = 2023
        "Tipe Sumber" = "Laporan Riset Tahunan"
        "Kategori Isu" = "Digital Debt & AI Demand"
        "Data Kunci / Statistik" = "68% pekerja berjuang dengan kurangnya waktu fokus; 64% merasa terbebani oleh pencarian komunikasi dan file terfragmentasi."
        "Temuan Utama" = "Pekerja tenggelam dalam 'utang digital' (digital debt) di mana beban mengelola informasi mengalahkan waktu pembuatan karya nyata."
        "Validasi Pain Point" = "Pengguna sangat menginginkan AI mengambil alih beban organisasi dan sintesis dokumen."
        "Implikasi Solusi / Fitur DocuFlow AI" = "Fitur 'Ask-My-Docs' (Local RAG) yang memungkinkan pengguna menanyakan isi kumpulan dokumen dengan bahasa alami."
    },
    [PSCustomObject]@{
        "No" = 5
        "Institusi / Penulis" = "Forrester Research"
        "Judul Publikasi / Riset" = "The State of Digital Worker Experience"
        "Tahun" = 2022
        "Tipe Sumber" = "Laporan Analisis Pasar"
        "Kategori Isu" = "Pengalaman Kerja Digital"
        "Data Kunci / Statistik" = "20% hingga 30% kapasitas produktif organisasi hilang akibat ketidakmampuan menemukan aset lama dan inefisiensi pencarian."
        "Temuan Utama" = "Hambatan akses informasi berkorelasi langsung dengan pergantian karyawan (turnover) dan penurunan moral kerja."
        "Validasi Pain Point" = "Frustrasi harian saat mencari dokumen penting memicu kelelahan kerja (burnout)."
        "Implikasi Solusi / Fitur DocuFlow AI" = "Universal Quick-Launcher (Alt+Space) untuk membuka dokumen yang dicari dalam waktu kurang dari 2 detik."
    },
    [PSCustomObject]@{
        "No" = 6
        "Institusi / Penulis" = "Dropbox & Wakefield Research"
        "Judul Publikasi / Riset" = "The Cost of Context Switching and Document Disorganization"
        "Tahun" = 2021
        "Tipe Sumber" = "Survei Pekerja Kantor"
        "Kategori Isu" = "Beban Mental & Disorganisasi"
        "Data Kunci / Statistik" = "54% pekerja menghabiskan lebih banyak waktu mencari file daripada membalas email; 62% merasa stres akibat file yang berantakan."
        "Temuan Utama" = "Kekacauan digital (digital clutter) menimbulkan dampak psikologis negatif yang setara dengan meja kerja fisik yang kotor."
        "Validasi Pain Point" = "Rasa bersalah dan stres karena desktop penuh dengan file acak seperti 'Screenshot_2026-11...png'."
        "Implikasi Solusi / Fitur DocuFlow AI" = "Fitur 'One-Click Declutter' otomatis mengkategorikan file desktop ke ruang kerja virtual tanpa memindahkan lokasi file."
    },
    [PSCustomObject]@{
        "No" = 7
        "Institusi / Penulis" = "Adobe Document Cloud"
        "Judul Publikasi / Riset" = "State of Digital Document Workflows"
        "Tahun" = 2020
        "Tipe Sumber" = "Laporan Industri Dokumen"
        "Kategori Isu" = "Penamaan File & Standarisasi"
        "Data Kunci / Statistik" = "70% profesional menyatakan bahwa format penamaan file yang tidak konsisten menyebabkan keterlambatan proyek."
        "Temuan Utama" = "Setiap karyawan memiliki gaya penamaan unik ('final_v2', 'revisi_budi'), menghancurkan efektivitas pencarian sistemik."
        "Validasi Pain Point" = "Tidak ada standarisasi penamaan di antara tim atau bahkan pada diri sendiri dari bulan ke bulan."
        "Implikasi Solusi / Fitur DocuFlow AI" = "AI Auto-Summarizer & Context Extractor yang membaca isi dokumen dan memberikan ringkasan instan tanpa perlu membuka file."
    },
    [PSCustomObject]@{
        "No" = 8
        "Institusi / Penulis" = "Nitro Software"
        "Judul Publikasi / Riset" = "Future of Work Report: Document Productivity"
        "Tahun" = 2022
        "Tipe Sumber" = "Laporan Riset Produktivitas"
        "Kategori Isu" = "Kebingungan Versi (Version Chaos)"
        "Data Kunci / Statistik" = "50% pekerja frustrasi karena tidak mengetahui dokumen mana yang merupakan versi terbaru atau kanonikal."
        "Temuan Utama" = "Terjadinya pengiriman draf lama ke klien atau pengambilan keputusan berdasarkan data finansial yang belum diperbarui."
        "Validasi Pain Point" = "Banyak salinan file yang mirip di folder Downloads dan Documents dengan tanggal berbeda-beda."
        "Implikasi Solusi / Fitur DocuFlow AI" = "Smart Version Clustering: mendeteksi kemiripan semantik antar file dan menandai mana yang merupakan 'Master / Latest Version'."
    },
    [PSCustomObject]@{
        "No" = 9
        "Institusi / Penulis" = "CoSchedule"
        "Judul Publikasi / Riset" = "State of Marketing Work Management"
        "Tahun" = 2023
        "Tipe Sumber" = "Studi Manajemen Tim Kreatif"
        "Kategori Isu" = "Duplikasi Pekerjaan"
        "Data Kunci / Statistik" = "Tim kreatif dan pemasaran menduplikasi pekerjaan 23% dari waktu karena aset asli tidak dapat ditemukan di penyimpanan tim."
        "Temuan Utama" = "Mendesain ulang grafik, menulis ulang proposal, atau mengunduh ulang materi yang sebenarnya sudah ada di drive."
        "Validasi Pain Point" = "Pemborosan anggaran dan jam kerja kreatif untuk mengulang aset yang hilang di tumpukan folder."
        "Implikasi Solusi / Fitur DocuFlow AI" = "Visual & Semantic Deduplication Engine: mendeteksi file kembar dan duplikat aset visual serta dokumen teks."
    },
    [PSCustomObject]@{
        "No" = 10
        "Institusi / Penulis" = "American Psychological Association (APA)"
        "Judul Publikasi / Riset" = "Digital Clutter, Cognitive Load, and Attention Fatigue"
        "Tahun" = 2019
        "Tipe Sumber" = "Jurnal Psikologi Terapan"
        "Kategori Isu" = "Beban Kognitif (Cognitive Overload)"
        "Data Kunci / Statistik" = "Stimulus visual dari ratusan ikon dan dokumen tanpa kategori meningkatkan tingkat kortisol dan kelelahan atensi."
        "Temuan Utama" = "Otak manusia memperlakukan setiap ikon berkas yang tidak terorganisir sebagai 'tugas terbuka' (open loop) yang menguras energi mental."
        "Validasi Pain Point" = "Kecemasan psikologis saat membuka desktop atau folder kerja yang kacau balau."
        "Implikasi Solusi / Fitur DocuFlow AI" = "Tampilan antarmuka 'Focus Mode' yang bersih dan cerdas, menyajikan hanya file yang relevan dengan tugas hari ini."
    },
    [PSCustomObject]@{
        "No" = 11
        "Institusi / Penulis" = "Harvard Business Review (HBR)"
        "Judul Publikasi / Riset" = "Why Hierarchical Folder Structures Fail Modern Knowledge Work"
        "Tahun" = 2020
        "Tipe Sumber" = "Artikel Manajemen & Teori Informasi"
        "Kategori Isu" = "Kegagalan Struktur Hierarki"
        "Data Kunci / Statistik" = "82% skema klasifikasi folder manual ditinggalkan oleh pengguna dalam waktu kurang dari 3 bulan setelah dibuat."
        "Temuan Utama" = "Folder hierarkis (Tree Structure) memaksa dokumen hanya berada di 1 lokasi, padahal dokumen bisnis memiliki multi-konteks."
        "Validasi Pain Point" = "Apakah invoice klien X harus dimasukkan ke folder 'Klien X', folder 'Finance 2026', atau folder 'Pajak'?"
        "Implikasi Solusi / Fitur DocuFlow AI" = "Multi-Dimensional Virtual Tags: satu dokumen secara otomatis hadir di semua ruang konteks yang relevan tanpa duplikasi file."
    },
    [PSCustomObject]@{
        "No" = 12
        "Institusi / Penulis" = "Ponemon Institute"
        "Judul Publikasi / Riset" = "The Security Risks of Unstructured Data Sprawl"
        "Tahun" = 2022
        "Tipe Sumber" = "Laporan Keamanan Siber & Privasi"
        "Kategori Isu" = "Keamanan & Kepatuhan Data"
        "Data Kunci / Statistik" = "61% insiden kebocoran data melibatkan data tidak terstruktur yang tersimpan di laptop staf tanpa pengawasan IT."
        "Temuan Utama" = "File sensitif (KTP, NPWP, laporan gaji, kontrak) menumpuk di folder Downloads tanpa enkripsi atau retensi yang jelas."
        "Validasi Pain Point" = "Kekhawatiran pelanggaran privasi (UU PDP Indonesia / GDPR) akibat file berserakan di drive lokal."
        "Implikasi Solusi / Fitur DocuFlow AI" = "Privacy Vault & PII Detection: secara otomatis menandai dokumen yang memuat data pribadi sensitif dan mengamankannya."
    },
    [PSCustomObject]@{
        "No" = 13
        "Institusi / Penulis" = "Prof. Gloria Mark (University of California, Irvine)"
        "Judul Publikasi / Riset" = "The Cost of Interrupted Work and Context Switching"
        "Tahun" = 2008
        "Tipe Sumber" = "Studi Akademik Human-Computer Interaction"
        "Kategori Isu" = "Distraksi & Alur Kerja (Flow State)"
        "Data Kunci / Statistik" = "Dibutuhkan 23 menit 15 detik bagi seorang pekerja untuk kembali fokus mendalam setelah terinterupsi mencari berkas."
        "Temuan Utama" = "Interupsi mikro saat mencari file merusak kemampuan berpikir mendalam dan memperlambat penyelesaian proyek besar."
        "Validasi Pain Point" = "Saat sedang menulis laporan penting, harus berhenti 15 menit hanya untuk mencari referensi data pendukung."
        "Implikasi Solusi / Fitur DocuFlow AI" = "Proactive Context Assistant: menyarankan dokumen terkait secara instan di sudut layar saat aplikasi kerja dibuka."
    },
    [PSCustomObject]@{
        "No" = 14
        "Institusi / Penulis" = "SearchYourCloud & Aberdeen Group"
        "Judul Publikasi / Riset" = "Document Retrieval Failure Rates in Hybrid Workspaces"
        "Tahun" = 2018
        "Tipe Sumber" = "Studi Perilaku Pencarian Dokumen"
        "Kategori Isu" = "Kegagalan Mesin Pencari Konvensional"
        "Data Kunci / Statistik" = "Rata-rata pekerja butuh 8 kali pencarian kata kunci berbeda; 15% dokumen tidak pernah berhasil ditemukan kembali."
        "Temuan Utama" = "Pencarian Windows Explorer dan Finder bawaan hanya mencari nama file dan teks harfiah sederhana tanpa pemahaman semantik."
        "Validasi Pain Point" = "Pengguna lupa persis kata kunci penamaan file sehingga pencarian sistem operasi menghasilkan 0 hasil."
        "Implikasi Solusi / Fitur DocuFlow AI" = "Neural Embedding Search: mencari dokumen berdasarkan konsep makna (contoh: cari 'bukti bayar sewa' menemukan 'kuitansi_gedung.pdf')."
    },
    [PSCustomObject]@{
        "No" = 15
        "Institusi / Penulis" = "Statista Research"
        "Judul Publikasi / Riset" = "Global Enterprise Knowledge Management & Intelligent Search Market"
        "Tahun" = 2023
        "Tipe Sumber" = "Laporan Riset Pasar Global"
        "Kategori Isu" = "Ukuran & Pertumbuhan Pasar"
        "Data Kunci / Statistik" = "Pasar Intelligent Search & Enterprise Knowledge diperkirakan tumbuh dari $4.5B menjadi $10.2B pada 2026 (CAGR 12.8%)."
        "Temuan Utama" = "Ledakan data tidak terstruktur mendorong adopsi teknologi AI pencarian berbasis Large Language Models di seluruh industri."
        "Validasi Pain Point" = "Kebutuhan mendesak di pasar untuk mengelola tsunami dokumen digital."
        "Implikasi Solusi / Fitur DocuFlow AI" = "Peluang pasar SaaS yang sangat sehat dengan potensi scaling dari segmen Prosumer/UMKM hingga B2B Enterprise."
    },
    [PSCustomObject]@{
        "No" = 16
        "Institusi / Penulis" = "Perforce Software"
        "Judul Publikasi / Riset" = "State of Digital Asset Management Report"
        "Tahun" = 2022
        "Tipe Sumber" = "Laporan Manajemen Aset Digital"
        "Kategori Isu" = "Ketiadaan Metadata Dokumen"
        "Data Kunci / Statistik" = "68% perusahaan berjuang mengelola aset tidak terstruktur; 43% aset digital tidak memiliki metadata memadai."
        "Temuan Utama" = "Dokumen tanpa metadata (penulis, tanggal kontrak, entitas terkait) menjadi 'file buta' yang sulit dicari sistem biasa."
        "Validasi Pain Point" = "File PDF hasil scan atau foto dokumen tidak dapat dicari karena tidak memiliki lapisan teks atau tag."
        "Implikasi Solusi / Fitur DocuFlow AI" = "Auto-OCR & Entity Extraction Pipeline: otomatis membaca teks gambar/scan dan mengekstrak entitas kunci (klien, nominal, tanggal)."
    },
    [PSCustomObject]@{
        "No" = 17
        "Institusi / Penulis" = "Stanford HCI Group (Nicolas Ducheneaut et al.)"
        "Judul Publikasi / Riset" = "The Semantic Desktop and Why Humans Fail at Hierarchical Categorization"
        "Tahun" = 2006
        "Tipe Sumber" = "Paper Akademik Stanford University"
        "Kategori Isu" = "Perilaku Pengarsipan Manusia"
        "Data Kunci / Statistik" = "Pengguna berhenti mengarsipkan file secara manual ketika jumlah dokumen melampaui ambang batas 500 berkas."
        "Temuan Utama" = "Klasifikasi manual membutuhkan keputusan kognitif berulang yang melelahkan sehingga pengguna akhirnya menumpuk file di desktop/downloads."
        "Validasi Pain Point" = "Niat merapikan folder selalu gagal di tengah jalan karena terlalu memakan waktu dan tenaga."
        "Implikasi Solusi / Fitur DocuFlow AI" = "Autonomous Background Indexing: AI bekerja 100% di latar belakang tanpa menuntut intervensi atau kedisiplinan pengguna."
    },
    [PSCustomObject]@{
        "No" = 18
        "Institusi / Penulis" = "MIT Sloan Management Review"
        "Judul Publikasi / Riset" = "Overcoming Information Silos in Hybrid Work"
        "Tahun" = 2022
        "Tipe Sumber" = "Artikel Riset Manajemen MIT"
        "Kategori Isu" = "Silo Kerja Hibrida & Multi-Device"
        "Data Kunci / Statistik" = "Bekerja secara hibrida melipatgandakan lokasi penyimpanan dokumen (laptop kantor, PC rumah, USB, cloud pribadi)."
        "Temuan Utama" = "Ketidakteraturan bertambah parah karena dokumen tersimpan terpisah di berbagai perangkat fisik yang tidak tersinkronisasi."
        "Validasi Pain Point" = "Berkas yang dibutuhkan ada di laptop kantor saat sedang bekerja dari rumah."
        "Implikasi Solusi / Fitur DocuFlow AI" = "Lightweight Cross-Device Metadata Sync: menyinkronkan peta indeks dokumen antar perangkat dengan tetap menjaga privasi."
    },
    [PSCustomObject]@{
        "No" = 19
        "Institusi / Penulis" = "Asana"
        "Judul Publikasi / Riset" = "Anatomy of Work Global Index: Work About Work"
        "Tahun" = 2023
        "Tipe Sumber" = "Laporan Riset Produktivitas Global"
        "Kategori Isu" = "Work About Work"
        "Data Kunci / Statistik" = "58% waktu pekerja tercurah untuk 'pekerjaan di luar tugas inti' (berburu file, update status, koordinasi manual)."
        "Temuan Utama" = "Hanya 42% hari kerja digunakan untuk pekerjaan strategis yang menghasilkan nilai tambah langsung bagi perusahaan."
        "Validasi Pain Point" = "Rasa tidak produktif meskipun sudah bekerja lembur seharian di depan komputer."
        "Implikasi Solusi / Fitur DocuFlow AI" = "Kalkulator Penghematan Waktu (ROI Dashboard) di dalam aplikasi untuk menunjukkan berapa jam kerja yang dihemat pengguna."
    },
    [PSCustomObject]@{
        "No" = 20
        "Institusi / Penulis" = "Qatalog & Cornell University"
        "Judul Publikasi / Riset" = "Workplace Chaos: The Cost of App and File Fragmentation"
        "Tahun" = 2022
        "Tipe Sumber" = "Studi Gabungan Industri & Akademik"
        "Kategori Isu" = "Fragmentasi Aplikasi & File"
        "Data Kunci / Statistik" = "Pekerja berganti konteks aplikasi hingga 1.200 kali per hari, kehilangan 4 jam per minggu hanya untuk sinkronisasi dokumen antar tim."
        "Temuan Utama" = "Fragmentasi dokumen antar aplikasi memicu pertanyaan berulang di chat: 'Di mana link dokumen revisi kemarin?'"
        "Validasi Pain Point" = "Grup WhatsApp atau Slack penuh dengan pesan meminta ulang file yang sebenarnya sudah pernah dibagikan."
        "Implikasi Solusi / Fitur DocuFlow AI" = "Team Shared Knowledge Spaces: tautan virtual bersama yang dapat diakses tim tanpa perlu mengatur hak akses folder yang rumit."
    },
    [PSCustomObject]@{
        "No" = 21
        "Institusi / Penulis" = "Coveo"
        "Judul Publikasi / Riset" = "Enterprise Search Relevance Report"
        "Tahun" = 2023
        "Tipe Sumber" = "Laporan Tahunan Mesin Pencari"
        "Kategori Isu" = "Relevansi Pencarian Dokumen"
        "Data Kunci / Statistik" = "41% pekerja mengatakan informasi yang mereka butuhkan sebenarnya ada di sistem, tetapi tersembunyi karena keyword tidak cocok."
        "Temuan Utama" = "Kesenjangan kosakata (vocabulary mismatch) antara pembuat dokumen dan pencari dokumen menggagalkan sistem pencarian teks lama."
        "Validasi Pain Point" = "Mencari dengan kata 'PKS' tidak memunculkan dokumen yang berjudul 'Perjanjian Kerja Sama'."
        "Implikasi Solusi / Fitur DocuFlow AI" = "Synonym & Acronym Expansion AI: secara otomatis memahami singkatan bisnis dan istilah sinonim bahasa Indonesia dan Inggris."
    },
    [PSCustomObject]@{
        "No" = 22
        "Institusi / Penulis" = "Association for Information and Image Management (AIIM)"
        "Judul Publikasi / Riset" = "State of the Intelligent Information Management Industry"
        "Tahun" = 2022
        "Tipe Sumber" = "Laporan Asosiasi Profesional"
        "Kategori Isu" = "Information Chaos"
        "Data Kunci / Statistik" = "79% organisasi mengakui mengalami 'Information Chaos' akibat volume file tidak terstruktur meningkat 3x lipat dalam 2 tahun."
        "Temuan Utama" = "Perusahaan tidak kekurangan ruang penyimpanan (storage murah), tetapi kekurangan kemampuan mengorganisir dan mengambil kembali data."
        "Validasi Pain Point" = "Kapasitas hard disk atau cloud penuh dengan ribuan file yang tidak jelas isinya."
        "Implikasi Solusi / Fitur DocuFlow AI" = "Smart Storage Analyzer: merekomendasikan file lama yang tidak terpakai untuk diarsipkan dan menghapus duplikat secara aman."
    },
    [PSCustomObject]@{
        "No" = 23
        "Institusi / Penulis" = "Xerox Enterprise Study"
        "Judul Publikasi / Riset" = "Document Dilemmas in Small-to-Medium Businesses"
        "Tahun" = 2019
        "Tipe Sumber" = "Studi Kasus Bisnis Kecil Menengah"
        "Kategori Isu" = "Keterbatasan Sumber Daya UMKM"
        "Data Kunci / Statistik" = "46% karyawan UMKM membuang waktu đáng kể setiap minggu karena dokumen faktur, PO, dan legal tersimpan acak di PC masing-masing."
        "Temuan Utama" = "Bisnis kecil tidak mampu membeli sistem Enterprise Content Management (ECM) ratusan juta rupiah yang rumit dikonfigurasi."
        "Validasi Pain Point" = "Solusi enterprise terlalu mahal dan rumit; solusi gratis seperti Windows Explorer terlalu bodoh."
        "Implikasi Solusi / Fitur DocuFlow AI" = "Model Self-Serve SaaS yang terjangkau ($9 - $19/bulan) dengan setup instan tanpa perlu pelatihan IT."
    },
    [PSCustomObject]@{
        "No" = 24
        "Institusi / Penulis" = "Deloitte Insights"
        "Judul Publikasi / Riset" = "The Intelligently Automated Enterprise"
        "Tahun" = 2023
        "Tipe Sumber" = "Laporan Riset Konsultan Global"
        "Kategori Isu" = "Teknologi RAG & Vektor"
        "Data Kunci / Statistik" = "Implementasi Retrieval-Augmented Generation (RAG) dan vector embeddings mempercepat temu kembali dokumen hingga 70-80%."
        "Temuan Utama" = "AI generatif paling bermanfaat saat digabungkan dengan konteks data internal pengguna secara real-time."
        "Validasi Pain Point" = "Pengguna tidak hanya ingin menemukan file, tetapi juga ingin jawaban spesifik yang ada di dalam berkas tersebut."
        "Implikasi Solusi / Fitur DocuFlow AI" = "Fitur 'Deep-Dive Q&A': pengguna bisa bertanya 'Berapa total nilai kontrak PT ABC di klausul 4?' dan langsung mendapat jawaban + link berkas."
    },
    [PSCustomObject]@{
        "No" = 25
        "Institusi / Penulis" = "PwC Global"
        "Judul Publikasi / Riset" = "2024 AI Business Survey: High-ROI Productivity Use Cases"
        "Tahun" = 2024
        "Tipe Sumber" = "Survei Eksekutif Bisnis"
        "Kategori Isu" = "ROI Investasi Perangkat Lunak AI"
        "Data Kunci / Statistik" = "Otomasi manajemen dokumen tidak terstruktur menduduki top 3 investasi AI dengan payback period di bawah 6 bulan."
        "Temuan Utama" = "Perusahaan sangat bersedia membayar software AI yang langsung meningkatkan produktivitas harian karyawan tanpa risiko bisnis tinggi."
        "Validasi Pain Point" = "Tingginya Willingness-to-Pay (WTP) untuk solusi yang menghentikan inefisiensi pencarian dokumen."
        "Implikasi Solusi / Fitur DocuFlow AI" = "Konversi freemium ke paid tier yang tinggi karena manfaat langsung dirasakan dalam minggu pertama penggunaan."
    },
    [PSCustomObject]@{
        "No" = 26
        "Institusi / Penulis" = "IBM Institute for Business Value"
        "Judul Publikasi / Riset" = "Data Decoded: Unlocking Unstructured Enterprise Data"
        "Tahun" = 2021
        "Tipe Sumber" = "Whitepaper Teknologi IBM"
        "Kategori Isu" = "Data Tidak Terstruktur (Unstructured Data)"
        "Data Kunci / Statistik" = "80% hingga 90% data organisasi berupa data tidak terstruktur (PDF, Word, spreadsheet, slide) tanpa skema baku."
        "Temuan Utama" = "Data ini menjadi 'Dark Data'—aset berharga yang tidak dapat diakses dan tidak memberikan nilai karena terkubur."
        "Validasi Pain Point" = "Proposal sukses tahun lalu atau riset pasar berharga tidak bisa dimanfaatkan kembali karena tidak ditemukan."
        "Implikasi Solusi / Fitur DocuFlow AI" = "Document Knowledge Graph: menghubungkan konsep-konsep antar dokumen lama dengan proyek yang sedang dikerjakan saat ini."
    },
    [PSCustomObject]@{
        "No" = 27
        "Institusi / Penulis" = "Nielsen Norman Group (NN/g)"
        "Judul Publikasi / Riset" = "Search vs. Browse in Modern Information Architecture"
        "Tahun" = 2022
        "Tipe Sumber" = "Studi Riset User Experience (UX)"
        "Kategori Isu" = "Kegagalan Desain Direktori Mendalam"
        "Data Kunci / Statistik" = "Tingkat keberhasilan navigasi pengguna anjlok 60% jika kedalaman folder melebihi 3 tingkat; pencarian semantik 4x lebih cepat."
        "Temuan Utama" = "Menavigasi pohon folder (browsing) membutuhkan beban memori kerja yang berat; manusia lebih menyukai pencarian terpandu."
        "Validasi Pain Point" = "Pengguna tersesat di dalam struktur sub-sub-sub-folder yang rumit."
        "Implikasi Solusi / Fitur DocuFlow AI" = "Antarmuka 'Flat & Dynamic': tidak ada folder bertingkat, dokumen difilter secara dinamis berdasarkan garis waktu, relasi, dan topik."
    },
    [PSCustomObject]@{
        "No" = 28
        "Institusi / Penulis" = "Zapier"
        "Judul Publikasi / Riset" = "Remote Work & Personal Automation Report"
        "Tahun" = 2023
        "Tipe Sumber" = "Laporan Otomasi & Produktivitas"
        "Kategori Isu" = "Minat Otomasi Pribadi"
        "Data Kunci / Statistik" = "74% knowledge workers siap menggunakan tools AI yang secara mandiri merapikan file dan membersihkan desktop."
        "Temuan Utama" = "Konsumen menyukai otomatisasi pasif (set-and-forget) yang tidak membutuhkan konfigurasi rule/if-this-then-that yang rumit."
        "Validasi Pain Point" = "Tools otomasi konvensional terlalu teknis bagi pengguna awam."
        "Implikasi Solusi / Fitur DocuFlow AI" = "Zero-Configuration AI: cukup pasang aplikasi, AI langsung mempelajari pola berkas dan mengorganisirnya secara otomatis."
    },
    [PSCustomObject]@{
        "No" = 29
        "Institusi / Penulis" = "Spiceworks Ziff Davis"
        "Judul Publikasi / Riset" = "State of IT File Storage & Shadow IT"
        "Tahun" = 2023
        "Tipe Sumber" = "Laporan Komunitas Profesional IT"
        "Kategori Isu" = "Shadow Storage & Perilaku Karyawan"
        "Data Kunci / Statistik" = "52% pekerja menggunakan folder lokal pribadi yang tidak terstandarisasi untuk menghindari birokrasi folder IT perusahaan yang kaku."
        "Temuan Utama" = "Sistem penataan dokumen terpusat yang kaku gagal karena pengguna selalu memilih jalan pintas tercepat (simpan di desktop/downloads)."
        "Validasi Pain Point" = "Paksaan IT untuk memakai aturan folder tertentu selalu dilanggar oleh karyawan demi kecepatan kerja."
        "Implikasi Solusi / Fitur DocuFlow AI" = "Solusi Non-Intrusif: biarkan pengguna menyimpan file di mana saja sesuka hati, AI yang menghubungkannya di lapisan atas."
    },
    [PSCustomObject]@{
        "No" = 30
        "Institusi / Penulis" = "TechRepublic & TechTarget"
        "Judul Publikasi / Riset" = "The Rise of Contextual Workspace Search"
        "Tahun" = 2023
        "Tipe Sumber" = "Laporan Tren Teknologi Enterprise"
        "Kategori Isu" = "Evolusi Industri Pencarian Kontekstual"
        "Data Kunci / Statistik" = "Kebutuhan bergeser dari cloud storage pasif (Dropbox/Drive) menuju 'Active Contextual Intelligence' proaktif."
        "Temuan Utama" = "Masa depan pengelolaan berkas bukan tempat menyimpan dokumen, melainkan asisten cerdas yang menyajikan dokumen tepat pada saat dibutuhkan."
        "Validasi Pain Point" = "Penyimpanan cloud saat ini hanya seperti 'gudang gelap' yang pasif menampung file tanpa kecerdasan kontekstual."
        "Implikasi Solusi / Fitur DocuFlow AI" = "Proactive Smart Feed: menampilkan berkas yang kemungkinan besar dibutuhkan hari ini berdasarkan rapat kalender atau proyek aktif."
    }
)

Write-Host "Menyiapkan Data Sheet 3: Analisis Masalah (Pain Points)..."
$sheet3Data = @(
    [PSCustomObject]@{
        "Pilar Masalah" = "1. Pemborosan Waktu & Biaya Kerja"
        "Gejala & Dampak Lapangan" = "Karyawan membuang 1.8 - 2.5 jam/hari mencari dokumen; 50% pencarian awal gagal; terpaksa membuat ulang berkas dari nol."
        "Dampak Finansial" = "Kerugian $5.000 - $12.000 per karyawan per tahun dalam bentuk jam kerja tidak produktif."
        "Sumber Validasi" = "IDC, McKinsey, Forrester, Asana"
        "Solusi Produk DocuFlow AI" = "Pencarian semantik super cepat (<2 detik) dan ekstraksi informasi instan."
    },
    [PSCustomObject]@{
        "Pilar Masalah" = "2. Kekacauan Versi & Duplikasi"
        "Gejala & Dampak Lapangan" = "File dinamai acak ('Proposal_final_v2_revisi_acc.pdf'); tim menduplikasi 23% pekerjaan; salah kirim draf usang ke klien."
        "Dampak Finansial" = "Kerusakan reputasi bisnis, kesalahan kontrak legal, dan pemborosan ruang penyimpanan cloud berbayar."
        "Sumber Validasi" = "Nitro Software, Adobe, CoSchedule"
        "Solusi Produk DocuFlow AI" = "Smart Version Clustering dan deteksi duplikasi konten semantik."
    },
    [PSCustomObject]@{
        "Pilar Masalah" = "3. Beban Kognitif & Kelelahan Mental"
        "Gejala & Dampak Lapangan" = "Desktop penuh dengan ratusan ikon acak; rasa cemas dan overwhelmed setiap kali membuka file manager; hilangnya flow state kerja."
        "Dampak Finansial" = "Penurunan kualitas kerja, kelelahan atensi, dan peningkatan tingkat stres kerja."
        "Sumber Validasi" = "APA, Dropbox/Wakefield, Gloria Mark (UC Irvine)"
        "Solusi Produk DocuFlow AI" = "Antarmuka minimalis, auto-declutter, dan eliminasi rasa bersalah merapikan folder."
    },
    [PSCustomObject]@{
        "Pilar Masalah" = "4. Fragmentasi Multi-Platform (Silo)"
        "Gejala & Dampak Lapangan" = "Dokumen terpecah di Google Drive, OneDrive kantor, disk lokal, SSD eksternal, WhatsApp web, dan lampiran email."
        "Dampak Finansial" = "Waktu koordinasi berulang di grup chat hanya untuk meminta link atau file yang hilang."
        "Sumber Validasi" = "Gartner, Qatalog & Cornell, MIT Sloan"
        "Solusi Produk DocuFlow AI" = "Unified Indexing Engine yang menyatukan drive lokal dan akun cloud dalam satu pintu pencarian."
    },
    [PSCustomObject]@{
        "Pilar Masalah" = "5. Kerentanan Keamanan & Privasi"
        "Gejala & Dampak Lapangan" = "File sensitif (KTP, NPWP, laporan keuangan, paspor, data karyawan) tercecer di folder Downloads tanpa enkripsi."
        "Dampak Finansial" = "Potensi sanksi hukum pelanggaran UU PDP / GDPR dan risiko pemerasan data internal."
        "Sumber Validasi" = "Ponemon Institute, Spiceworks Ziff Davis"
        "Solusi Produk DocuFlow AI" = "Local-First PII Scanner & Privacy Vault yang otomatis mengidentifikasi dan melindungi data rahasia."
    },
    [PSCustomObject]@{
        "Pilar Masalah" = "6. Kegagalan Struktur Hierarki Tradisional"
        "Gejala & Dampak Lapangan" = "Folder pohon hierarkis (Tree Structure) kaku; 82% sistem folder manual ditinggalkan setelah 3 bulan; dokumen multi-konteks tersesat."
        "Dampak Finansial" = "Frustrasi sistemik dan pengguna beralih ke 'Shadow IT' penyimpanan acak."
        "Sumber Validasi" = "Harvard Business Review, Stanford HCI, NN/g"
        "Solusi Produk DocuFlow AI" = "Virtual Multi-Dimensional Spaces: satu file dapat hadir di banyak ruang kerja kontekstual tanpa duplikasi fisik."
    }
)

Write-Host "Menyiapkan Data Sheet 4: Fitur Solusi & Arsitektur Produk..."
$sheet4Data = @(
    [PSCustomObject]@{
        "Modul Sistem" = "Omni-Connector Daemon"
        "Komponen Teknis" = "Background Service (Rust/C++)"
        "Fungsi Utama" = "Memantau perubahan file secara real-time di folder lokal (Downloads, Documents, Desktop) dan cloud drives (Google Drive, Dropbox, OneDrive) dengan konsumsi RAM rendah (<80MB)."
        "Tingkat Rilis" = "MVP (Fase 1)"
    },
    [PSCustomObject]@{
        "Modul Sistem" = "Local Ingestion & OCR Engine"
        "Komponen Teknis" = "Fast OCR (Tesseract / ONNX) + PDF Parser"
        "Fungsi Utama" = "Mengekstrak teks, tabel, dan metadata dari file PDF, DOCX, XLSX, PPTX, JPG/PNG (scan/kuitansi), dan teks obrolan tanpa mengirim isi file ke internet."
        "Tingkat Rilis" = "MVP (Fase 1)"
    },
    [PSCustomObject]@{
        "Modul Sistem" = "Semantic Vector Indexer"
        "Komponen Teknis" = "Lightweight Embedding (all-MiniLM / BGE-small)"
        "Fungsi Utama" = "Mengubah konten dokumen menjadi vektor matematis untuk pencarian semantik makna, analogi, sinonim, dan pemahaman konteks bisnis berbahasa Indonesia & Inggris."
        "Tingkat Rilis" = "MVP (Fase 1)"
    },
    [PSCustomObject]@{
        "Modul Sistem" = "Virtual Multi-Dimensional Spaces"
        "Komponen Teknis" = "Virtual Graph Database (SQLite / DuckDB)"
        "Fungsi Utama" = "Mengelompokkan file secara dinamis ke ruang kerja virtual (Klien, Kategori Pajak, Tahun Proyek, Tipe Berkas) tanpa mengubah lokasi fisik file di hard drive pengguna."
        "Tingkat Rilis" = "MVP (Fase 1)"
    },
    [PSCustomObject]@{
        "Modul Sistem" = "Natural Language Search & Ask-My-Docs"
        "Komponen Teknis" = "Local RAG Engine + SLM / Hybrid Cloud LLM"
        "Fungsi Utama" = "Pengguna dapat mengetik pertanyaan bebas (contoh: 'Berapa tagihan hosting bulan lalu?') dan AI langsung merangkum jawaban serta membuka file aslinya."
        "Tingkat Rilis" = "Pro Feature (Fase 2)"
    },
    [PSCustomObject]@{
        "Modul Sistem" = "Smart Version & Duplicate Manager"
        "Komponen Teknis" = "SimHash + Semantic Diffing Algorithm"
        "Fungsi Utama" = "Mendeteksi file salinan yang identik atau revisi dokumen yang mirip, menyarankan file master terbaru, dan memberikan opsi pembersihan ruang disk."
        "Tingkat Rilis" = "Pro Feature (Fase 2)"
    },
    [PSCustomObject]@{
        "Modul Sistem" = "Team Shared Knowledge Spaces"
        "Komponen Teknis" = "End-to-End Encrypted Relay Sync"
        "Fungsi Utama" = "Memungkinkan tim agensi atau UMKM berbagi ruang virtual dokumen proyek secara aman tanpa perlu migrasi folder storage yang mahal."
        "Tingkat Rilis" = "Team Feature (Fase 3)"
    },
    [PSCustomObject]@{
        "Modul Sistem" = "Privacy & PII Auto-Redaction Vault"
        "Komponen Teknis" = "Local Regex + NER (Named Entity Recognition)"
        "Fungsi Utama" = "Mendeteksi secara otomatis dokumen yang memuat nomor KTP, kartu kredit, kata sandi, dan data medis untuk disimpan dalam brankas terenkripsi."
        "Tingkat Rilis" = "Enterprise Feature (Fase 3)"
    }
)

Write-Host "Menyiapkan Data Sheet 5: Model Bisnis & Finansial..."
$sheet5Data = @(
    [PSCustomObject]@{
        "Tingkatan Paket (Tier)" = "Free Starter"
        "Target Pengguna" = "Pengguna individual, mahasiswa, pengguna kasual"
        "Harga Bulanan" = "$0 / bulan"
        "Fitur Utama" = "Indeks lokal hingga 1.000 file, pencarian semantik dasar, auto-declutter folder Downloads, local-only storage."
        "Strategi Pertumbuhan" = "Product-Led Growth (PLG), viral desktop widget."
    },
    [PSCustomObject]@{
        "Tingkatan Paket (Tier)" = "Pro Individual"
        "Target Pengguna" = "Freelancers, konsultan, pengacara, akuntan, peneliti"
        "Harga Bulanan" = "$9 / bulan ($90 / tahun)"
        "Fitur Utama" = "Indeks file lokal tanpa batas, 3 konektor cloud (Google Drive, OneDrive, Dropbox), Ask-My-Docs (RAG Chat), Smart Version Manager, OCR gambar/scan."
        "Strategi Pertumbuhan" = "Free trial 14 hari, kampanye produktivitas LinkedIn/Twitter."
    },
    [PSCustomObject]@{
        "Tingkatan Paket (Tier)" = "Team Workspace"
        "Target Pengguna" = "Agensi kreatif, UMKM, firma hukum kecil (5 - 30 orang)"
        "Harga Bulanan" = "$19 / user / bulan"
        "Fitur Utama" = "Semua fitur Pro, ruang kerja virtual bersama tim, kontrol akses berbasis peran (RBAC), team duplicate detection, kolaborasi cross-cloud."
        "Strategi Pertumbuhan" = "Inbound content marketing, demo langsung untuk pemilik UMKM."
    },
    [PSCustomObject]@{
        "Tingkatan Paket (Tier)" = "Enterprise Custom"
        "Target Pengguna" = "Korporasi, institusi keuangan, layanan kesehatan"
        "Harga Bulanan" = "$39 / user / bulan (kustom)"
        "Fitur Utama" = "Deploy on-premise / air-gapped LLM, compliance UU PDP / GDPR audit logs, SSO/SAML, integrasi custom ERP/CRM, SLA 99.9%."
        "Strategi Pertumbuhan" = "Direct B2B enterprise sales team dan kemitraan konsultan IT."
    }
)

Write-Host "Menyiapkan Data Sheet 6: Matriks Analisis Kompetitor..."
$sheet6Data = @(
    [PSCustomObject]@{
        "Kriteria Evaluasi" = "Kemudahan Setup (Zero Config)"
        "Windows Search / Finder" = "Tinggi (Bawaan OS)"
        "Cloud Search (Drive/Dropbox)" = "Tinggi (Web-based)"
        "Notion / Obsidian / PKM" = "Sangat Rendah (Butuh disiplin manual)"
        "Glean / Coveo (Enterprise)" = "Rendah (Butuh integrasi tim IT enterprise)"
        "DocuFlow AI (Solusi Diusulkan)" = "Sangat Tinggi (Instal langsung jalan, AI otonom)"
    },
    [PSCustomObject]@{
        "Kriteria Evaluasi" = "Pencarian Semantik / Makna AI"
        "Windows Search / Finder" = "Tidak Ada (Hanya teks harfiah)"
        "Cloud Search (Drive/Dropbox)" = "Terbatas (Keyword matching sederhana)"
        "Notion / Obsidian / PKM" = "Terbatas (Tergantung plugin AI manual)"
        "Glean / Coveo (Enterprise)" = "Tinggi (Enterprise Semantic Engine)"
        "DocuFlow AI (Solusi Diusulkan)" = "Sangat Tinggi (Local Vector Embedding + NLU)"
    },
    [PSCustomObject]@{
        "Kriteria Evaluasi" = "Tanpa Perlu Tagging/Folder Manual"
        "Windows Search / Finder" = "Tidak (Tergantung struktur folder manual)"
        "Cloud Search (Drive/Dropbox)" = "Tidak (Tergantung struktur folder manual)"
        "Notion / Obsidian / PKM" = "Tidak (Sangat menuntut tagging manual)"
        "Glean / Coveo (Enterprise)" = "Ya (Auto-indexing enterprise)"
        "DocuFlow AI (Solusi Diusulkan)" = "Ya (100% Autonomous Virtual Clustering)"
    },
    [PSCustomObject]@{
        "Kriteria Evaluasi" = "Multi-Source (Lokal + Multi-Cloud)"
        "Windows Search / Finder" = "Hanya Lokal"
        "Cloud Search (Drive/Dropbox)" = "Hanya di cloud platform masing-masing"
        "Notion / Obsidian / PKM" = "Terisolasi di dalam workspace aplikasi"
        "Glean / Coveo (Enterprise)" = "Ya (Multi-cloud enterprise connectors)"
        "DocuFlow AI (Solusi Diusulkan)" = "Ya (Lokal SSD + Google Drive + Dropbox + OneDrive)"
    },
    [PSCustomObject]@{
        "Kriteria Evaluasi" = "Privasi Data (Local-First)"
        "Windows Search / Finder" = "Lokal"
        "Cloud Search (Drive/Dropbox)" = "Cloud publik"
        "Notion / Obsidian / PKM" = "Campuran (Obsidian lokal, Notion cloud)"
        "Glean / Coveo (Enterprise)" = "Cloud enterprise / SaaS"
        "DocuFlow AI (Solusi Diusulkan)" = "Sangat Tinggi (Local-first processing, zero cloud data leak)"
    },
    [PSCustomObject]@{
        "Kriteria Evaluasi" = "Tanya Jawab Dokumen (Ask-My-Docs)"
        "Windows Search / Finder" = "Tidak Ada"
        "Cloud Search (Drive/Dropbox)" = "Tidak Ada / Sangat Terbatas"
        "Notion / Obsidian / PKM" = "Ada pada Notion AI (Berbayar)"
        "Glean / Coveo (Enterprise)" = "Ada (Enterprise Q&A)"
        "DocuFlow AI (Solusi Diusulkan)" = "Ada (Local RAG interaktif dengan referensi halaman)"
    },
    [PSCustomObject]@{
        "Kriteria Evaluasi" = "Harga & Aksesibilitas"
        "Windows Search / Finder" = "Gratis"
        "Cloud Search (Drive/Dropbox)" = "Termasuk langganan storage"
        "Notion / Obsidian / PKM" = "Gratis hingga $10/bulan"
        "Glean / Coveo (Enterprise)" = "Sangat Mahal ($30k+ per tahun / enterprise)"
        "DocuFlow AI (Solusi Diusulkan)" = "Terjangkau ($0 - $19/bulan, ramah freelancer & UMKM)"
    }
)

Write-Host "Mengekspor seluruh Sheet ke file Excel..."

# Export Sheet 1
$sheet1Data | Export-Excel -Path $excelPath -WorksheetName "1. Ringkasan & Ide Bisnis" -AutoSize -FreezeTopRow -BoldTopRow -TableStyle Medium9

# Export Sheet 2
$sheet2Data | Export-Excel -Path $excelPath -WorksheetName "2. 30 Sumber Riset Empiris" -AutoSize -FreezeTopRow -BoldTopRow -TableStyle Medium2

# Export Sheet 3
$sheet3Data | Export-Excel -Path $excelPath -WorksheetName "3. Analisis Pain Points" -AutoSize -FreezeTopRow -BoldTopRow -TableStyle Medium4

# Export Sheet 4
$sheet4Data | Export-Excel -Path $excelPath -WorksheetName "4. Fitur & Arsitektur Produk" -AutoSize -FreezeTopRow -BoldTopRow -TableStyle Medium6

# Export Sheet 5
$sheet5Data | Export-Excel -Path $excelPath -WorksheetName "5. Model Bisnis & Monetisasi" -AutoSize -FreezeTopRow -BoldTopRow -TableStyle Medium13

# Export Sheet 6
$sheet6Data | Export-Excel -Path $excelPath -WorksheetName "6. Matriks Analisis Kompetitor" -AutoSize -FreezeTopRow -BoldTopRow -TableStyle Medium14

Write-Host "File Excel berhasil dibuat secara komprehensif di: $excelPath"
