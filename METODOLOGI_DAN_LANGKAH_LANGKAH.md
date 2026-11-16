# Metodologi & Langkah-Langkah: Dari Masalah File Clutter Menjadi Ide Bisnis

Dokumen ini menguraikan secara rinci metodologi riset dan **langkah-langkah terstruktur** yang diambil dalam menganalisis permasalahan kesulitan menemukan file hingga bertransformasi menjadi ide bisnis produk kecerdasan buatan (*AI-powered SaaS*) yang tervalidasi dan siap dieksekusi.

---

## 🧭 Ikhtisar 8 Langkah Utama (Step-by-Step Process)

```
[Langkah 1: Perumusan Masalah & Scope]
       │
       ▼
[Langkah 2: Pengumpulan 30 Sumber Riset Empiris]
       │
       ▼
[Langkah 3: Pemetaan Pain Points & Pemodelan Finansial]
       │
       ▼
[Langkah 4: Analisis Akar Masalah (Root Cause Analysis)]
       │
       ▼
[Langkah 5: Konseptualisasi Solusi & Arsitektur Produk DocuFlow AI]
       │
       ▼
[Langkah 6: Penyusunan Business Model Canvas & Monetisasi]
       │
       ▼
[Langkah 7: Desain & Otomasi Spreadsheet Excel Multi-Sheet]
       │
       ▼
[Langkah 8: Version Control, Deployment & GitHub Repository Setup]
```

---

## 📌 Rincian Langkah-Langkah Eksekusi

### Langkah 1: Perumusan Masalah & Batasan Riset (Problem Definition & Scoping)
- **Topik Masalah**: *"Sering kesulitan menemukan file tertentu karena terlalu banyak dokumen yang tersimpan tanpa pengelompokkan yang jelas."*
- **Tujuan**: Memetakan apakah masalah ini merupakan kendala individual yang sepele atau masalah sistemik global yang layak diubah menjadi sebuah ide bisnis komersial.
- **Hipotesis Awal**:
  1. Pengguna mengalami friksi mental tinggi saat harus memutuskan folder penyimpanan setiap kali mengunduh atau menyimpan dokumen baru.
  2. Kebiasaan menyimpan langsung di folder *Downloads* atau *Desktop* terjadi karena kebutuhan kecepatan, bukan kemalasan semata.
  3. Struktur folder hierarkis konvensional tidak lagi relevan dengan volume informasi digital modern.

---

### Langkah 2: Pengumpulan & Telaah 30 Sumber Riset Empiris (Systematic Literature & Industry Review)
Untuk menjamin validitas riset, dilakukan pengumpulan data sekunder dari **30 sumber kredibel lintas domain**:
- **Lembaga Riset Pasar & Konsultan Manajemen**: IDC, McKinsey Global Institute, Gartner, Forrester Research, Deloitte Insights, PwC Global.
- **Raksasa Teknologi & Vendor Dokumen**: Microsoft, Adobe Document Cloud, Dropbox, Nitro Software, Asana, Zapier, IBM Institute for Business Value.
- **Institusi Akademik & Riset Psikologi**: Stanford University (HCI Group), UC Irvine (Prof. Gloria Mark), Cornell University & Qatalog, American Psychological Association (APA), Harvard Business Review (HBR), Nielsen Norman Group (NN/g).
- **Asosiasi Profesional & Industri DAM**: AIIM, Perforce Software, CoSchedule, Coveo, SearchYourCloud & Aberdeen Group, Xerox, Statista, Spiceworks Ziff Davis, TechRepublic/TechTarget.

Setiap sumber diekstraksi ke dalam matriks data: nama penulis, judul riset, tahun rilis, tipe data (kuantitatif/kualitatif), metrik statistik kunci, temuan utama, dan relevansinya terhadap rancangan fitur produk.

---

### Langkah 3: Pemetaan Pain Points & Pemodelan Finansial Masalah (Impact Modeling)
Dari ke-30 sumber, dilakukan pengelompokan (*clustering*) masalah menjadi 6 pilar dampak negatif:
1. **Dampak Waktu & Produktivitas**: Rata-rata 1.8 – 2.5 jam per hari terbuang hanya untuk mencari dokumen (McKinsey, IDC).
2. **Duplikasi Aset & Kesalahan Revisi**: 23% pekerjaan diduplikasi karena aset asli tidak ditemukan (CoSchedule); 50% pekerja frustrasi karena ketidakpastian versi kanonikal (Nitro).
3. **Beban Kognitif & Stres**: 62% pekerja merasa stres akibat desktop dan folder yang berantakan (Dropbox/Wakefield); penumpukan file meningkatkan hormon kortisol (APA).
4. **Silo & Fragmentasi Penyimpanan**: Dokumen terpecah di Google Drive, OneDrive, SSD lokal, dan chat WhatsApp (Gartner, Qatalog).
5. **Kerentanan Keamanan & Regulasi**: 61% kebocoran data melibatkan data tidak terstruktur di laptop karyawan tanpa enkripsi (Ponemon Institute, UU PDP/GDPR).
6. **Kegagalan Taksonomi Manual**: 82% skema folder manual ditinggalkan dalam waktu 3 bulan (Harvard Business Review).

**Pemodelan Finansial**:
- Jika seorang karyawan memiliki gaji rata-rata Rp 10.000.000 / bulan (Rp 62.500 / jam kerja):
- Kehilangan 2 jam/hari = 40 jam/bulan = **Rp 2.500.000 pemborosan gaji per karyawan setiap bulan**.
- Untuk perusahaan dengan 20 karyawan, kerugian mencapai **Rp 50.000.000 per bulan** hanya karena inefisiensi pencarian dokumen.

---

### Langkah 4: Analisis Akar Masalah (Root Cause Analysis / The "Why")
Mengapa orang tidak mengelompokkan dokumen mereka dengan rapi?
- **Cognitive Friction**: Otak manusia bekerja secara asosiatif dan multi-konteks, sedangkan sistem operasi komputer (Windows/macOS) memaksakan struktur hierarki pohon tunggal (*Single-Parent Tree Structure*).
- **Inkonsistensi Waktu**: Kategori folder yang dibuat bulan Januari terasa tidak masuk akal saat dicari bulan November.
- **Friction of Filing vs Speed of Working**: Saat dikejar tenggat waktu (*deadline*), tindakan mengklik 4 subfolder untuk menyimpan berkas dianggap menghambat, sehingga tombol *"Save to Downloads"* selalu menjadi pilihan termudah.

**Kesimpulan Kunci**: *Jangan memaksa manusia mengubah perilakunya untuk melayani folder; biarkan kecerdasan buatan melayani kebiasaan alami manusia.*

---

### Langkah 5: Konseptualisasi Solusi & Desain Produk (DocuFlow AI)
Berdasarkan akar masalah di atas, dirancang produk solusi bernama **DocuFlow AI**:
- **Arsitektur Local-First Daemon**: Aplikasi desktop ringan (berjalan di background dengan RAM <80MB) yang memindai folder lokal dan menyinkronkan metadata cloud tanpa memindahkan lokasi file fisik pengguna.
- **Virtual Multi-Dimensional Spaces**: Menggunakan basis data grafis virtual untuk mengaitkan 1 file ke banyak kategori (contoh: Klien X, Pajak 2026, Kontrak) secara otomatis.
- **On-Device Vector Embedding**: Mengekstrak teks melalui OCR lokal dan menghasilkan representasi vektor semantik agar pencarian bahasa alami dapat memahami makna dan sinonim.
- **Interactive Local RAG ('Ask-My-Docs')**: Pengguna dapat mengajukan pertanyaan spesifik dan menerima jawaban yang dirangkum langsung dari isi berkas disertai referensi halaman.

---

### Langkah 6: Penyusunan Model Bisnis & Strategi Go-To-Market (GTM)
- **Model Pendapatan (Freemium SaaS)**:
  - *Free Tier*: Indeks lokal hingga 1.000 berkas (Product-Led Growth).
  - *Pro Tier ($9/bulan)*: Berkas lokal tanpa batas, integrasi cloud storage, fitur RAG Q&A.
  - *Team Tier ($19/user/bulan)*: Ruang kerja virtual bersama, deteksi duplikasi tim.
  - *Enterprise ($39+/user/bulan)*: On-premise air-gapped LLM, kepatuhan UU PDP/GDPR.
- **Unit Economics**:
  - Customer Acquisition Cost (CAC) diperkirakan $35 via organic virality.
  - Lifetime Value (LTV) $216 (asumsi retensi 24 bulan @ $9/bulan).
  - Rasio LTV/CAC = 6.17x (sangat sehat untuk model bisnis SaaS).

---

### Langkah 7: Desain & Otomasi Spreadsheet Excel Multi-Sheet
Untuk menyajikan data riset secara profesional, dibangun skrip otomasi PowerShell (`generate_excel.ps1`) menggunakan pustaka `ImportExcel`:
1. Menyusun 6 lembar kerja (*Worksheets*) berstruktur logis.
2. Menerapkan skema pewarnaan modern (*TableStyle Medium*), membekukan baris judul (*Freeze Top Row*), dan mengaktifkan penyesuaian lebar kolom otomatis (*AutoSize*).
3. Memastikan kompatibilitas penuh dengan Microsoft Excel, Google Sheets, dan LibreOffice Calc.

---

### Langkah 8: Pengendalian Versi (Git) & GitHub Setup
1. Menginisialisasi repositori Git lokal di direktori proyek.
2. Mengonfigurasi identitas repositori (`user.name` dan `user.email`).
3. Menambahkan seluruh artefak (Excel, Markdown dokumentasi, skrip otomatisasi) ke dalam *staging area* (`git add .`).
4. Melakukan *commit* terstruktur dengan deskripsi komprehensif.
5. Menghubungkan *remote origin* ke akun GitHub target (`salsabilaivo-netizen`).
6. Melakukan upaya dorong (*push*) ke repositori GitHub dan menyiapkan panduan jika token memerlukan penyesuaian cakupan izin (*scope*).
