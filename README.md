# Arahan Penggunaan — HEP Kolej Komuniti Mas Gading

## Cara Buka Aplikasi

1. Buka **File Explorer**, navigasi ke folder ini.
2. Klik dua kali **`index.html`**.  
   Ia akan dibuka dalam Chrome atau Edge secara automatik.
3. Tiada pemasangan, tiada internet, tiada pelayan diperlukan.

---

## Penjana Kertas Kerja (Fasa 1)

### Isi borang

| Bahagian | Kandungan |
|---|---|
| **01 Maklumat Program** | Nama, penganjur, kluster, peringkat, tarikh, masa, tempat, sasaran, peserta, kolaborasi |
| **02 Tujuan & Hasil** | Latar belakang, objektif (satu per baris), impak, kesimpulan |
| **03 Pelaksanaan** | Carta perbatuan, tentatif program, senarai AJK — menggunakan jadual dengan butang tambah/buang baris |
| **04 Anggaran Perbelanjaan** | Sumber peruntukan dan butiran bajet dalam jadual |
| **05 Peserta & Pegawai** | Senarai peserta, pegawai dihubungi, e-mel, alamat |

### Jadual dinamik (tambah/buang baris)

- Klik **+ Tambah baris** untuk tambah baris baharu dalam jadual.
- Klik butang **×** di hujung baris untuk buang baris tersebut.
- Kolum **Jumlah (RM)** dalam bajet dikira **secara automatik**.

### Jana dokumen Word

1. Lengkapkan semua medan bertanda `*`.
2. Klik butang **Jana Word ➜** (di bahagian bawah).
3. Fail `.docx` akan dimuat turun ke folder *Downloads*.
4. Buka fail dalam **Microsoft Word** atau **WPS Office**.
5. Semak format, cetak dan proses kelulusan secara manual.

> **Nota:** Ruangan tandatangan (disediakan oleh, disemak oleh, semakan kewangan, kelulusan) **dibiarkan kosong** dalam dokumen. Lengkapkan secara tulis tangan atau taip dalam Word.

### Simpan & buka draf

- **Simpan draf** — Muat turun fail `Draf_Kertas_Kerja.json` ke *Downloads*. Simpan fail ini.
- **Buka draf** — Pilih fail JSON yang disimpan untuk sambung mengisi borang.
- Data borang **tidak disimpan automatik** — simpan draf sebelum menutup pelayar.

### Contoh rekaan

Klik **Isi contoh rekaan** untuk melihat contoh borang yang lengkap (data rekaan sahaja, bukan data sebenar).

---

## Kandungan Dokumen yang Dijana

Dokumen Word mengandungi:

1. **Muka depan** — nama program, tarikh, tempat, penganjur, ruangan disediakan oleh
2. **Jadual maklumat program** — semua butiran asas
3. **Seksyen bernombor** (1.0–8.0):
   - 1.0 Latar Belakang
   - 2.0 Objektif Program
   - 3.0 Carta Jawatankuasa (→ Lampiran 1)
   - 4.0 Carta Perbatuan
   - 5.0 Anggaran Peruntukan (→ Lampiran 2)
   - 6.0 Tentatif Program
   - 7.0 Impak Program
   - 8.0 Kesimpulan
4. **Pengesahan institusi** — tandatangan disediakan oleh, disemak oleh, semakan kewangan, kelulusan (semua kosong)
5. **Pegawai dihubungi** dan alamat institusi
6. **Lampiran 1** — Senarai AJK
7. **Lampiran 2** — Anggaran perbelanjaan dengan jumlah
8. **Lampiran 3** — Senarai peserta

---

## Modul Yang Tersedia

| Menu | Status | Penerangan |
|---|---|---|
| **Kertas Kerja** | ✅ Sedia | Jana kertas kerja program lengkap (.docx) |
| **Memo** | ✅ Sedia | Jana memo rasmi HEP |
| **Surat Lantikan AJK** | ✅ Sedia | Surat pelantikan ahli jawatankuasa |
| **Surat Lantikan Penceramah** | ✅ Sedia | Surat jemputan penceramah |
| **Penjanaan Sijil** | ✅ Sedia | Sijil bernombor siri untuk semua peserta |
| **Laporan Program** | ✅ Sedia | Laporan pelaksanaan program |
| **Penilaian Kursus** | ✅ Sedia | Borang penilaian peserta + ringkasan Word |
| **Analisis Keberkesanan** | ✅ Sedia | Carta statistik & log rekod dokumen dijana |

---

## Analisis Keberkesanan

Menu **Analisis Keberkesanan** memaparkan:

- **Statistik ringkasan** — jumlah dokumen dijana mengikut jenis
- **Carta kluster** — program mengikut kluster (Kertas Kerja)
- **Carta peringkat** — Kolej / Zon / Kebangsaan dll
- **Trend bulanan** — semua jenis dokumen, 12 bulan terkini
- **Log terkini** — 20 rekod terakhir dengan nama, kluster dan tarikh

Rekod disimpan automatik ke **localStorage** pelayar setiap kali dokumen dijana.
Klik **🗑 Padam Semua Rekod** untuk memadam statistik (tidak mempengaruhi draf).

---

## Fail Projek

| Fail | Keterangan |
|---|---|
| `index.html` | Aplikasi utama — buka dalam pelayar |
| `docx.js` | Penjana Word (DOCX) tanpa pergantungan rangkaian |
| `test-docx-ps.ps1` | Skrip ujian PowerShell (74 ujian) |
| `README.md` | Arahan penggunaan ini |

---

## Nota Teknikal

- Semua data diproses **dalam pelayar** sahaja — tiada data dihantar ke pelayan.
- Fail draf JSON dan dokumen DOCX mungkin mengandungi maklumat peserta — simpan di lokasi yang selamat.
- Ujian dijalankan dengan: `powershell -ExecutionPolicy Bypass -File test-docx-ps.ps1`
- Aplikasi berfungsi dalam Chrome, Edge dan Firefox (versi terkini).
- Fail DOCX yang dijana telah disahkan mengandungi pakej ZIP sah, XML Word betul dan semua lampiran.

---

## Ujian Penerimaan

Sebelum penggunaan rasmi, jalankan semakan berikut:

- [ ] Isi borang kosong sepenuhnya dan jana dokumen Word
- [ ] Buka dokumen dalam Microsoft Word — tiada mesej ralat
- [ ] Semua 8 seksyen dan 3 lampiran hadir
- [ ] Jumlah bajet sepadan dengan kiraan manual
- [ ] Ruangan tandatangan dan kelulusan kosong
- [ ] Simpan draf → tutup tab → buka semula → buka draf → semak data pulih
- [ ] Isi contoh rekaan → jana dokumen → sahkan data contoh tidak menjadi nilai lalai selepas muat semula halaman
- [ ] Semak dokumen panjang (10+ peserta, 10+ baris bajet) dalam Word

*Aplikasi ini untuk kegunaan setempat seorang pengguna. Untuk penggunaan berkongsi, tambah pelayan, pengesahan pengguna dan kawalan akses sebelum diterbitkan.*
