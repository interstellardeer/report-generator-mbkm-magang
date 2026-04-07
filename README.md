# Generate Laporan MBKM menggunakan LLM

Repositori ini menyediakan sistem bagi mahasiswa untuk mengekstrak data dari dokumen magang menjadi laporan MBKM (berformat `.docx`) secara otomatis dengan bantuan kecerdasan buatan (AI). 

Proyek ini dilengkapi dengan panduan "Agent Prompt" spesifik pada berkas `AGENT_LAPORAN_MBKM.md` yang memastikan AI memproses dokumen magang Anda dengan batasan privasi yang jelas, serta mengembalikan data terstruktur yang dijamin kompatibel dengan generator laporan kita.

## Persiapan Dokumen
Sebelum memulai, kumpulkan dokumen-dokumen terkait kegiatan magang Anda (boleh berupa teks biasa, PDF, atau CSV jika didukung oleh LLM Anda):
- SK Magang / Surat Pengajian awal (Offering Letter)
- Rekap absensi / Kehadiran kerja
- Logbook pekerjaan mingguan / Catatan harian
- Form Evaluasi Penyelia (jika diizinkan untuk diproses)

## Cara Penggunaan

Terdapat dua pendekatan utama untuk menggunakan sistem AI ini: melalui website chatbot antarmuka, atau melalui tool Agent/IDE *developer*. 

### Metode 1: Menggunakan Website Chatbot LLM (Claude.ai, ChatGPT, Gemini, dll.)
Metode ini adalah cara paling sederhana jika Anda tidak menggunakan aplikasi *coding* bertenaga AI.
1. Buka berkas `AGENT_LAPORAN_MBKM.md` pada komputer Anda dan **copy / salin seluruh isinya**.
2. Masuk ke chatbot LLM preferensi Anda, paste seluruh isi tersebut sebagai awalan (atau masukkan ke fitur *System Prompt* /*Custom Instructions* jika ada), lalu kirimkan.
3. Chatbot sekarang telah bertindak sebagai "Agent Laporan MBKM". 
4. Di pesan selanjutnya, **lampirkan (upload) berkas `.js` generator (contoh: `generate_laporan_mbkm.js` atau `generate_laporan_mbkm_claude.js`)** agar AI memahami bentuk asli konsistensi laporannya, lalu lampirkan juga dokumen magang Anda (SK, logbook, absensi).
5. LLM akan membaca dokumen-dokumen tersebut, memilah data yang bersifat publik/rahasia, dan membalas dengan struktur kode variabel bernama `CONFIG`.

### Metode 2: Menggunakan AI Agent Tools / IDE (Sangat Direkomendasikan)

> **Rekomendasi:** Kami **sangat menyarankan** pendekatan ini jika AI Agent Anda mendukung fitur *Tool Use* (eksekusi alat mandiri berkapabilitas terminal/berkas). Agent seperti Claude Code, Antigravity, atau Cursor dapat mengambil alih keseluruhan orkestrasi pembuatannya. Mereka tidak bertele-tele mencetak kode di layar, layaknya obrolan biasa, namun dapat merubah *script* `.js` dan langsung menyerahkan *file output* `.docx` ke *workspace* Anda!

Metode ini digunakan jika Anda bekerja menggunakan IDE atau agen terminal pintar. Ini jelas lebih mudah karena agen AI dapat membaca lingkungan *folder* secara otomatis:
1. Pertama, pastikan Anda telah *clone* atau mengunduh repositori ini ke komputer Anda. Kemudian, buka folder proyek (`report-generator`) tersebut di dalam IDE Anda atau terminal yang sudah terpasang CLI agen AI.
2. Panggil agen Anda dengan perintah (Command / *Chat Box*).
3. Sematkan referensi terhadap berkas instruksi dan dokumen magangnya. Berikan *prompt* seperti contoh berikut ini:
   > "Tolong gunakan panduan di `@AGENT_LAPORAN_MBKM.md` untuk memproses dokumen magang saya. Ini referensi kode laporannya `@generate_laporan_mbkm.js`, serta data saya: `@logbook.txt` dan `@sk-magang.pdf`. Lakukan ekstraksi sesuai instruksi privasinya dan berikan blok `CONFIG` nya."
4. Agent berkapabilitas *Tool Use* akan secara leluasa membaca instruksi, menata dokumen yang ditunjuk, kemudian otomatis mengganti kode `CONFIG` di skrip (`generate_laporan_mbkm.js`) lalu mengeksekusinya untuk menghasilkan berkas `.docx` instan tanpa perlu campur tangan *copy-paste* Anda!

---

## Tahap Akhir: Menghasilkan Laporan (Generate)

Setelah AI Agent (dari Metode 1 maupun Metode 2) rampung dan memberikan Anda blok **CONFIG**, langkah mengunci generasinya adalah sebagai berikut:

1. Salin hasil kode `CONFIG` yang diberikan AI.
2. Buka berkas `generate_laporan_mbkm.js` menggunakan *text editor* atau IDE Anda.
3. Cari letak variabel `const CONFIG = { ... }` (umumnya berada pada baris 38 hingga 120).
4. **Paste / timpa (Replace)** blok variabel tersebut dengan hasil konfigurasi dari AI.
5. Perhatikan beberapa indikator yang bertuliskan `"PERLU DITULIS MAHASISWA"`. AI sengaja mengosongkan *field* tersebut (seperti Latar Belakang, Kesimpulan) karena hal itu menuntut refleksi pribadi Anda sendiri. Isi bagian tersebut dengan narasi manual.
6. Buka terminal atau konsol perintah (*Command Prompt*), jalankan *script* dengan *command*:
   ```bash
   node generate_laporan_mbkm.js
   ```
7. *Ta-da!* Berkas Laporan Anda (`Laporan_MBKM_Template.docx`) telah jadi dan siap dikaji ulang!

## ⚠️ Perhatian Privasi Data
Meskipun instruksi di dalam `AGENT_LAPORAN_MBKM.md` telah mewanti-wanti sistem AI agar menyortir dan menahan konten rahasia perusahaan, **tetaplah bijaksana sebagai penentu akhir**. JANGAN mengunggah atau melampirkan berkas yang tertera sangat jelas merujuk pada Kerahasiaan Tingkat Tinggi, kekayaan intelektual murni tak terpublikasi (source code internal, data keuangan pelanggan), atau dokumen rahasia instansi, ke alat bantu kecerdasan buatan pihak ketiga mana pun.
