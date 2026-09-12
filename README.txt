SMART eRPH AI PWA — Pemasangan

1. Dalam Google Sheet eRPH, buka Extensions > Apps Script.
2. Gantikan Code.gs dengan Code.gs yang disediakan bersama PWA ini.
3. Dalam Project Settings > Script Properties, tetapkan:
   - GEMINI_API_KEY : API key Gemini anda.
   - PWA_API_TOKEN : cipta satu token rahsia sendiri, contohnya gabungan huruf, nombor dan simbol yang panjang.
4. Pilih Deploy > New deployment > Web app.
   Jalankan sebagai: Me.
   Pilih akses yang sesuai untuk guru. Jika memilih Anyone, PWA_API_TOKEN wajib dirahsiakan.
5. Salin URL yang berakhir dengan /exec.
6. Host folder PWA ini di Vercel atau Firebase Hosting.
7. Buka PWA, tekan ikon tetapan, dan masukkan URL /exec serta PWA_API_TOKEN sekali sahaja.

Penggunaan

1. Pilih Minggu 1 hingga 45.
2. Pilih tarikh Isnin. PWA akan mengisi tarikh Selasa hingga Jumaat secara automatik.
3. Semua hari ditanda secara default. Nyah tanda hari cuti.
4. Tekan Jana eRPH Mingguan.

Sistem akan mengemas kini semua tab dan semua slot bagi hari yang ditanda. Tajuk dalam slot sedia ada digunakan jika sepadan dengan DSKP/sukatan. Jika tiada tajuk, sistem memilih tajuk berdasarkan minggu RPT.
