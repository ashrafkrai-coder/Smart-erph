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
6. Pada Google Sheet eRPH, muat semula halaman. Pilih Smart eRPH AI > Sambungkan fail eRPH semasa untuk PWA sekali sahaja.
7. Deploy folder PWA ini ke Vercel.
8. Dalam Vercel Project Settings > Environment Variables, tambah:
   - ERPH_SCRIPT_URL : URL Apps Script yang berakhir dengan /exec.
   - ERPH_API_TOKEN : nilai yang sama seperti PWA_API_TOKEN dalam Apps Script.
9. Deploy semula Vercel.

Penggunaan

1. Pilih Minggu 1 hingga 45.
2. Pilih tarikh Isnin. PWA akan mengisi tarikh Selasa hingga Jumaat secara automatik.
3. Semua hari ditanda secara default. Nyah tanda hari cuti.
4. Tekan Jana eRPH Mingguan.

Sistem akan mengemas kini semua tab dan semua slot bagi hari yang ditanda. Tajuk dalam slot sedia ada digunakan jika sepadan dengan DSKP/sukatan. Jika tiada tajuk, sistem memilih tajuk berdasarkan minggu RPT.

PWA memanggil API Vercel sendiri. Vercel menyambung ke Apps Script di belakang, jadi pelayar tidak lagi disekat oleh CORS dan token tidak dipaparkan dalam PWA.
