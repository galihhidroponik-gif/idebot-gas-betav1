/**
 * =========================================================================
 * PANDUAN INSTALASI & SETUP SISTEM OMNIBOT
 * =========================================================================
 * STEP 1: DAPATKAN AKSES LIBRARY
 *   - Hubungi admin Idekreatifa untuk mendapatkan IDs library OmniBot melalui WA 089509063690.
 * 
 * STEP 2: PASANG LIBRARY
 *   - Di menu sebelah kiri editor script, klik tanda tambah (+) pada bagian "Libraries" (Pustaka).
 *   - Masukkan IDs library OmniBot, lalu klik "Cari".
 *   - Pilih versi terbaru, pastikan pengidentifikasinya bernama "OmniBot", lalu klik "Tambahkan" (Save).
 * 
 * STEP 3: OTORISASI SISTEM (WAJIB PERTAMA KALI)
 *   - Di editor script ini, lihat kotak pilihan fungsi di bagian atas (sebelah tombol "Jalankan").
 *   - Pilih fungsi "otorisasiSistem".
 *   - Klik tombol "Jalankan".
 *   - Ikuti pop-up Google: Review Permissions -> Pilih Email -> Advanced -> Go to... -> Allow.
 * 
 * STEP 4: DEPLOY APLIKASI WEB
 *   - Klik tombol biru "Deploy" (Terapkan) di pojok kanan atas -> "New deployment" (Penerapan baru).
 *   - Pilih ikon gerigi (Select type) -> "Web app" (Aplikasi web).
 *   - Execute as (Jalankan sebagai): "Me" (Email Anda).
 *   - Who has access (Siapa yang memiliki akses): "Anyone" (Siapa saja) -> INI WAJIB!
 *   - Klik Deploy dan salin (copy) "Web app URL" yang muncul.
 * 
 * STEP 5: PASANG WEBHOOK
 *   - Paste URL Web App dari Step 4 ke sheet setting B2 (URL WEEBHOOK)
 *   - Paste URL Web App dari Step 4 ke menu Webhook di dashboard Starsender atau OneSender.
 * 
 * STEP 6: TEST KONEKSI PENGIRIMAN PESAN
 *   - Kembali ke tab Google Sheet Anda, refresh halamannya.
 *   - Klik menu baru "🤖 OmniBot Menu" di deretan atas -> klik "🚀 Jalankan Pengiriman Pesan".
 *   - Jika muncul pop-up "✅ PROSES SELESAI", berarti koneksi ke Script Master SUKSES!
 * 
 * STEP 7: BUKA APLIKASI OMNICHANNEL
 *   - Buka tab/jendela browser baru, lalu paste URL Web App (dari Step 4).
 *   - Jika UI Omnichannel terbuka dan sinkron, Instalasi Anda telah SUKSES 100%!
 *   - (Langkah Akhir: Pasang Trigger Waktu 1 menit ke fungsi `actionKirimPesan` di menu Triggers).
 * =========================================================================
 */
