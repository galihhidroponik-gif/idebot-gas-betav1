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

// =========================================================================
// VARIABEL SISTEM (JANGAN DIUBAH)
// =========================================================================
var CLIENT_SHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
var CLIENT_SHEET_URL = SpreadsheetApp.getActiveSpreadsheet().getUrl();

// =========================================================
// [STEP 3] FUNGSI PANCINGAN OTORISASI (Jalankan Manual)
// =========================================================
function otorisasiSistem() {
  var email = Session.getEffectiveUser().getEmail();
  Logger.log("Otorisasi berhasil untuk: " + email);
}

// =========================================================================
// [STEP 6] MENU KUSTOM DI SPREADSHEET (TOMBOL UI)
// =========================================================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🤖 OmniBot Menu')
      .addItem('🚀 Jalankan Pengiriman Pesan', 'manualKirimPesan')
      .addToUi();
}

function manualKirimPesan() {
  var ui = SpreadsheetApp.getUi();
  
  // PENGAMAN: Cek apakah library sudah terpasang
  if (typeof OmniBot === 'undefined') {
    ui.alert('❌ LIBRARY BELUM TERPASANG', 'Anda belum memiliki IDs Library OmniBot.\n\nSilakan hubungi Admin Idekreatifa di WA 089509063690 untuk mendapatkan akses dan panduan.', ui.ButtonSet.OK);
    return; // Hentikan proses
  }

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast("Memproses pengiriman pesan...", "OmniBot", 3);
    
    // Memanggil fungsi eksekusi dari library master
    OmniBot.actionKirimPesan(CLIENT_SHEET_ID); 
    
    ui.alert('✅ PROSES SELESAI', 'Perintah pengiriman pesan otomatis telah dieksekusi oleh sistem Master tanpa masalah fatal.', ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('❌ PROSES GAGAL', 'Terjadi kesalahan saat memproses koneksi ke Master:\n\n' + error.message, ui.ButtonSet.OK);
  }
}

// =========================================================================
// CORE SYSTEM (ROUTER & TRIGGER)
// =========================================================================

// Webhook Receiver (Dari Provider WA)
function doPost(e) {
  if (typeof OmniBot === 'undefined') return ContentService.createTextOutput("LIBRARY_MISSING");
  return OmniBot.doPostLibrary(e, CLIENT_SHEET_ID);
}

// Web App UI Router (Menampilkan HTML)
function doGet(e) {
  if (typeof OmniBot === 'undefined') return HtmlService.createHtmlOutput("<h2>❌ ERROR: LIBRARY BELUM TERPASANG</h2><p>Silakan hubungi Admin Idekreatifa di WA 089509063690.</p>");
  var clientGasUrl = ScriptApp.getService().getUrl();
  return OmniBot.doGetLibrary(e, CLIENT_SHEET_URL, clientGasUrl);
}

// Trigger Pengiriman Pesan Otomatis
function actionKirimPesan() {
  // PENGAMAN: Cek apakah library sudah terpasang untuk trigger/manual run Apps Script
  if (typeof OmniBot === 'undefined') {
    var errorMsg = "❌ LIBRARY BELUM TERPASANG: Anda belum memiliki IDs Library OmniBot. Silakan hubungi admin idekreatifa 089509063690";
    Logger.log(errorMsg);
    throw new Error(errorMsg); // Memunculkan alert merah (Exception) di Apps Script
  }

  return OmniBot.actionKirimPesan(CLIENT_SHEET_ID); 
}

function cekDanIsiOtomatisB2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetSetting = ss.getSheetByName("Setting");
  
  if (sheetSetting) {
    var cellB2 = sheetSetting.getRange("B2");
    var nilaiB2 = cellB2.getValue();
    
    // Jika sel B2 kosong, otomatis isi dengan ID spreadsheet aktif saat ini!
    if (!nilaiB2 || nilaiB2.toString().trim() === "") {
      var currentId = ss.getId();
      cellB2.setValue(currentId);
      Logger.log("ID Spreadsheet berhasil diisi otomatis ke B2: " + currentId);
    }
  }
}

// =========================================================================
// JEMBATAN PENGHUBUNG UI HTML (JALUR DATA BACKEND)
// =========================================================================
function getSheetData(sheetInput, userAccountsStr) { 
  if (typeof OmniBot === 'undefined') return null;
  return OmniBot.getSheetData(sheetInput, userAccountsStr); 
}

function getAvailableLabels(sheetInput) { 
  if (typeof OmniBot === 'undefined') return null;
  return OmniBot.getAvailableLabels(sheetInput); 
}

function updateContactData(sheetInput, uniqueId, newName, newLabels, newBiodata) { 
  if (typeof OmniBot === 'undefined') return "error: Library Missing";
  return OmniBot.updateContactData(sheetInput, uniqueId, newName, newLabels, newBiodata); 
}

function markAsRead(sheetInput, uniqueId) { 
  if (typeof OmniBot === 'undefined') return "error: Library Missing";
  return OmniBot.markAsRead(sheetInput, uniqueId); 
}

function processUIPostNative(data) { 
  if (typeof OmniBot === 'undefined') return JSON.stringify({status: "error", message: "Library Missing"});
  return OmniBot.processUIPostNative(data); 
}

// Mengambil informasi Token AI dari Sheet Setting J1
function getAITokenStatus(sheetUrl) {
  try {
    const ss = SpreadsheetApp.openByUrl(sheetUrl);
    const sheet = ss.getSheetByName("Setting");
    if (sheet) {
      // Mengambil nilai text dari sel J1
      const tokenInfo = sheet.getRange("J1").getValue(); 
      return tokenInfo;
    }
    return "Sheet Setting tidak ditemukan";
  } catch (e) {
    return "Error memuat token";
  }
}

// TAMBAHAN JEMBATAN CONTROL PANEL UI & SERVER SEARCH
function getSettingsForUI(sheetInput) { return OmniBot.getSettingsForUI(sheetInput); }

// PERUBAHAN: Selaraskan parameter menjadi h22 dan h20
function saveGeneralSettings(sheetInput, h22, h20) { return OmniBot.saveGeneralSettings(sheetInput, h22, h20); }

// JEMBATAN BARU (Untuk Edit & Tambah)
function saveDeviceDataNative(sheetInput, index, dataObj) { return OmniBot.saveDeviceDataNative(sheetInput, index, dataObj); }
function deleteDeviceFromUI(sheetInput, index) { return OmniBot.deleteDeviceFromUI(sheetInput, index); }
function searchContactNative(sheetInput, keyword) { return OmniBot.searchContactNative(sheetInput, keyword); }

// =========================================================================
// [JEMBATAN ROUTER UI FINANCE]
// =========================================================================
// =========================================================================
// [MODUL CLIENT-SIDE SELF-HEALING] - CEK & BUAT SHEET FINANCE OTOMATIS
// =========================================================================
function checkAndCreateFinanceSheet_(sheetUrl) {
  try {
    var ss = sheetUrl ? SpreadsheetApp.openByUrl(sheetUrl) : SpreadsheetApp.getActiveSpreadsheet();
    var sheet = null;
    var sheets = ss.getSheets();
    
    // Cari tab sheet bernama "Finance" secara case-insensitive
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getName().trim().toLowerCase() === "finance") {
        sheet = sheets[i];
        break;
      }
    }
    
    // Jika tab Finance tidak ada, buat secara otomatis dengan 22 Header baku
    if (!sheet) {
      sheet = ss.insertSheet("Finance");
      sheet.appendRow([
        "Timestaps", "Akun", "Agen", "ID Networking", "Nama Konsumen", 
        "Status", "Jenis", "Kategori", "Tgl Transaksi", "Nama Item", 
        "Qty", "Satuan", "Harga Satuan", "Total", "Note", 
        "Id Group", "Nama Group", "No Inv", "Link Drive", "Nama File", "Proses AI", "Bulan"
      ]);
      Logger.log("✅ Tab 'Finance' berhasil dibuat otomatis di spreadsheet klien.");
    }
    return sheet;
  } catch (e) {
    Logger.log("❌ Gagal inisialisasi sheet Finance: " + e.message);
    return null;
  }
}

// JEMBATAN ROUTER UI FINANCE DENGAN AUTO-HEALING KLIEN
function getFinanceDataMaster(sheetUrl) { 
  if (typeof OmniBot === 'undefined') return [];
  
  // Lakukan pengecekan dan pembuatan sheet otomatis di sisi klien sebelum ditarik master
  checkAndCreateFinanceSheet_(sheetUrl);
  
  return OmniBot.getFinanceDataMaster(sheetUrl); 
}

function updateFinanceRowStatus(sheetUrl, rowIndex, newStatus) { 
  if (typeof OmniBot === 'undefined') return "error: Library Missing";
  
  // Pastikan sheet ada sebelum update status
  checkAndCreateFinanceSheet_(sheetUrl);
  
  return OmniBot.updateFinanceRowStatus(sheetUrl, rowIndex, newStatus); 
}

// =========================================================================
// [BARU] FUNGSI HAPUS PROFIL KONTAK (EKSEKUSI NATIVE)
// =========================================================================


// =========================================================================
// [BARU] FUNGSI HAPUS PROFIL KONTAK (EKSEKUSI NATIVE)
// =========================================================================
function deleteContactData(sheetUrl, contactId) {
  try {
    var ss = SpreadsheetApp.openByUrl(sheetUrl);
    // CATATAN: Pastikan nama sheet Anda benar "Database" (sesuaikan jika berbeda)
    var sheet = ss.getSheetByName("Networking"); 
    
    if (!sheet) {
      return "error: Sheet Database tidak ditemukan";
    }
    
    var data = sheet.getDataRange().getValues();
    
    // Cari baris berdasarkan uniqueId (Asumsi ID berada di Kolom B / Array Index 1)
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() === String(contactId).trim()) {
        // Hapus baris (i + 1 karena array dimulai dari 0, baris sheet dari 1)
        sheet.deleteRow(i + 1);
        return "success";
      }
    }
    
    return "error_not_found";
  } catch (e) {
    return "error: " + e.message;
  }
}

function verifyLoginNative(sheetInput, user, pass) { 
  return OmniBot.verifyLogin(sheetInput, user, pass); 
}

// [BARU] Trigger Harian untuk Kalkulasi CRM & RFM
function runDailyCRMSync() {
 // PENGAMAN: Cek apakah library sudah terpasang
 if (typeof OmniBot === 'undefined') {
   var errorMsg = "❌ LIBRARY BELUM TERPASANG: Anda belum memiliki IDs Library OmniBot. Silakan hubungi admin idekreatifa 089509063690";
   Logger.log(errorMsg);
   throw new Error(errorMsg);
 }
 // Mengeksekusi fungsi CRM dari Library Master dengan mengirimkan ID Sheet Klien
 return OmniBot.dailyCRMSync(CLIENT_SHEET_ID);
}
