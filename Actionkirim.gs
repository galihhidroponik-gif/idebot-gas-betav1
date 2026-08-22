/**
 * Kirim pesan WhatsApp dari Google Spreadsheet ke Starsender
 * Versi: Full Log Eksekusi, Single Call Optimized & AI Queue Engine
 * Update: Global History Model Reset (H16 & H17)
 * Lookup Laporan: Kirim Kolom A <--> Networking Kolom B
 * Lookup Provider: Kirim Kolom H <--> Setting Kolom A
 */

function actionKirimPesan(clientSheetId) { 
  // 1. Inisialisasi Lock Service
  var lock = LockService.getDocumentLock();
  
  // 2. Coba dapatkan kunci. Tunggu maksimal 3 detik.
  // Jika ada skrip yang masih berjalan, skrip baru ini akan berhenti.
  if (!lock.tryLock(3000)) {
    Logger.log("⚠️ Skrip sebelumnya masih berjalan. Eksekusi dibatalkan untuk mencegah tumpang tindih.");
    return; // Keluar dari fungsi
  }

  try {
  // 2. Gunakan ID klien untuk membuka Spreadsheet (Aman dari Trigger Error)
  var ss = clientSheetId ? SpreadsheetApp.openById(clientSheetId) : SpreadsheetApp.getActiveSpreadsheet();
  var tabSetting = ss.getSheetByName("Setting");
  var tabNetworking = ss.getSheetByName("Networking"); 
  var tabSending = ss.getSheetByName("Send");

  if (!tabSending || !tabSetting || !tabNetworking) {
    Logger.log("❌ EROR: Salah satu Sheet (Setting/Networking/Send) tidak ditemukan!");
    return;
  }

  Logger.log("📥 Mengambil seluruh data awal (Single Call Matrix)...");

  // =========================================================================
  // 1. SINGLE CALL: SHEET SETTING
  // =========================================================================
  var lastRowSetting = tabSetting.getLastRow();
  var settingMatrix = lastRowSetting > 0 ? tabSetting.getRange(1, 1, lastRowSetting, 17).getValues() : [];
  Logger.log("✅ Data Setting diambil: " + settingMatrix.length + " baris.");
  
  var row_mulai = 2;
  var jumlahPengiriman = 10;
  var waktuKirimB18 = "";
  var fufData = [];
  var providerMap = {};
  
var globalAI = {
    tempGemini: 0.2, geminiApiKey: "", geminiModel: "",
    promptWarmerName: "", maxOutputTokens: 2048, 
    idtglUpdatemodel: "", historymodel: "" 
  };

  if (settingMatrix.length > 0) {
    row_mulai = toAngkaInt(settingMatrix[15] ? settingMatrix[15][1] : 2) || 2;       // B16
    jumlahPengiriman = toAngkaInt(settingMatrix[16] ? settingMatrix[16][1] : 10);    // B17
    waktuKirimB18 = safeTrim(settingMatrix[17] ? settingMatrix[17][1] : "");         // B18

    // Mapping Config AI dari Webhook
    globalAI.tempGemini       = settingMatrix[15] ? (settingMatrix[15][5] || 0.2) : 0.2; // F16
    globalAI.geminiApiKey     = settingMatrix[17] ? (settingMatrix[17][5] || "") : ""; // F18
    globalAI.geminiModel      = settingMatrix[17] ? (settingMatrix[17][6] || "") : ""; // G18
    globalAI.promptWarmerName = settingMatrix[19] ? (settingMatrix[19][5] || "") : ""; // F20
    globalAI.maxOutputTokens  = settingMatrix[15] ? (settingMatrix[15][4] || 4096) : 4096; // Pindah ke E16
    globalAI.fallbackModels   = settingMatrix[18] ? (settingMatrix[18][7] || "") : ""; // Pindah ke H19
    
    // Mapping Global History
    globalAI.idtglUpdatemodel = settingMatrix[15] ? safeTrim(settingMatrix[15][7]) : ""; // H16
    globalAI.historymodel     = settingMatrix[15] ? safeTrim(settingMatrix[15][6]) : ""; // Pindah ke G16

    // Mapping FUF Data (A23:B30 -> Index 22 s/d 29) >> diubah ke pembacaaan personal dari kolom AF (Sheet networking)
    // HAPUS ATAU BERI KOMENTAR KODE INI:
    // for (var i = 22; i <= 29; i++) {
    //  if (settingMatrix[i]) fufData.push([settingMatrix[i][0], settingMatrix[i][1]]);
    // }

    // Mapping Provider (A34:Q -> Index 33 sampai baris terakhir)
    for (var i = 33; i < settingMatrix.length; i++) {
      var r = settingMatrix[i];
      var key = safeTrim(r[0]);
      if (key) {
        providerMap[key] = {
          settingRow: i + 1, // <--- TAMBAHAN: Menyimpan nomor baris posisi device di Sheet Setting
          csName: safeTrim(r[1]), 
          status: safeTrim(r[5]),
          apiKey: safeTrim(r[8]),
          apiUrl: safeTrim(r[9]),
          prefixUrl: safeTrim(r[10]),
          devicePrompt: safeTrim(r[12]), 
          delayjadwalKirim: safeTrim(r[14])
        };
      }
    }
    Logger.log("✅ Ditemukan " + Object.keys(providerMap).length + " provider/device di Sheet Setting.");
  }

  // =========================================================================
  // 1.1. LOGIKA RESET GLOBAL HISTORY MODEL (H16 & H17)
  // =========================================================================
  // Generate tanggal hari ini dalam format 'idYYMMDD'
  var todayDate = new Date();
  var tz = Session.getScriptTimeZone();
  var todayFormat = "id" + Utilities.formatDate(todayDate, tz, "yyMMdd");

  if (globalAI.idtglUpdatemodel !== todayFormat) {
    Logger.log("🔄 Beda hari terdeteksi. Mereset history model global (H16 & H17).");
    globalAI.idtglUpdatemodel = todayFormat;
    globalAI.historymodel = ""; // Reset memory
    try {
      // Update Fisik ke Master Setting (H16 = Tanggal Baru, G16 = Kosong)
      tabSetting.getRange("H16").setValue(todayFormat);
      tabSetting.getRange("G16").setValue("");
    } catch(e) { 
      Logger.log("❌ Error reset H16 & G16 - " + e.message); 
    }
  }

  // =========================================================================
  // 2. SINGLE CALL: SHEET Networking
  // =========================================================================
  var lastRowNetworking = tabNetworking.getLastRow();
  var colRangeNetworking = Math.max(tabNetworking.getLastColumn(), 33); 
  var NetworkingMatrix = lastRowNetworking > 0 ? tabNetworking.getRange(1, 1, lastRowNetworking, colRangeNetworking).getDisplayValues() : [];
  var NetworkingMap = {};
  for (var i = 0; i < NetworkingMatrix.length; i++) {
    var key = safeTrim(NetworkingMatrix[i][1]); 
    if (key) NetworkingMap[key] = i + 1; 
  }
  Logger.log("✅ Data Networking dimapping: " + Object.keys(NetworkingMap).length + " report ID ditemukan.");

  // =========================================================================
  // 3. SINGLE CALL: SHEET SEND 
  // =========================================================================
  // Ambil batas row maksimal dari Setting B16 (Default ke last row jika kosong/0)
  var maxBatasRow = toAngkaInt(settingMatrix[15] ? settingMatrix[15][1] : 0); 
  var lastRowSending = tabSending.getLastRow();
  
  // Jika Setting B16 diisi, maka kunci batas pembacaan sesuai nilai B16 tersebut
  if (maxBatasRow > 0) {
    Logger.log("⚙️ Batas row maksimal dikunci berdasarkan Setting B16: hingga baris " + maxBatasRow);
    lastRowSending = maxBatasRow;
  }

  // Sesuai instruksi: Pembacaan antrean selalu dimulai secara mutlak dari baris 2
  var row_mulai_send = 2; 
  var sendingData = [];
  
  if (lastRowSending >= row_mulai_send) {
    // Rumus baris: (Batas Row - Baris Mulai + 1)
    var totalBarisDitarik = lastRowSending - row_mulai_send + 1;
    sendingData = tabSending.getRange(row_mulai_send, 1, totalBarisDitarik, tabSending.getLastColumn()).getDisplayValues();
  }
  Logger.log("✅ Data Send siap dieksekusi dari baris " + row_mulai_send + " sampai baris " + lastRowSending + ". Total baris antrean: " + sendingData.length);

  if (sendingData.length === 0) {
    Logger.log("⚠️ Tidak ada data untuk dikirim di Sheet Send. Proses dihentikan.");
    return;
  }

  Logger.log("▶️ MEMULAI PROSES PENGIRIMAN & AI ENGINE...");

  // =========================================================================
  // 0. BATCH UPDATE MANAGER (PENGUMPUL SETVALUE)
  // =========================================================================
  var BatchManager = {
    netWrites: [], setWrites: [], 
    colorNetGreen: [], colorNetRed: [], colorNetDarkRed: [], colorSetRed: [],

    // Fungsi untuk menabung data yang akan ditulis (harus array 2D)
    pushNet: function(r, c, vArray) { this.netWrites.push({row: r, col: c, vals: vArray}); },
    pushSet: function(r, c, vArray) { this.setWrites.push({row: r, col: c, vals: vArray}); },

    // Fungsi eksekusi massal di akhir kode
    commit: function(tabNet, tabSet) {
      Logger.log("💾 Menyimpan batch update ke Spreadsheet...");
      
      // 1. Tulis massal Data Networking
      for(var i = 0; i < this.netWrites.length; i++) {
         var w = this.netWrites[i];
         tabNet.getRange(w.row, w.col, 1, w.vals[0].length).setValues(w.vals);
      }
      
      // 2. Tulis massal Data Setting
      for(var i = 0; i < this.setWrites.length; i++) {
         var w = this.setWrites[i];
         tabSet.getRange(w.row, w.col, 1, w.vals[0].length).setValues(w.vals);
      }
      
      // 3. Tulis massal Warna menggunakan RangeList (Sangat Cepat)
      if(this.colorNetGreen.length > 0) tabNet.getRangeList(this.colorNetGreen).setBackground("#008000").setFontColor("#FFFFFF");
      if(this.colorNetRed.length > 0) tabNet.getRangeList(this.colorNetRed).setBackground("#FF0000").setFontColor("#FFFFFF");
      if(this.colorNetDarkRed.length > 0) tabNet.getRangeList(this.colorNetDarkRed).setBackground("#8B0000").setFontColor("#FFFFFF");
      if(this.colorSetRed.length > 0) tabSet.getRangeList(this.colorSetRed).setBackground("#FF0000").setFontColor("#FFFFFF");
    }
  };

  // --- 4. SISTEM BATCHING + AI LOGGING ---
  var StarSender = {
    kloterPengiriman: (jumlahPengiriman > 0 ? jumlahPengiriman : 10),
    cacheMessage: [], NetworkingRows: [], apiKeys: [], apiUrls: [], aiLogs: [],
    
    pushMessage: function(msg, targetRow, apiKey, apiUrl, aiLogData) {
      this.cacheMessage.push(JSON.parse(JSON.stringify(msg)));
      this.NetworkingRows.push(targetRow);
      this.apiKeys.push(apiKey);
      this.apiUrls.push(apiUrl);
      this.aiLogs.push(aiLogData || ["", "", "", "", ""]); 
      Logger.log("📦 Pesan baris " + targetRow + " masuk ke antrean Batching (" + this.cacheMessage.length + "/" + this.kloterPengiriman + ").");
      if (this.cacheMessage.length >= this.kloterPengiriman) this.flush(tabNetworking);
    },
    
    // --- BAGIAN DALAM StarSender.flush YANG SUDAH DIMODIFIKASI ---

    flush: function(sheetNetworking) {
      if (this.cacheMessage.length === 0) return;
      Logger.log("🚀 MENGIRIM BATCH (" + this.cacheMessage.length + " pesan) ke API secara PARALEL...");
      
      var now = new Date();
      var waktuBerhasil = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
      
      // 1. Kumpulkan semua request ke dalam array
      var requests = [];
      for (var i = 0; i < this.cacheMessage.length; i++) {
        var authHeader = this.apiKeys[i];
        
        // Deteksi jika OneSender dan belum ada kata 'Bearer ', tambahkan 'Bearer '
        if (this.apiUrls[i].toLowerCase().indexOf("starsender") === -1 && !authHeader.startsWith("Bearer ")) {
          authHeader = "Bearer " + authHeader;
        }

        requests.push({
          url: this.apiUrls[i],
          method: "post",
          headers: { "Authorization": authHeader, "Content-Type": "application/json" },
          payload: JSON.stringify(this.cacheMessage[i]),
          muteHttpExceptions: true
        });
      }

      // 2. Tembak semua API SEKALI JALAN (Sangat Cepat!)
      var responses = UrlFetchApp.fetchAll(requests);

      // 3. Proses hasilnya
      for (var i = 0; i < responses.length; i++) {
        var targetRow = this.NetworkingRows[i]; 
        var logAI = this.aiLogs[i];
        
        try {
          var responseCode = responses[i].getResponseCode();
          var responseText = responses[i].getContentText();
          
          if (responseCode === 200 || responseCode === 201) {
             Logger.log("✅ [FETCH SUCCESS] HTTP " + responseCode + " | Row " + targetRow);
          } else {
             Logger.log("❌ [FETCH FAILED] HTTP " + responseCode + " | Row " + targetRow + " | Alasan: " + responseText);
          }

          BatchManager.pushNet(targetRow, 17, [[waktuBerhasil]]);

          if (logAI[0] !== "") {
            BatchManager.pushNet(targetRow, 22, [["AI Reply", logAI[1], logAI[2], logAI[3]]]);
          } else {
            BatchManager.pushNet(targetRow, 22, [["AI Reply"]]);
          }
          BatchManager.colorNetGreen.push("V" + targetRow);

        } catch (e) {
          BatchManager.pushNet(targetRow, 22, [["Error"]]);
          BatchManager.colorNetRed.push("V" + targetRow);
        }
      }
      
      this.cacheMessage = []; this.NetworkingRows = []; this.apiKeys = []; this.apiUrls = []; this.aiLogs = [];
      Logger.log("🧹 Batching di-reset.");
    }
  };

  // --- 5. LOOP DATA SHEET KIRIM ---
  for (var i = 0; i < sendingData.length; i++) {
    var realRow = row_mulai_send + i;
    var rowArr = sendingData[i];
    
    var reportID = safeTrim(rowArr[0]); // Kolom A
    Logger.log("--------------------------------------------------");
    Logger.log("🔍 Mengecek Sheet Send Baris: " + realRow + " | Report ID: " + reportID);

    if (reportID === "") {
        Logger.log("⚠️ Lewati baris " + realRow + ": Report ID kosong.");
        continue;
    }

    var statusQueueMode = safeTrim(rowArr[2]).toLowerCase(); // Kolom C
    var historyMsgData = safeTrim(rowArr[3]); // Kolom D
    var providerID = safeTrim(rowArr[7]); // Kolom H
    var nilaiBaru = toAngkaInt(rowArr[8]); // Kolom I
    
    // ========================================================================
    // MODIFIKASI: Deteksi Grup Lebih Awal dari Kolom A untuk Kebutuhan Bypass
    // ========================================================================
    var isGroupMsg = reportID.toLowerCase().indexOf("@g.us") !== -1;

    var targetNetworkingRow = NetworkingMap[reportID];
    if (!targetNetworkingRow) {
        Logger.log("⏭️ Skip baris " + realRow + ": Report ID '" + reportID + "' TIDAK ADA di Sheet Networking.");
        continue;
    }

    try {
      var provider = providerMap[providerID];
      var statusDiijinkan = ["run", "aer run"];
      var statusSekarang = safeTrim(provider ? provider.status : "").toLowerCase();

      if (!provider) {
          Logger.log("⏭️ Skip baris " + realRow + ": Provider ID '" + providerID + "' tidak ditemukan di Setting.");
          continue;
      }

      // ========================================================================
      // MODIFIKASI: Logika Bypass "Stop" (Lolos jika status "Queue" ATAU "Grup")
      // ========================================================================
      var isBypassState = (statusQueueMode === "queue" || isGroupMsg);

      if (!statusDiijinkan.includes(statusSekarang) && !isBypassState) {
          Logger.log("⏭️ Skip baris " + realRow + ": Status Provider '" + providerID + "' adalah '" + statusSekarang + "' dan bukan Queue / Grup.");
          continue;
      }
      
      if (!statusDiijinkan.includes(statusSekarang) && isBypassState) {
          var alasanBypassState = (statusQueueMode === "queue") ? "Queue" : "Target Grup (@g.us)";
          Logger.log("⚠️ Bypass Status: Provider '" + providerID + "' sedang '" + statusSekarang + "', tapi baris " + realRow + " adalah " + alasanBypassState + ". Tetap diproses.");
      }

      if (waktuKirimB18 !== "") {
        if (!isDalamJamKerja(waktuKirimB18)) {
            if (isBypassState) {
                var alasanBypassJam = (statusQueueMode === "queue") ? "Queue" : "Target Grup (@g.us)";
                Logger.log("⚠️ Bypass Jam Kerja: Baris " + realRow + " adalah " + alasanBypassJam + ". Tetap diproses meski di luar jam kerja.");
            } else {
                Logger.log("⏸️ PENGIRIMAN DI-SKIP: Baris " + realRow + " bukan Queue/Grup dan di luar jam kerja (" + waktuKirimB18 + ").");
                continue;
            }
        }
      }

      if (provider.delayjadwalKirim) {
        var waktuSekarang = new Date();
        var waktuDelay = new Date(provider.delayjadwalKirim);
        if (!isNaN(waktuDelay.getTime()) && waktuSekarang < waktuDelay) {
            Logger.log("⏭️ Skip baris " + realRow + ": Menunggu delay jadwal kirim Provider (" + waktuDelay + ").");
            continue; 
        }
      }

      // --- EKTRAKSI NOMOR TELEPON & DETEKSI GRUP ---
      // (Kode selanjutnya tetap sama seperti sebelumnya...)
      var recipientRaw = safeTrim(rowArr[4]); // Kolom E

      // --- EKTRAKSI NOMOR TELEPON & DETEKSI GRUP ---
      var recipientRaw = safeTrim(rowArr[4]); // Kolom E
      var isGroup = recipientRaw.toLowerCase().indexOf("@g.us") !== -1;
      var recipientDigits = "";

      if (isGroup) {
        // Jika format grup (@g.us), gunakan string asli tanpa diubah
        recipientDigits = recipientRaw;
      } else {
        // Jika nomor HP biasa, bersihkan dan beri prefix 62
        var digits = (recipientRaw + "").match(/\d+/g);
        if (!digits) {
            Logger.log("⏭️ Skip baris " + realRow + ": Tidak ada angka/nomor telepon valid di kolom E (" + recipientRaw + ").");
            continue;
        }
        recipientDigits = autoPrefixPhone(digits.join(''));
      }
      
      // HAPUS DUPLIKASI BARIS INI:
      // var recipientDigits = autoPrefixPhone(digits.join('')); 

      var isStarsender = provider.apiUrl.toLowerCase().indexOf("starsender") !== -1;
      Logger.log("👀 DEBUG URL API Provider (" + providerID + "): [" + provider.apiUrl + "]");
      
      // ========================================================================
      // [MODIFIKASI] LOGIKA BYPASS PENGECEKAN WA (KOLOM B BARU & KOLOM C == QUEUE)
      // ========================================================================
      var skipCekWA = false; // <--- TAMBAHAN WAJIB
      var alasanBypass = ""; // <--- TAMBAHAN WAJIB
      
      var lastActiveRaw = rowArr[1]; // Kolom B
      // Tambahkan kondisi bypass jika penerima adalah Grup WhatsApp
      if (isGroup) {
          skipCekWA = true;
      }

      // 1. Cek Syarat Bypass Kondisi A: Jika Kolom C bernilai "queue"
      if (statusQueueMode === "queue") {
          skipCekWA = true;
          alasanBypass = "Kolom C berstatus 'queue'.";
      } 
      // 2. Cek Syarat Bypass Kondisi B: Jika Kolom B aktif baru-baru ini (Hari ini, Kemarin, Lusa Kemarin)
      else if (lastActiveRaw && lastActiveRaw !== "") {
          var lastActiveDate = null;
          var parts = lastActiveRaw.toString().trim().split(/[\/\-]/);
          
          if (parts.length === 3) {
              var p0 = parseInt(parts[0], 10); // Bulan (06)
              var p1 = parseInt(parts[1], 10); // Tanggal (12)
              var p2 = parseInt(parts[2], 10); // Tahun (2026)
              
              // Cek Cerdas: Format MM/dd/yyyy (06/12/2026 -> p0=6 Juni, p1=12 Tanggal)
              if (p0 <= 12 && p1 <= 31) {
                  lastActiveDate = new Date(p2, p0 - 1, p1);
              } else if (p0 > 12) { // Antisipasi format dd/MM/yyyy jika sewaktu-waktu berubah
                  lastActiveDate = new Date(p2, p1 - 1, p0);
              } else {
                  lastActiveDate = new Date(lastActiveRaw);
              }
          } else {
              lastActiveDate = new Date(lastActiveRaw);
          }

          if (lastActiveDate && !isNaN(lastActiveDate.getTime())) {
              var today = new Date();
              
              // Reset waktu ke jam 00:00:00 untuk perbandingan murni tanggal hari
              today.setHours(0, 0, 0, 0);
              lastActiveDate.setHours(0, 0, 0, 0);
              
              // Menggunakan Math.round agar terhindar dari selisih desimal jam akibat timezone script
              var diffTime = today.getTime() - lastActiveDate.getTime();
              var diffDays = Math.round(diffTime / (1000 * 60 * 60 * 24)); 
              
              // Kebijakan: 0 = Hari ini, 1 = Kemarin, 2 = Lusa Kemarin
              if (diffDays >= 0 && diffDays < 3) { 
                  skipCekWA = true;
                  alasanBypass = "Nomor tercatat aktif di Kolom B (" + diffDays + " hari yang lalu).";
              }
          }
      }

      var hasilCek = true;
      if (skipCekWA) {
          // Log Bypass dicetak rapi di sini
          Logger.log("⏩ [Cek WA] Bypass pengecekan API. Alasan: " + alasanBypass);
          hasilCek = true; // Langsung lolos tanpa tembak API WhatsApp
      } else {
          // Jika gagal bypass, baru sistem menguji nomor via API resmi
          Logger.log("🔎 [Cek WA] Menguji nomor " + recipientDigits + " (Is Starsender: " + isStarsender + ")...");
          if (isStarsender) {
              hasilCek = cekNomorStarsender(recipientDigits, provider.apiKey);
          } else {
              hasilCek = cekNomorOneSender(recipientDigits, provider.apiKey, provider.prefixUrl);
          }
      }
      // ========================================================================
      
      if (hasilCek === "DEVICE_ERROR") {
        // PENANGANAN JIKA DEVICE ADMIN MATI / ERROR API
        Logger.log("⚠️ [Cek WA] Skip baris " + realRow + ": Koneksi Device Admin terputus. Menulis status ke Setting.");

        // HANYA UPDATE DI SHEET SETTING: Tulis "Device OFF" di Kolom C (Indeks 3)
        if (provider.settingRow) {
            BatchManager.pushSet(provider.settingRow, 3, [["Device OFF"]]);
            BatchManager.colorSetRed.push("C" + provider.settingRow);
        }
        
        continue; // Lanjut ke antrean berikutnya. Kolom V dan C di Networking tidak disentuh.
        
      } else if (hasilCek === false) {
        // PENANGANAN JIKA NOMOR CUSTOMER BENAR-BENAR MATI
        Logger.log("⏭️ [Cek WA] Skip baris " + realRow + ": Nomor " + recipientDigits + " TIDAK AKTIF di WA.");
        
        // 1. Bersihkan value lama dan ubah kolom C (Indeks 3) menjadi 'XX'
        BatchManager.pushNet(targetNetworkingRow, 3, [["XX"]]);
        
        // 2. Ubah kolom V (Indeks 22) menjadi 'Tidak Aktif'
        BatchManager.pushNet(targetNetworkingRow, 22, [["Tidak Aktif"]]);
        BatchManager.colorNetDarkRed.push("V" + targetNetworkingRow);
        
        continue; 
      } else {
        Logger.log("✅ [Cek WA] Nomor " + recipientDigits + " valid dan aktif.");
      }

      var messageType = safeTrim(rowArr[10]); 
      var mediaLink = safeTrim(rowArr[12]);
      var messageBuilt = "";
      var aiLogDataUpdate = ["", "", "", ""]; // Setup log kosong (V,W,X,Y)

      // =========================================================================
      // 🎯 EKSEKUSI AI JIKA STATUS KOLOM C == "QUEUE" (HANYA TEKS)
      // =========================================================================
      if (statusQueueMode === "queue" && (!mediaLink || mediaLink === "")) {
        Logger.log("🤖 [AI MODE] Memulai proses AI untuk baris " + realRow);
        
        var rawCustName = safeTrim(rowArr[5]); // Simpan nama asli untuk dicek
        var custNameAI = rawCustName || "Kakak";
        var labelKirimAI = safeTrim(rowArr[6]).toUpperCase();

        // --- LOGIKA DINAMIS ELEMEN KE-3 (NAMA vs RESUME) ---
        var instruksiElemen3 = "";
        var labelElemen3 = "";
        
        if (!rawCustName || rawCustName.includes("CEK_")) {
            labelElemen3 = "Nama";
            instruksiElemen3 = "3. Nama Sapaan Akrab : diisi nama sapaan costumer dari history percakapan beserta sapaan akrab yang disukai costumer misalnya bang, kak, pak atau bu contoh pak galih, kak selly, bunda fatma dan lain lain\n";
        } else {
            labelElemen3 = "Resume";
            instruksiElemen3 = "3. resumeAI : meresume konteks percakapan dari history pesan\n";
        }
        
        // 1. Ubah string multi-label menjadi Array agar pencarian lebih akurat
        var arrayLabelAktif = labelKirimAI.split(",").map(function(s) { return s.trim(); });
        var isModeWarmerAktif = arrayLabelAktif.includes("AE") || arrayLabelAktif.includes("AER");
        var isModeKaryawanAktif = arrayLabelAktif.includes("ON");

        custNameAI = custNameAI.replace(/CEK_/gi, "Kak ");
        
        // --- SELALU GUNAKAN GEMINI ---
        var modelTerpakai = globalAI.geminiModel;

        var saPaan = getPanggilanDepan(custNameAI);
        var statusNama = "";

      if (!rawCustName || rawCustName.includes("CEK_")) { // Jika mengandung CEK_ atau Kosong, arahkan AI cari di history atau sapaan default
      namaUntukAI = "USER";
      let fullNama = ("")
      .replace(/nama\s*:?\s*/i,"")
      .trim();
      saPaan = (fullNama.split(/\s+/).length > 1
     ? getPanggilanDepan(fullNama)
     : "") || "Kak";

      // Buat variabel pengecekan yang aman di luar teks
      let namaDariHistory = safeTrim(rowArr[23]);
      statusNama =
      "NAMA COSTUMER BELUM TERVERIFIKASI, TERCATAT SAAT INI : "+ ( custNameAI || "kakak") + " atau " + namaDariHistory + 
      "\n- VALIDASI SEBELUM MENGIRIM: Jika nama disebut tanpa sapaan → REGENERATE JAWABAN gunakan default Kak " +
      "\n- Pancing pertanyaan untuk menggali informasi sapaan akrab yang disukai costumer : 'Izin supaya lebih akrab bisa kakak admin panggil apa ya Pak/Bu/Bang/Kak?\n"+
      "- Jika costumer menyebutkan sapaan akrab + nama maka simpan dan LOCK gunakan sesuai preferensi customer." +
      " (contoh: Bapak Ahmad, Ibu Rina, Kak Fajar)"
      ;
    } else {
      // Jika sudah bersih (tanpa CEK_), artinya sudah diverif manual oleh admin
      namaUntukAI =  rawCustName;
      saPaan = getPanggilanDepan(namaUntukAI); // Ambil panggilan depan (Contoh: "bang")
      statusNama = "TerVerifikasi, Gunakan nama '" + rawCustName + "' 1 Kali per percakapan";
    }
        
        
        // --- 1. Set nilai default jika sapaan tidak ada di daftar bawah ---
        var instruksiSpesifikSapaan = "- Tipe sapaan santun dan hangat kepada " + (saPaan || custNameAI); 

        // --- 2. LOGIKA PENGGOLONGAN SAPAAN ---
        var sapaanClean = saPaan.toLowerCase(); // Pastikan semuanya jadi huruf kecil

        if (["bang", "abang"].includes(sapaanClean)) {
            instruksiSpesifikSapaan = "- Tipe sapaan AKRAB: Gunakan '" + saPaan + "' khas pelayanan Indonesia kepada customer pria.\n" +
            "- Variasikan sapaan abang boleh di depan atau belakang sedangkan bang hanya dipakai di depan nama.";
        } 
        else if (["kakak", "kak"].includes(sapaanClean)) {
            instruksiSpesifikSapaan = "- Tipe sapaan AKRAB: Gunakan '" + saPaan + "' khas pelayanan Indonesia kepada semua gender." +
            "- Variasikan sapaan kakak boleh di depan atau belakang sedangkan kak hanya dipakai di depan nama.";
        } 
        else if (["mas"].includes(sapaanClean)) { // Huruf kecil agar cocok dengan toLowerCase()
            instruksiSpesifikSapaan = "- Gunakan '" + saPaan + "' boleh di depan atau belakang nama dengan nada bicara profesional dan sedikit lebih akrab kepada laki-laki khas pelayanan Indonesia.";
        } 
        else if (["mba", "mbak"].includes(sapaanClean)) { // Huruf kecil + tambah variasi 'mbak'
            instruksiSpesifikSapaan = "- Gunakan '" + saPaan + "' boleh di depan atau belakang nama dengan nada bicara profesional dan sedikit lebih akrab khas pelayanan Indonesia kepada customer perempuan.";
        } 
        else if (["pak", "bapak"].includes(sapaanClean)) {
            instruksiSpesifikSapaan = "- Tipe sapaan FORMAL kepada laki-laki dengan nada bicara profesional dan sedikit lebih santun. Variasikan 'Bapak' boleh di depan atau belakang nama sedangkan sapaan 'Pak' untuk di depan.";
        } 
        else if (["ibu", "bu"].includes(sapaanClean)) {
            instruksiSpesifikSapaan = "- Tipe sapaan santun dan hangat. Gunakan variasi sapaan 'Ibu' boleh di depan atau belakang nama sedangkan sapaan bu selalu di depan. Gunakan '" + saPaan + "' dengan nada bicara menghormati kepada customer perempuan seperti ibu sendiri.";
        } 
        else if (["bunda", "ustadzah"].includes(sapaanClean)) {
            instruksiSpesifikSapaan = "- Tipe sapaan santun dan hangat. Gunakan '" + saPaan + "' boleh di depan atau belakang nama dengan nada bicara menghormati kepada customer perempuan seperti ibu sendiri.";
        } 
        else if (["ayah", "ustadz"].includes(sapaanClean)) {
            instruksiSpesifikSapaan = "- Tipe sapaan KELUARGA: Gunakan gaya bahasa yang hangat dan akrab. Panggilan '" + saPaan + "' sebaiknya di depan nama.";
        }
        
        // --- PERBAIKAN: MENYESUAIKAN VARIABEL DENGAN SCOPE SHEET SEND ---
        
        // 1. Cek apakah Label (arrayLabelAktif) mengandung SV
        let saveC = arrayLabelAktif.includes("SV");

        // 2. Menentukan Nama Customer (menghapus awalan CEK_ jika ada)
        let fallbackPushName = rawCustName ? rawCustName.replace(/CEK_/gi, "").trim() : "Kak";
        if (fallbackPushName === "") fallbackPushName = "Kak";
        
        // 3. Tentukan instruksi saveC
        let instruksiSaveC = saveC === false ? "- TUGAS TAMBAHAN: Sisipkan ajakan yang ramah, hangat, dan natural agar " + fallbackPushName + " menyimpan nomor kontak admin ini.\n- Berikan alasan yang masuk akal secara halus (misal: agar komunikasi layanan lebih lancar, atau agar bisa melihat update menarik di Story WhatsApp).\n- ATURAN KRITIKAL: JANGAN gunakan template kalimat yang kaku atau berulang. Rangkai kalimat ajakan dengan gaya bahasa sendiri yang bervariasi (sopan dan lebih formal) dan pastikan transisinya menyatu/nyambung dengan konteks topik obrolan saat ini.\n" : "";

        // 4. Waktu Sapaan (Kosongkan parameter agar otomatis pakai waktu sistem saat ini)
        let sapaanWaktu = getWaktuSapaan();

        // 5. Eksekusi Replace ke dalam Prompt (Langsung pakai devicePrompt standar)
        let rawDevicePrompt = provider.devicePrompt || "Anda adalah asisten virtual";
        
        // Replace [NAMAAKTIF] menggunakan custNameAI yang valid di fungsi ini
        let processedDevicePrompt = rawDevicePrompt.replace(/\[NAMAAKTIF\]/g, custNameAI);
        
        // Replace [saveC]
        processedDevicePrompt = processedDevicePrompt.replace(/\[saveC\]/g, instruksiSaveC);

        // Replace [sapaanWaktu]
        processedDevicePrompt = processedDevicePrompt.replace(/\[sapaanWaktu\]/g, sapaanWaktu);

        var historyFix = cleanHistoryLama(historyMsgData);
        var promptTerpakai = "";

        // Tentukan Format Prompt
        if (isModeWarmerAktif || isModeKaryawanAktif ) {
            Logger.log("🤖 [AI MODE] Menggunakan Prompt Warmer/AE");
            promptTerpakai = 
            (!isModeKaryawanAktif ? 
             "###DECISION TREE CHATBOT FOR WARMER WHATSAPP\n" : 
             "###DECISION TREE CHATBOT FOR REMAINDER SDM INTERNAL\n" ) +
             (!isModeKaryawanAktif ? "###[DATA PANGGILAN]\n" +
            "- PANGGILAN UTAMA USER/COSTUMER : "+ custNameAI +"\n"+
            "- " + statusNama + "\n" +
            "- Hindari kata 'anda' dan 'kamu' dalam sapaan.\n\n" +
            "###MEMORY PROCESSING HISTORY\n" +
            "- WAJIB baca seluruh riwayat percakapan sebelum menjawab.\n" +
            "- Ambil hanya INFORMASI FAKTA dari history:\n" +
            "- jawaban wajib berbeda/memiliki variasi kalimat 80% dari history pesan\n\n" +
            (globalAI.promptWarmerName || "Anda adalah asisten warmer.") 
            + "\n\n" +
            "###[STEP PROSEDUR KERJA FORMAT OUTPUT]\n" +
            "Tulisakan dalam Format Ouput 2 elemen masing masing ditulis hanya 1x wajib dengan pembatas 1 pipe (|) saja\n " +
            "Format Output: Label|Jawaban\n" +
            "1. Label: 'Balas|'\n" +
            "2. Jawaban: Kolom ini WAJIB berisi hanya teks pesan yang benar-benar akan dikirim ke WhatsApp customer. ATURAN FATAL: DILARANG KERAS menggunakan karakter garis lurus atau pipe (|) di dalam isi teks pesan ini karena akan merusak sistem database. Gunakan gaya WhatsApp yang ringkas, manusiawi, profesional, dan empatik — bukan format artikel atau penjelasan panjang. Jika lebih dari 1 paragraf, WAJIB beri spacer untuk menghasilkan spasi 1 baris kosong antar paragraf agar enak dibaca." +
            "###EKSEKUSI SEKARANG DENGAN FORMAT OUTPUT: Label|Jawaban\n\n" : 
            "" ) // Promt jika ini untuk karyawan
            ;
        } else {
            Logger.log("🤖 [AI MODE] Menggunakan Prompt Decision Tree");
            promptTerpakai = 
            "###DECISION TREE CHATBOT FOR QUEUE\n" +
            "###[STEP 1. IDENTIFIKASI NAMA USER/COSTUMER]\n" +
            "- " + statusNama + "\n" +
            "- Hindari kata 'anda' dan 'kamu' dalam sapaan.\n"+
            instruksiSpesifikSapaan + "\n\n" +
            "###[STEP 2. MEMORY PROCESSING & CONTEXT AWARENESS]\n" +
            "- Pahami konteks percakapan dengan mencocokkan data dari history pesan di bawah.\n" +
            "- Pahami saat ini adalah " + sapaanWaktu  + " (time : " + waktuDelayText + ") sesuai kebutuhan misalkan untuk sapaan selamat " + sapaanWaktu  + " di awal atau ucapan diakhir pecakapan\n" +
            "- ANTI-ROBOT RULE: Jawaban WAJIB bervariasi minimal 80% dari chat admin sebelumnya. Jangan gunakan pola kalimat yang persis sama berulang kali.\n" +
            "- Jika informasi (misal: alamat, harga, atau prosedur) sudah pernah dijawab di history, JANGAN dijelaskan ulang dari awal. Lanjutkan saja dengan konfirmasi atau instruksi selanjutnya.\n\n"
            + processedDevicePrompt + "\n\n" +
            "###[STEP PROSEDUR KERJA FORMAT OUTPUT]\n" +
            "Tulisakan dalam Format Ouput 3 elemen masing masing ditulis hanya 1x wajib dengan pembatas pipe (|)\n " +
            "Format Output: Label|Jawaban|" + (!rawCustName || rawCustName.includes("CEK_") ? "Nama" : "resumeAI") + "\n" +
            "1. Label: Isi 'Balas' atau 'Tidak Balas'.\n" +
            "2. Jawaban: Kolom ini WAJIB berisi hanya teks pesan yang benar-benar akan dikirim ke WhatsApp customer. ATURAN FATAL: DILARANG KERAS menggunakan karakter garis lurus atau pipe (|) di dalam isi teks pesan ini karena akan merusak sistem database. Gunakan gaya WhatsApp yang ringkas, manusiawi, profesional, dan empatik — bukan format artikel atau penjelasan panjang. Maksimal 3 kalimat per paragraf. Jika lebih dari 1 paragraf, WAJIB beri spacer (gunakan kode Regex '\\n\\n') untuk menghasilkan spasi 1 baris kosong antar paragraf agar enak dibaca. Jika harus memasukan link url maka pastikan taruh ditenggah paragraf, diselingi teks penutup setelah link agar tidak terpotong menjadi broken link\n" +
            instruksiElemen3 + "\n" +
            "###EKSEKUSI SEKARANG DENGAN FORMAT OUTPUT: Label|Jawaban|" + (!rawCustName || rawCustName.includes("CEK_") ? "Nama" : "resumeAI") + "\n"       
            ;
        }

        // 1. Buat variabel pesan user
        var waktuDelayText = provider.delayjadwalKirim ? new Date(provider.delayjadwalKirim) : new Date();
        var historyFix = cleanHistoryLama(historyMsgData);
        var pesanUser = "Tolong balas riwayat antrean pesan whatsapp ini: \n" + (historyFix || "Belum ada history");
        
        // 2. Gabungkan persis seperti yang dilakukan di dalam fungsi callAI
        var teksYangDibacaGemini = "System Instructions: " + promptTerpakai + "\n" + pesanUser;

        // --- Panggil API AI ---
        Logger.log("⏳ Memanggil API Generative Language (" + modelTerpakai + ")...");
        var rawResponse = callAI(globalAI.geminiApiKey, modelTerpakai, globalAI.fallbackModels, promptTerpakai, pesanUser, globalAI.tempGemini, globalAI.maxOutputTokens, clientSheetId);
        Logger.log("🤖 Jawaban Raw AI: " + rawResponse);

       // --- Parsing Jawaban AI ---
        var aiLabel = "Tidak Balas", aiReply = "", resumeAi = "";
        var isAIError = false; // Penanda khusus jika API AI gagal

        if (rawResponse && rawResponse !== "AI No Response") {
            // Pengaman Kasuistik: Mengekstrak baris output riil dari bawah ke atas
            var cleanResponse = rawResponse;
            if (rawResponse.indexOf("###[STEP PROSEDUR KERJA FORMAT OUTPUT]") !== -1 || rawResponse.indexOf("Format Output:") !== -1) {
              Logger.log("⚠️ Deteksi Halusinasi: AI menyalin ulang System Prompt. Mencoba mengekstrak baris output riil...");
              var lines = rawResponse.split("\n");
              for (var l = lines.length - 1; l >= 0; l--) {
                var currentLine = lines[l].trim();
                if (currentLine.toLowerCase().startsWith("balas|") || currentLine.toLowerCase().startsWith("tidak balas|") || currentLine.toLowerCase().startsWith("queue|")) {
                  cleanResponse = currentLine;
                  Logger.log("🎯 Berhasil menyelamatkan baris jawaban riil: " + cleanResponse);
                  break;
                }
              }
            }

            if (cleanResponse.indexOf("|") !== -1) {
                var parts = cleanResponse.split("|");
                aiLabel = safeTrim(parts[0]);
                var cleanReply = parts[1] ? parts[1].split(/Alasan:/i)[0].trim() : "";
                cleanReply = cleanReply.replace(/\\n/g, '\n');
                aiReply = (aiLabel.toLowerCase() === "balas" || aiLabel.toLowerCase() === "queue") ? cleanReply : "";
                
                // Mengamankan resumeAi
                if (parts[2]) {
                  resumeAi = parts[2].trim(); 
                } else if (!rawCustName || rawCustName.includes("CEK_")) {
                  resumeAi = "Kak";
                } else {
                  resumeAi = "Meresume percakapan";
                }
            } else {
                isAIError = true;
                aiLabel = "Queue";
            }
        } else {
            // JIKA API ERROR ATAU GAGAL PARSING ("AI No Response")
            isAIError = true;
            aiLabel = "Queue"; // Paksa log menjadi Queue
        }
        
        // Update memori model HANYA jika sukses membalas (Bukan Error)
        if (aiLabel.toLowerCase() === "balas" && aiReply !== "" && !isAIError) {
            var newHistoryGemi = globalAI.geminiApiKey + "~" + modelTerpakai; 
            
            // Format Baru | Lama
            globalAI.historymodel = globalAI.historymodel ? newHistoryGemi + " | " + globalAI.historymodel : newHistoryGemi;
            
            // Langsung update ke Sheet Master Setting G16
            try {
                tabSetting.getRange("G16").setValue(globalAI.historymodel);
            } catch(e) {
                Logger.log("❌ Error update G16: " + e.message);
            }
        }

        // SET ARRAY TEPAT 4 KOLOM (V, W, X, Y)
        aiLogDataUpdate = [aiLabel, teksYangDibacaGemini, resumeAi, rawResponse];

        // =========================================================================
        // KEPUTUSAN PENGIRIMAN & UPDATE SHEET
        // =========================================================================
        if (isAIError) {
            Logger.log("⚠️ AI Error / Gagal Parsing. Status dikembalikan ke 'Queue'.");
            BatchManager.pushNet(targetNetworkingRow, 22, [aiLogDataUpdate]);
            continue;
        } else if ((aiLabel.toLowerCase() === "balas" || aiLabel.toLowerCase() === "queue") && aiReply !== "") {
            
            // INI KUNCI UTAMANYA: Pindahkan teks balasan AI ke variabel pengiriman
            messageBuilt = aiReply; 
            Logger.log("🤖 AI Berhasil memotong pesan. Hasil: " + messageBuilt.substring(0, 60) + "...");
            
        } else {
            Logger.log("⚠️ AI Memutuskan TIDAK BALAS atau jawaban kosong. Skip pengiriman WA untuk baris " + realRow);
            BatchManager.pushNet(targetNetworkingRow, 22, [aiLogDataUpdate]);
            continue; 
        }

      } else {
        // =========================================================================
        // EKSEKUSI FOLLOW UP / BLAST (DENGAN AI REWRITE & SPINTEXT)
        // =========================================================================
        Logger.log("💬 [FOLLOW UP MODE] Memulai eksekusi blok non-queue.");

        // --- 1. MEMBUAT PESAN SPINTEXT ---
        var materiFollowUp = spinText(safeTrim(rowArr[11])
                              .replace(/{NAMAAKTIF}/gi, safeTrim(rowArr[5]))
                              .replace(/{NAMACS}/gi, provider.csName));
        Logger.log("📝 [PROSES 1] Materi SpinText berhasil dibentuk: " + materiFollowUp.substring(0, 50) + "...");

        // --- SETUP VARIABEL PENDUKUNG AI UNTUK BLOK INI ---
        var rawCustName = safeTrim(rowArr[5]);
        var custNameAI = rawCustName || "Kakak";
        custNameAI = custNameAI.replace(/CEK_/gi, "Kak ");
        var saPaan = getPanggilanDepan(custNameAI) || "Kak";
        var statusNama = rawCustName && !rawCustName.includes("CEK_") ? 
                         "Gunakan nama '" + rawCustName + "' 1 Kali per percakapan" : 
                         "NAMA COSTUMER BELUM TERVERIFIKASI, TERCATAT SAAT INI : " + custNameAI;
        
        var sapaanWaktu = getWaktuSapaan();
        var waktuDelayText = provider.delayjadwalKirim ? new Date(provider.delayjadwalKirim) : new Date();

        var sapaanClean = saPaan.toLowerCase();
        var instruksiSpesifikSapaan = "- Tipe sapaan santun dan hangat kepada " + (saPaan || custNameAI);
        if (["bang", "abang"].includes(sapaanClean)) instruksiSpesifikSapaan = "- Tipe sapaan AKRAB: Gunakan '" + saPaan + "' khas pelayanan Indonesia kepada customer pria.";
        else if (["kakak", "kak"].includes(sapaanClean)) instruksiSpesifikSapaan = "- Tipe sapaan AKRAB: Gunakan '" + saPaan + "' khas pelayanan Indonesia kepada semua gender.";
        else if (["mas"].includes(sapaanClean)) instruksiSpesifikSapaan = "- Gunakan '" + saPaan + "' boleh di depan atau belakang nama dengan nada bicara profesional dan sedikit lebih akrab kepada laki-laki.";
        else if (["mba", "mbak"].includes(sapaanClean)) instruksiSpesifikSapaan = "- Gunakan '" + saPaan + "' boleh di depan atau belakang nama dengan nada bicara profesional dan sedikit lebih akrab kepada customer perempuan.";
        else if (["pak", "bapak"].includes(sapaanClean)) instruksiSpesifikSapaan = "- Tipe sapaan FORMAL kepada laki-laki dengan nada bicara profesional dan sedikit lebih santun.";
        else if (["ibu", "bu", "bunda", "ustadzah"].includes(sapaanClean)) instruksiSpesifikSapaan = "- Tipe sapaan santun dan hangat kepada customer perempuan seperti ibu sendiri.";

        var instruksiElemen3 = (!rawCustName || rawCustName.includes("CEK_")) ? 
                               "4. Nama Sapaan Akrab : diisi nama sapaan costumer dari history percakapan beserta sapaan akrab yang disukai costumer\n" : 
                               "4. resumeAI : meresume konteks percakapan dari history pesan\n";

        var historyFix = cleanHistoryLama(historyMsgData);
        var modelTerpakai = globalAI.geminiModel; // Selalu pakai Gemini

    // =========================================================================
    // MODIFIKASI: CEK KONDISI AEr -> AE (KOLOM R=0, S>0)
    // =========================================================================
    var isPeralihanAER = false;
    if (NetworkingMatrix[targetNetworkingRow - 1]) {
        var labelC_Networking = safeTrim(NetworkingMatrix[targetNetworkingRow - 1][2]); // Index 2 = Kolom C
        var valR_Networking = toAngkaInt(NetworkingMatrix[targetNetworkingRow - 1][17]); // Index 17 = Kolom R
        var valS_Networking = toAngkaInt(NetworkingMatrix[targetNetworkingRow - 1][18]); // Index 18 = Kolom S

        if (labelC_Networking.toUpperCase().includes("AER") && valR_Networking === 0 && valS_Networking > 0) {
            Logger.log("🔄 [PERUBAHAN LABEL] Kondisi AEr terpenuhi (R=0, S>0). Mengubah AEr menjadi AE...");
            
            // 1. Ubah tulisan AEr menjadi AE di memori
            var newLabelC = labelC_Networking.replace(/AEr/gi, "AE");
            
            // 2. Tulis langsung ke Sheet Networking Kolom C (Indeks 3)
            BatchManager.pushNet(targetNetworkingRow, 3, [[newLabelC]]);
            
            // 3. Tandai agar mode AI menggunakan Prompt Notifikasi Khusus
            isPeralihanAER = true;
        }
    }
    // =========================================================================

        // --- 2. SETUP PROMPT DECISION TREE & NOTIFIKASI ---
        var promptTerpakai = "";
        if (isPeralihanAER) {
            Logger.log("🤖 [FOLLOW UP MODE] Menggunakan Prompt KHUSUS Peralihan AEr -> AE (Notifikasi Warmer Berhenti)");
            promptTerpakai = 
            "###[INSTRUKSI KHUSUS: NOTIFIKASI SISTEM AEr -> AE]\n" +
            "###[STEP 1. IDENTIFIKASI NAMA USER/COSTUMER]\n" +
            "-  " + statusNama + "\n" +
            "- Hindari kata 'anda' dan 'kamu' dalam sapaan.\n" +
            instruksiSpesifikSapaan + "\n\n" +
            "- Pahami saat ini adalah waktu " + sapaanWaktu  + " (time : " + waktuDelayText + ").\n" +
            "###[TUGAS UTAMA]\n" +
            "Buat 1 pesan notifikasi yang sopan, profesional, dan empatik kepada customer dengan panduan berikut:\n" +
            "1. Sampaikan secara halus bahwa karena pesan admin pada hari sebelumnya belum mendapat balasan, maka mulai besok sistem pesan otomatis (warmer) dari kami akan dijeda/dihentikan.\n" +
            "2. Beritahukan bahwa admin tidak akan mengirimkan pesan duluan lagi secara otomatis demi menjaga kenyamanan customer agar tidak merasa terganggu.\n" +
            "3. Tutup pesan dengan ramah: Sampaikan bahwa jika suatu saat customer membutuhkan bantuan, memiliki pertanyaan, atau ingin melanjutkan obrolan, mereka bisa membalas pesan ini kapan saja dan admin akan dengan senang hati membantu.\n" +
            "4. TONE: Ramah, melayani, tidak menyalahkan customer (no guilt-tripping), dan sangat profesional.\n\n" +
            "###[STEP PROSEDUR KERJA FORMAT OUTPUT]\n" +
            "Tulisakan dalam Format Ouput 2 elemen masing masing ditulis hanya 1x wajib dengan pembatas pipe (|)\n " +
            "Format Output: Label|Jawaban\n" +
            "1. Label: Isi 'Balas|'.\n" +
            "2. Jawaban: WAJIB berisi teks pesan WhatsApp sesuai instruksi di atas. DILARANG menggunakan karakter garis lurus atau pipe (|) di dalam isi teks. Maksimal 2 paragraf pendek.\n" +
            "###EKSEKUSI SEKARANG DENGAN FORMAT OUTPUT: Label|Jawaban\n:";
        } else {
            Logger.log("🤖 [PROSES 2] Menggunakan Prompt Decision Tree Follow Up (Soft Selling).");
            promptTerpakai = 
            "###DECISION TREE CHATBOT FOR FU SOFT SELLING\n" +
            "###[STEP 1. IDENTIFIKASI NAMA USER/COSTUMER]\n" +
            "-  " + statusNama + "\n" +
            "- Hindari kata 'anda' dan 'kamu' dalam sapaan.\n" +
            instruksiSpesifikSapaan + "\n\n" +
            "###[STEP 2. FOLLOW UP COSTUMER (SOFT SELLING)]\n" +
            "- Pahami saat ini adalah waktu " + sapaanWaktu  + " (time : " + waktuDelayText + ").\n" +
            "- TUGAS UTAMA: Rangkai pesan WhatsApp untuk soft-selling menggunakan [MATERI FOLLOW UP] di bawah secara luwes.\n\n" +
            "###[ATURAN SAPAAN PEMBUKA (PROFILING TEKS)]\n" +
            "- Analisa riwayat chat dan materi FU secara cermat untuk menentukan salam pembuka:\n" +
            "  * JIKA ADA indikator Islami (contoh: kata wakaf, sedekah, ustadz, alhamdulillah, atau ada salam Islami di chat sebelumnya): WAJIB awali pesan dengan 'Assalamualaikum' (atau variasi salam Islami yang hangat).\n" +
            "  * JIKA TIDAK ADA indikator Islami (konteks percakapan netral/umum): DILARANG KERAS menggunakan 'Assalamualaikum' atau istilah agama tertentu. WAJIB gunakan sapaan universal yang ramah dan profesional (misal: 'Halo', 'Hai', atau 'Selamat " + sapaanWaktu + "').\n\n" +
            "###[STRATEGI BRIDGING (PILIH SALAH SATU SESUAI KONDISI HISTORY)]\n" +
            "- KONDISI 1 (CHAT TERAKHIR DARI CUSTOMER & RELEVAN): Buat 1 kalimat transisi/basa-basi empati yang merespons/menyambung chat terakhir customer, kemudian baru masuk ke materi FU secara halus. Jangan tiba-tiba mengirim link.\n" +
            "- KONDISI 2 (CHAT TERAKHIR DARI CUSTOMER TAPI TIDAK NYAMBUNG): Abaikan topik history sebelumnya. Buat kalimat pembuka natural seolah ini adalah pesan penyapaan baru yang segar, lalu sampaikan materi FU.\n" +
            "- KONDISI 3 (CHAT TERAKHIR DARI ADMIN & BELUM DIBALAS/MENGGANTUNG): Jangan kaku. Ulas tipis atau singgung sedikit pesan admin sebelumnya agar terasa nyambung (misal: 'Melanjutkan info sebelumnya ya...', atau 'Semoga pesan sebelumnya sudah dibaca...'), lalu masuk ke materi FU yang baru.\n\n" +
            "- ATURAN PENUTUP (WAJIB): Apapun kondisinya, setiap pesan Follow Up WAJIB diakhiri dengan kalimat tanya pancingan terbuka (Call to Action) yang sopan untuk memancing feedback/balasan dari customer.\n\n" +
            "MATERI FOLLOW UP:\n" + materiFollowUp + "\n\n" +
            "###[STEP PROSEDUR KERJA FORMAT OUTPUT]\n" +
            "Tulisakan dalam Format Ouput 3 elemen masing masing ditulis hanya 1x wajib dengan pembatas pipe (|)\n " +
            "Format Output: Label|Jawaban|" + (!rawCustName || rawCustName.includes("CEK_") ? "Nama" : "resumeAI") + "\n" +
            "1. Label: Isi 'Balas' atau 'Tidak Balas'.\n" +
            "2. Jawaban: Kolom ini WAJIB berisi hanya teks pesan yang benar-benar akan dikirim ke WhatsApp customer. ATURAN FATAL: DILARANG KERAS menggunakan karakter garis lurus atau pipe (|) di dalam isi teks pesan ini karena akan merusak sistem database. Gunakan gaya WhatsApp yang ringkas, manusiawi, profesional, dan empatik — bukan format artikel atau penjelasan panjang. Maksimal 3 kalimat per paragraf. Jika lebih dari 1 paragraf, WAJIB beri spacer (gunakan kode Regex '\\n\\n') untuk menghasilkan spasi 1 baris kosong antar paragraf agar enak dibaca. Jika harus memasukan link url maka pastikan taruh ditenggah paragraf, diselingi teks penutup setelah link agar tidak terpotong menjadi broken link\n" +
            instruksiElemen3 + "\n" +
            "###EKSEKUSI SEKARANG DENGAN FORMAT OUTPUT: Label|Jawaban|" + (!rawCustName || rawCustName.includes("CEK_") ? "Nama" : "resumeAI") + "\n:";
        }

        // 1. Buat variabel pesan user
        var historyFix = cleanHistoryLama(historyMsgData);
        var pesanUser = "Tolong balas riwayat antrean pesan whatsapp ini: \n" + (historyFix || "Belum ada history");
        
        // 2. Gabungkan persis seperti yang dilakukan di dalam fungsi callAI
        var teksYangDibacaGemini = "System Instructions: " + promptTerpakai + "\n" + pesanUser;

        Logger.log("⏳ [PROSES 2] Memanggil API AI (" + modelTerpakai + ")...");
        var rawResponse = callAI(globalAI.geminiApiKey, modelTerpakai, globalAI.fallbackModels, promptTerpakai, pesanUser, globalAI.tempGemini, globalAI.maxOutputTokens);
        Logger.log("🤖 [PROSES 2] Jawaban Raw AI diterima.");

        // Parsing Jawaban AI
        var aiLabel = "Tidak Balas", aiReply = "", resumeAi = "";
        if (rawResponse && rawResponse.indexOf("|") !== -1) {
            var parts = rawResponse.split("|");
            aiLabel = safeTrim(parts[0]);
            var cleanReply = parts[1] ? parts[1].split(/Alasan:/i)[0].trim() : "";
            cleanReply = cleanReply.replace(/\\n/g, '\n');
            aiReply = (aiLabel.toLowerCase() === "balas" || aiLabel.toLowerCase() === "queue") ? cleanReply : "";
            
            // PERUBAHAN KRITIKAL: Index 2 untuk Kolom X
            if (parts[2]) resumeAi = parts[2].trim(); 
        }

        // Fallback safety: Jika AI gagal merespons dengan format yang benar, gunakan SpinText murni
        if (aiReply === "") {
            Logger.log("⚠️ [PROSES 2] AI gagal merespons dengan format benar. Fallback ke teks SpinText statis.");
            aiReply = materiFollowUp;
            aiLabel = "Fallback Balas";
        }

        messageBuilt = aiReply;

        if (aiLabel.toLowerCase() === "balas" || aiLabel.toLowerCase() === "queue") {
            var newHistoryGemi = globalAI.geminiApiKey + "~" + modelTerpakai; 
            
            // Format Baru | Lama
            globalAI.historymodel = globalAI.historymodel ? newHistoryGemi + " | " + globalAI.historymodel : newHistoryGemi;
            
            // Langsung update ke Sheet Master Setting G16
            try {
                tabSetting.getRange("G16").setValue(globalAI.historymodel);
            } catch(e) {
                Logger.log("❌ Error update G16: " + e.message);
            }
        }

        // --- 3. PENCATATAN LOG AI (KOLOM V, W, X, Y, Z) ---
        Logger.log("📋 [PROSES 3] Mencatat Log AI ke array memori (Kolom V-Z).");
        var historyGemiMode = ""; // Sesuai aturan arsitektur batching 4 elemen
        aiLogDataUpdate = [aiLabel, teksYangDibacaGemini, resumeAi, rawResponse];

       // --- 4. UPDATE KE SHEET Networking (LOGIKA PERSONAL KOLOM AF) ---
        Logger.log("🗓️ [PROSES 4] Menghitung jadwal Reminder dari Kolom AF Networking.");
        
        // MENGAMBIL DAY REMINDER DARI KOLOM AF NETWORKING (Indeks Array ke-31)
        var rawDayReminder = 0;
        if (NetworkingMatrix[targetNetworkingRow - 1] && NetworkingMatrix[targetNetworkingRow - 1][31]) {
            rawDayReminder = NetworkingMatrix[targetNetworkingRow - 1][31];
        }
        
        var hariTambah = toAngkaInt(rawDayReminder);
        
        // JIKA DATANYA KOSONG ATAU 0, GUNAKAN DEFAULT 30 HARI
        if (hariTambah <= 0) {
            hariTambah = 30;
        }

        Logger.log("⏳ Nilai Hari Tambahan untuk baris ini: " + hariTambah + " hari.");

        // LOGIKA PENULISAN AC (29) dan AD (30) DIBUAT JADI SATU BARIS
        var d = new Date(); 
        d.setDate(d.getDate() + hariTambah);
        var tglReminder = Utilities.formatDate(d, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss");
        var nilaiLama = 0;
        if (NetworkingMatrix[targetNetworkingRow - 1] && NetworkingMatrix[targetNetworkingRow - 1][28]) {
            nilaiLama = toAngkaInt(NetworkingMatrix[targetNetworkingRow - 1][28]);
        }
        var totalAC = nilaiLama + ((Number(nilaiBaru) > 0) ? 1 : 0);
        
        BatchManager.pushNet(targetNetworkingRow, 29, [[totalAC, tglReminder]]); 
        
        Logger.log("✅ [PROSES 4] Update Networking Kolom AC & AD selesai (Masuk Antrean Batch).");
      }
      
      // --- PUSH KE SERVER ---
      var msg = {};
      var finalApiUrl = provider.apiUrl; // Variabel penampung URL yang bisa diubah

      if (isStarsender) {
        msg = (messageType === 'text' || !mediaLink) ? 
          { "messageType": "text", "body": messageBuilt, "to": recipientDigits } : 
          { "messageType": "media", "body": messageBuilt, "file": mediaLink, "to": recipientDigits };
          
        // ===============================================================
        // MODIFIKASI: Ubah endpoint URL Starsender khusus untuk Grup
        // ===============================================================
        if (isGroup) {
          // Jika URL dasar di Sheet Setting berakhiran "/send", tambahkan "/grup"
          if (finalApiUrl.endsWith("/send")) {
            finalApiUrl = finalApiUrl + "/grup";
          } else if (finalApiUrl.indexOf("/send/grup") === -1) {
            // Berjaga-jaga jika format penulisan di Sheet Setting ada garis miring ekstra
            finalApiUrl = finalApiUrl.replace(/\/$/, "") + "/grup";
          }
        }
        
      } else {
        // Pilihan untuk OneSender
        msg = (messageType === 'text' || !mediaLink) ? 
          buildOneSenderText(recipientDigits, messageBuilt, isGroup) : 
          buildOneSenderImage(recipientDigits, mediaLink, messageBuilt, isGroup);
      }

      // [TAMBAHAN LOG] Debug Payload sebelum masuk Batching
      Logger.log("--------------------------------------------------");
      Logger.log("📦 [DEBUG PAYLOAD] Persiapan kirim untuk Baris: " + realRow);
      Logger.log("🔗 URL Endpoint : " + finalApiUrl);
      Logger.log("🔑 API Key      : " + (provider.apiKey ? provider.apiKey.substring(0, 10) + "..." : "KOSONG/UNDEFINED"));
      Logger.log("📄 JSON Payload : " + JSON.stringify(msg));
      Logger.log("--------------------------------------------------");

      // Pastikan mem-push menggunakan finalApiUrl yang sudah difilter
      StarSender.pushMessage(msg, targetNetworkingRow, provider.apiKey, finalApiUrl, aiLogDataUpdate);
      
    } catch (err) {
      Logger.log("❌ GAGAL FATAL di baris " + realRow + ": " + err.message);
    }
  } // <-- Ini adalah penutup dari looping utama (for loop)

  StarSender.flush(tabNetworking);
  
  // =======================================================================
  // FINAL BATCH COMMIT: Eksekusi semua perubahan ke Sheet dalam 1 tarikan napas
  // =======================================================================
  BatchManager.commit(tabNetworking, tabSetting);
  
  Logger.log("🎉 Selesai semua proses antrean.");
} catch (error) {
    Logger.log("❌ Terjadi kesalahan: " + error.message);
  } finally {
    // 3. Lepaskan kunci setelah selesai atau jika terjadi error
    lock.releaseLock();
  }
}

/* =========================================================================
   FUNGSI HELPER & AI ENGINE
   ========================================================================= */

function callAI(apiKey, primaryModel, fallbackStr, systemPrompt, userMessage, temp, maxTokens, clientSheetId) { // <-- Tambah parameter di sini
  var fallbackArr = fallbackStr ? fallbackStr.split("|").map(function(m){ return m.trim(); }) : [];
  var modelsToTry = [primaryModel].concat(fallbackArr);
  
  modelsToTry = modelsToTry.filter(function(item, pos) {
    return modelsToTry.indexOf(item) == pos && item !== "";
  });

  var cleanTemp = temp.toString().replace(",", ".");
  var finalTemp = parseFloat(cleanTemp) || 0.7;
  var finalMaxTokens = parseInt(maxTokens) || 2048; 

  var payload = { 
    "contents": [{ "role": "user", "parts": [{ "text": "System Instructions: " + systemPrompt + "\n\nUser Message: " + userMessage }] }],
    "generationConfig": { "temperature": finalTemp, "maxOutputTokens": finalMaxTokens },
    "safetySettings": [
      { "category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE" },
      { "category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE" },
      { "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE" },
      { "category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE" }
    ]
  };

  var options = { 
    method: "POST", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true 
  };

  var lastError = "";

  // =========================================================
  // FUNGSI HELPER INTERNAL UNTUK MENULIS KE SHEET LOG
  // =========================================================
  function tulisLogKeSheet(judul, detail) {
    try {
      var ss = clientSheetId ? SpreadsheetApp.openById(clientSheetId) : SpreadsheetApp.getActiveSpreadsheet();
      var sheetLog = ss.getSheetByName("LOG");
      if (!sheetLog) {
        sheetLog = ss.insertSheet("LOG");
        sheetLog.appendRow(["Timestamp", "Activity Title", "Detail Pipeline"]);
        sheetLog.getRange("A1:C1").setFontWeight("bold");
      }
      sheetLog.appendRow([new Date(), judul, detail]);
    } catch(e) { /* Abaikan jika gagal nulis LOG */ }
  }
  // =========================================================

  for (var i = 0; i < modelsToTry.length; i++) {
    var currentModel = modelsToTry[i];
    var url = "https://generativelanguage.googleapis.com/v1beta/models/" + currentModel + ":generateContent?key=" + apiKey;
    
    Logger.log("🔄 [AUTO-FALLBACK] Mencoba model (" + (i+1) + "/" + modelsToTry.length + "): " + currentModel);

    try {
      var res = UrlFetchApp.fetch(url, options);
      var resCode = res.getResponseCode();
      var resText = res.getContentText();
      var json = JSON.parse(resText);
      
      if (resCode === 200 && json.candidates && json.candidates[0].content) {
        var rawText = json.candidates[0].content.parts[0].text;
        
        // JIKA INI ADALAH MODEL CADANGAN (Bukan model urutan pertama / i > 0)
        if (i > 0) {
          tulisLogKeSheet("⚠️ [AI FALLBACK AKTIF]", "Model utama gagal. Sistem berhasil menggunakan model cadangan: " + currentModel);
        }
        
        return rawText; 
      } else {
        // JIKA API GEMINI MENOLAK (Limit, Error, dll)
        var responPendek = resText.length > 200 ? resText.substring(0, 200) + "..." : resText;
        tulisLogKeSheet("❌ [AI API ERROR]", "Model " + currentModel + " GAGAL.\nKode HTTP: " + resCode + "\nAlasan: " + responPendek);
        lastError = resText;
      }
    } catch (e) {
      // JIKA KONEKSI INTERNET/FETCH GAGAL TOTAL
      tulisLogKeSheet("🚨 [AI SYSTEM CRASH]", "Gagal menghubungi model " + currentModel + ".\nError: " + e.message);
      lastError = e.message;
    }
  }

  // Jika semua model gagal
  tulisLogKeSheet("🔥 [AI FATAL ERROR]", "SEMUA MODEL GAGAL DIEKSEKUSI. Sistem menghentikan balasan AI.");
  return "AI No Response";
}

function cleanHistoryLama(rawText) {
  if (!rawText) return "";
  return rawText.replace(/^\s*_\s*$/gm, "")
                .replace(/\d{2}\/\d{2}\s+\d{2}:\d{2}:\d{2}\s+\[[A-Z0-9]+\]/g, "")
                // --- MODIFIKASI: Menghapus teks [Gambar Disimpan: ERROR_UPLOAD] ---
                .replace(/\[Gambar Disimpan:\s*ERROR_UPLOAD\]/gi, "")
                .replace(/\[Gambar Error:\s*Gagal Membuat File di GDrive \(Limit\/Timeout\)\]/gi, "")
                .replace(/[ \t]+/g, " ")
                .replace(/\n{2,}/g, "\n").trim();
}

function getPanggilanDepan(nama) {
  if (!nama || nama.includes("[")) return ""; 
  var parts = nama.trim().split(" ");
  return parts[0].toLowerCase(); 
}

/**
 * Fungsi Cek Nomor Aktif untuk Starsender
 */
function cekNomorStarsender(nomor, apiKey) {
  try {
    var resp = UrlFetchApp.fetch("https://api.starsender.online/api/check-number", {
      method: "post", headers: { "Authorization": apiKey, "Content-Type": "application/json" },
      payload: JSON.stringify({ "number": nomor }), muteHttpExceptions: true
    });
    var resText = resp.getContentText();
    Logger.log("   -> API Starsender Response untuk " + nomor + ": " + resText); 
    
    var res = JSON.parse(resText);
    
    // Cek apakah API sukses dihubungi dan device admin terkoneksi
    if (res.success === true) {
        // Device admin aktif. Baru kita percaya hasil pengecekan nomor customernya
        if (res.data && res.data.status === true) {
            return true;  // Nomor Customer Aktif
        } else {
            return false; // Nomor Customer BENAR-BENAR Mati
        }
    } else {
        // Jika success: false, artinya Device Admin yang putus / API Error
        Logger.log("⚠️ API Starsender: Device Admin Offline / Error.");
        return "DEVICE_ERROR";
    }
  } catch (e) { 
    Logger.log("⚠️ Gagal konek API Starsender. Error: " + e.message);
    return "DEVICE_ERROR"; 
  }
}

/**
 * Fungsi Cek Nomor Aktif untuk OneSender (Menggunakan Prefix URL Kolom K)
 */
function cekNomorOneSender(nomor, apiKey, prefixUrl) {
  try {
    if (!prefixUrl) {
       Logger.log("⚠️ Prefix URL Kolom K kosong. Cek nomor OneSender dilewati (Bypass).");
       return true; 
    }

    var cleanPrefix = prefixUrl.replace(/\/$/, ""); 
    var checkUrl = cleanPrefix + "/api/v1/check-number?phone_number=" + nomor;

    var options = {
      method: "get", 
      headers: { "Authorization": apiKey },
      muteHttpExceptions: true
    };

    var resp = UrlFetchApp.fetch(checkUrl, options);
    var resText = resp.getContentText();
    
    Logger.log("   -> API OneSender Response untuk " + nomor + ": " + resText); 
    var res = JSON.parse(resText);
    
    // Deteksi apakah respon API valid dan device admin berhasil mengecek
    if (res && res.data && typeof res.data.has_account !== "undefined") {
      if (res.data.has_account === true) {
        return true;  // Nomor Customer Aktif
      } else {
        return false; // Nomor Customer BENAR-BENAR Mati
      }
    } else {
      // Jika data tidak lengkap, artinya Device Admin Offline atau Limit API
      Logger.log("⚠️ API OneSender: Device Admin Offline / Error JSON.");
      return "DEVICE_ERROR";
    }
    
  } catch (e) { 
    Logger.log("⚠️ Gagal konek API OneSender. Error: " + e.message);
    return "DEVICE_ERROR"; 
  }
}

function isDalamJamKerja(waktuKirim) {
  var tz = Session.getScriptTimeZone();
  var jamString = Utilities.formatDate(new Date(), tz, "HH");
  var menitString = Utilities.formatDate(new Date(), tz, "mm");
  var currentMinutes = parseInt(jamString, 10) * 60 + parseInt(menitString, 10);
  var ranges = waktuKirim.split(",");
  for (var i = 0; i < ranges.length; i++) {
    var part = ranges[i].trim().split("-");
    if (part.length < 2) continue;
    var s = part[0].split(":"), e = part[1].split(":");
    var startMins = parseInt(s[0], 10) * 60 + parseInt(s[1], 10);
    var endMins = parseInt(e[0], 10) * 60 + parseInt(e[1], 10);
    if (startMins < endMins) { if (currentMinutes >= startMins && currentMinutes < endMins) return true; } 
    else { if (currentMinutes >= startMins || currentMinutes < endMins) return true; }
  }
  return false;
}

function safeTrim(v) { return (v === null || v === undefined) ? "" : (v + "").trim(); }

function toAngkaInt(angka) {
  if (!angka || angka === "") return 0;
  if (typeof angka === 'number') return Math.floor(angka);
  var clean = angka.toString().replace(/[^0-9]/g, '');
  return clean ? parseInt(clean, 10) : 0;
}

function autoPrefixPhone(s) {
  s = s.toString();
  if (s.startsWith('0')) return '62' + s.substring(1);
  if (s.startsWith('8')) return '62' + s;
  return s;
}

function spinText(text) {
  return text ? text.replace(/{([^{}]+)}/g, function(_, o) {
    var opts = o.split('|'); return opts[Math.floor(Math.random() * opts.length)];
  }) : text;
}

// GANTI / UPDATE FUNGSI PEMBENTUK PAYLOAD ONESENDER
function buildOneSenderImage(r, l, m, isGroup) {
  var recipientType = isGroup ? "group" : "individual";
  var imageObject = {};
  
  // Jika input berupa data Base64
  if (l && l.startsWith("data:image")) {
    imageObject = { "raw": l, "caption": m };
  } else {
    // Jika input berupa Link URL
    imageObject = { "link": l, "caption": m };
  }

  return {
    "recipient_type": recipientType,
    "to": r,
    "type": "image",
    "image": imageObject
  };
}

function buildOneSenderText(r, m, isGroup) { 
  return { 
    "recipient_type": isGroup ? "group" : "individual", 
    "to": r, 
    "type": "text", 
    "text": { "body": m } 
  }; 
}

/**
 * FUNGSI MENDETEKSI WAKTU (Pagi/Siang/Sore/Malam)
 * Aman menerima input teks waktu, objek Date, atau dikosongkan.
 */
function getWaktuSapaan(timeInput) {
  let dateObj;
  
  // Validasi input agar tidak error
  if (timeInput instanceof Date) {
    dateObj = timeInput; // Jika sudah format waktu, langsung pakai
  } else if (timeInput && typeof timeInput === 'string') {
    dateObj = new Date(timeInput.replace(/-/g, "/")); // Jika teks, ubah ke waktu
  } else {
    dateObj = new Date(); // Jika kosong, pakai waktu saat ini
  }

  // Ambil angka jamnya saja (Format 24 Jam: 0 - 23)
  const hour = dateObj.getHours();

  // Logika Pembagian Waktu Khas Indonesia
  if (hour >= 3 && hour < 11) {
    return "Pagi";
  } else if (hour >= 11 && hour < 15) {
    return "Siang";
  } else if (hour >= 15 && hour < 18) {
    return "Sore";
  } else {
    return "Malam";
  }
}

// =========================================================================
// JEMBATAN DATA UTAMA (SERVER-SIDE FILTERING, SORTING & VIRTUAL VLOOKUP)
// =========================================================================
function getSheetData(sheetUrl, userAccountsStr) {
  try {
    var ss = sheetUrl ? SpreadsheetApp.openByUrl(sheetUrl) : SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Networking");
    var tabSetting = ss.getSheetByName("Setting");
    
    if (!sheet) return [];

    // =========================================================
    // 1. BACA RASIO, API KEY & API URL DARI SETTING (VLOOKUP VIRTUAL)
    // =========================================================
    var deviceConfig = {};
    if (tabSetting) {
      var lastSetRow = tabSetting.getLastRow();
      if (lastSetRow >= 34) {
        var settingData = tabSetting.getRange(34, 1, lastSetRow - 33, 10).getDisplayValues();
        for (var i = 0; i < settingData.length; i++) {
          var devName = String(settingData[i][0]).trim(); // Kolom A (Nama Device)
          var rasioVal = String(settingData[i][3]).replace(/[^0-9.]/g, ''); // Kolom D (Rasio)
          var apiKeyStr = String(settingData[i][8]).trim(); // Kolom I (API Key)
          var apiUrlStr = String(settingData[i][9]).trim(); // Kolom J (API URL)
          
          deviceConfig[devName] = {
            rasio: parseFloat(rasioVal) || 0,
            apiKey: apiKeyStr,
            apiUrl: apiUrlStr
          };
        }
      }
    }
    
    // Ambil limit dari H20
    var limitRaw = tabSetting.getRange("H20").getValue();
    var limitData = parseInt(limitRaw, 10);
    if (isNaN(limitData) || limitData <= 0) limitData = 150;
    
    var data = sheet.getDataRange().getDisplayValues();
    if (data.length <= 1) return []; // Hanya header
    
    var headers = data[0];
    var rows = data.slice(1);
    
    // Helper untuk mengubah tanggal teks menjadi angka waktu
    function parseDate(t) {
      if (!t) return 0;
      var match = t.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
      if (match) return new Date(match[2] + '/' + match[1] + '/' + match[3]).getTime();
      return new Date(t).getTime() || 0;
    }
    
    // =========================================================
    // BARU: SERVER SIDE FILTERING BERDASARKAN ROLE
    // =========================================================
    var allowedAccounts = [];
    var isMaster = false;
    
    // Mengurai parameter string menjadi array jika ada
    if (userAccountsStr) {
      allowedAccounts = userAccountsStr.split(",").map(function(s) { return s.trim(); });
      isMaster = allowedAccounts.includes("Master Admin");
    } else {
      isMaster = true; // Jika dipanggil tanpa parameter (misal testing), anggap Master
    }

    var filteredRows = rows;
    
    // Jika BUKAN Master Admin, filter data berdasarkan Kolom A (Akun)
    if (!isMaster && allowedAccounts.length > 0) {
      filteredRows = rows.filter(function(r) {
        var akunChat = String(r[0] || "").trim(); // Kolom A (Akun)
        return allowedAccounts.includes(akunChat);
      });
    }

    // Sortir baris hasil filter berdasarkan kolom P atau Q
    filteredRows.sort(function(a, b) {
      var timeA = Math.max(parseDate(a[15]), parseDate(a[16]));
      var timeB = Math.max(parseDate(b[15]), parseDate(b[16]));
      return timeB - timeA; 
    });
    
    // Limit data sesuai H20 HANYA SETELAH DI-FILTER DAN DISORTIR
    var limitedRows = filteredRows.slice(0, limitData);

    // =========================================================
    // 2. SUNTIKKAN DATA VIRTUAL KE DALAM DATA UI
    // =========================================================
    limitedRows = limitedRows.map(function(r) {
      var fullId = String(r[1] || "").trim(); // Kolom B Networking
      var parts = fullId.split('.');
      var devName = parts.length > 1 ? parts.slice(1).join('.') : "Default"; // Ekstrak Nama Device
      
      // Ambil konfigurasi (Rasio, Key, URL) dari memori virtual Setting
      var config = deviceConfig[devName] || { rasio: 0, apiKey: "", apiUrl: "" };
      
     // MEMASTIKAN ARRAY CUKUP PANJANG (Penting!)
      // Perpanjang array menjadi 33 agar memiliki ruang aman di belakang
      while (r.length <= 33) { r.push(""); }
      
      // SUNTIKKAN VIRTUAL VLOOKUP UNTUK DIKIRIM KE UI
      // Pindahkan dari index 8 & 9 (yang menabrak Biodata) ke index yang aman
      r[31] = config.apiKey; // Suntik API Key ke Index 31
      r[32] = config.apiUrl; // Suntik API URL ke Index 32
      r[33] = config.rasio;  // Suntik Rasio ke Index 33
      return r;
    });
    
    return [headers].concat(limitedRows);
    
  } catch(e) {
    Logger.log("Error getSheetData: " + e.message);
    return [];
  }
}

// =========================================================================
// SERVER-SIDE SEARCH (MENCARI KONTAK DI LUAR LIMIT 150)
// =========================================================================
function searchContactNative(sheetUrl, keyword) {
  try {
    var ss = sheetUrl ? SpreadsheetApp.openByUrl(sheetUrl) : SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Networking");
    if (!sheet) return [];
    
    var data = sheet.getDataRange().getDisplayValues();
    if (data.length <= 1) return [];
    
    var headers = data[0];
    var kw = String(keyword).toLowerCase().trim();
    
    // Filter seluruh 2000+ baris mencari ID (Kolom B) atau Nama (Kolom H)
    var results = data.slice(1).filter(function(row) {
      var id = String(row[1]).toLowerCase();
      var name = String(row[7]).toLowerCase();
      return id.indexOf(kw) !== -1 || name.indexOf(kw) !== -1;
    });
    
    return [headers].concat(results); // Mengembalikan hanya baris yang cocok
  } catch(e) {
    return [];
  }
}

// =========================================================================
// FITUR CONTROL PANEL UI (PENGATURAN & DEVICE)
// =========================================================================
// UBAH FUNGSI INI KESELURUHAN
function getSettingsForUI(sheetInput) {
  try {
    var ss = sheetInput ? SpreadsheetApp.openByUrl(sheetInput) : SpreadsheetApp.getActiveSpreadsheet();
    var tabSetting = ss.getSheetByName("Setting");
    
    // Baca Array Range (Lebih cepat daripada memanggil 1 per 1)
    var labelsData = tabSetting.getRange("E3:H12").getValues(); // Index 0=E, 3=H
    var sendSettings = tabSetting.getRange("B17:B24").getValues();
    var adminData = tabSetting.getRange("A27:C30").getValues();
    
    var data = {
      webhookUrl: tabSetting.getRange("A2").getValue(),
      ui: { limitData: tabSetting.getRange("H20").getValue() },
      ai: {
        maxTokens: tabSetting.getRange("E16").getValue(),
        temp: tabSetting.getRange("F16").getValue(),
        limitHistory: tabSetting.getRange("E20").getValue(),
        tokenGemini: tabSetting.getRange("H22").getValue(),
        modelGemini: tabSetting.getRange("H18").getValue(),
        warmer: tabSetting.getRange("F20").getValue(),
        qtyWarmer: tabSetting.getRange("G20").getValue()
      },
      send: sendSettings.map(r => r[0]), // Mengambil kolom B saja
      labels: labelsData.map(r => ({ nama: r[0], ket: r[3] })), // E & H
      admins: adminData.map(r => ({ user: r[0], pass: r[1], role: r[2] })), // A, B, C
          devices: [],
          akunList: [] // <--- Wadah untuk List Akun
        };

        var lastRow = tabSetting.getLastRow();
        var akunSet = new Set(); // <--- Menggunakan Set untuk mencegah nama akun ganda

        if (lastRow >= 34) {
          // 2. Tarik Penuh 26 Kolom (A sampai Z)
          var deviceData = tabSetting.getRange(34, 1, lastRow - 33, 26).getDisplayValues();
          for (var i = 0; i < deviceData.length; i++) {
            if (deviceData[i][0] !== "") {
              data.devices.push({
                index: i, 
                providerId: deviceData[i][0], 
                csName: deviceData[i][1], 
                status: deviceData[i][5],
                fullRow: deviceData[i] // 3. Lempar Data Utuh ke UI untuk Form Edit
              });

              // Kumpulkan Nama Akun (Kolom B)
              let namaAkun = String(deviceData[i][1]).trim();
              if (namaAkun !== "") {
                  akunSet.add(namaAkun);
              }
            }
          }
        }
        // Masukkan daftar akun yang sudah disaring ke dalam data
        data.akunList = Array.from(akunSet);
        return data;
  } catch(e) { throw new Error(e.message); }
}

// FORMAT OBJECT PROFESIONAL
function saveGeneralSettings(sheetInput, dataObj) { 
  var ss = sheetInput ? SpreadsheetApp.openByUrl(sheetInput) : SpreadsheetApp.getActiveSpreadsheet();
  var tabSetting = ss.getSheetByName("Setting");
  
  // 1. Simpan UI & AI
  tabSetting.getRange("H20").setValue(dataObj.ui.limitData);
  tabSetting.getRange("E16").setValue(dataObj.ai.maxTokens);
  tabSetting.getRange("F16").setValue(dataObj.ai.temp);
  tabSetting.getRange("E20").setValue(dataObj.ai.limitHistory);
  tabSetting.getRange("H22").setValue(dataObj.ai.tokenGemini);
  tabSetting.getRange("H18").setValue(dataObj.ai.modelGemini);
  tabSetting.getRange("F20").setValue(dataObj.ai.warmer);
  tabSetting.getRange("G20").setValue(dataObj.ai.qtyWarmer);

  // 2. Simpan SEND (Dari B17 sampai B24)
  let arrSend = dataObj.send.map(s => [s]);
  tabSetting.getRange("B17:B24").setValues(arrSend);

  // 3. Simpan LABELS (Menyimpan Kolom E & H secara terpisah agar F tidak tertimpa)
  let arrE = dataObj.labels.map(l => [l.nama]);
  let arrH = dataObj.labels.map(l => [l.ket]);
  tabSetting.getRange("E3:E12").setValues(arrE);
  tabSetting.getRange("H3:H12").setValues(arrH);

  // 4. Simpan ADMINS (A27 sampai C30)
  let arrAdmins = dataObj.admins.map(a => [a.user, a.pass, a.role]);
  tabSetting.getRange("A27:C30").setValues(arrAdmins);
  
  return "Pengaturan sistem (UI, AI, Send, Label, Admin) berhasil diperbarui!";
}

// FUNGSI GABUNGAN: EDIT & TAMBAH BARU (ANTI UNDEFINED)
function saveDeviceDataNative(sheetInput, index, dataObj) {
  try {
    var ss = sheetInput ? SpreadsheetApp.openByUrl(sheetInput) : SpreadsheetApp.getActiveSpreadsheet();
    var tabSetting = ss.getSheetByName("Setting");
    
    var targetRow;
    // Jika index kosong, berarti ini perintah "Tambah Baru"
    if (index === "" || index === null || index === undefined) {
        var lastRow = tabSetting.getLastRow();
        targetRow = lastRow < 33 ? 34 : lastRow + 1; // Taruh di baris paling bawah
    } else {
        // Jika index ada angkanya, berarti ini perintah "Edit"
        targetRow = 34 + parseInt(index);
    }
    
    // Pengaman: Jika data dari UI hilang/undefined, ubah jadi string kosong ("") agar script tidak Crash
    function safeVal(val) { return val === undefined ? "" : val; }

    tabSetting.getRange(targetRow, 1).setValue(safeVal(dataObj.col0));   // Kolom A (Device Name)
    tabSetting.getRange(targetRow, 2).setValue(safeVal(dataObj.col1));   // Kolom B (Akun CS)
    tabSetting.getRange(targetRow, 3).setValue(safeVal(dataObj.col2));   // Kolom C (Kirim Pesan)
    tabSetting.getRange(targetRow, 5).setValue(safeVal(dataObj.col4));   // Kolom E (Auto Label) <--- DIPINDAH KE SINI
    tabSetting.getRange(targetRow, 7).setValue(safeVal(dataObj.col6));   // Kolom G (No Server)
    tabSetting.getRange(targetRow, 8).setValue(safeVal(dataObj.col7));   // Kolom H (BSUID)
    tabSetting.getRange(targetRow, 9).setValue(safeVal(dataObj.col8));   // Kolom I (API Key)
    tabSetting.getRange(targetRow, 10).setValue(safeVal(dataObj.col9));  // Kolom J (API URL)
    tabSetting.getRange(targetRow, 11).setValue(safeVal(dataObj.col10)); // Kolom K (Prefix URL)
    tabSetting.getRange(targetRow, 12).setValue(safeVal(dataObj.col11)); // Kolom L (AI Replay)
    tabSetting.getRange(targetRow, 13).setValue(safeVal(dataObj.col12)); // Kolom M (Prompt Gemini)
    tabSetting.getRange(targetRow, 14).setValue(safeVal(dataObj.col13)); // Kolom N (promt Vision)
    tabSetting.getRange(targetRow, 24).setValue(safeVal(dataObj.col23)); // Kolom X (Jenis)
    tabSetting.getRange(targetRow, 25).setValue(safeVal(dataObj.col24)); // Kolom Y (Masa Aktif)
    tabSetting.getRange(targetRow, 26).setValue(safeVal(dataObj.col25)); // Kolom Z (Set Rasio)
    
    return (index === "" || index === null) ? "✅ Device Baru Berhasil Ditambahkan!" : "✅ Data Device Berhasil Diperbarui!";
  } catch (e) {
    throw new Error("Gagal menyimpan ke Spreadsheet: " + e.message);
  }
}

function deleteDeviceFromUI(sheetInput, index) {
  var ss = sheetInput ? SpreadsheetApp.openByUrl(sheetInput) : SpreadsheetApp.getActiveSpreadsheet();
  var tabSetting = ss.getSheetByName("Setting");
  var targetRow = 34 + parseInt(index);
  tabSetting.deleteRow(targetRow);
  return "Device berhasil dihapus secara permanen.";
}

function deleteContactData(sheetUrl, contactId) {
  try {
    const ss = SpreadsheetApp.openByUrl(sheetUrl);
    const sheet = ss.getSheetByName("Database"); // Ganti "Database" dengan nama sheet database kontakmu
    const data = sheet.getDataRange().getValues();
    
    // Cari baris berdasarkan contactId (Asumsi ID ada di Kolom B / index 1)
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() === String(contactId).trim()) {
        // Hapus baris (i + 1 karena array index mulai dari 0, row sheet mulai dari 1)
        sheet.deleteRow(i + 1);
        return "success";
      }
    }
    
    return "error_not_found";
  } catch (e) {
    return "error: " + e.message;
  }
}

// =========================================================================
// FITUR LOGIN & AUTENTIKASI
// =========================================================================
function verifyLogin(sheetUrl, user, pass) {
  try {
    var ss = SpreadsheetApp.openByUrl(sheetUrl);
    var sheet = ss.getSheetByName("Setting");
    if (!sheet) return { success: false, msg: "Sheet Setting tidak ditemukan." };
    
    var adminData = sheet.getRange("A27:C30").getValues();
    
    for (var i = 0; i < adminData.length; i++) {
      var dbUser = String(adminData[i][0]).trim();
      var dbPass = String(adminData[i][1]).trim();
      var dbRoleStr = String(adminData[i][2]).trim();
      
      if (dbUser !== "" && dbUser === user && dbPass === pass) {
        var roles = dbRoleStr.split(",").map(function(s) { return s.trim(); });
        var isMaster = roles.includes("Master Admin");
        return { success: true, isMaster: isMaster, roles: roles };
      }
    }
    return { success: false, msg: "Username atau Password salah." };
  } catch (e) {
    return { success: false, msg: "Error Server: " + e.message };
  }
}
