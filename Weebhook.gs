/************************************************************
 * WEBHOOK UNIFIED V6.0 + UI OMNICHANNEL (TRAFFIC CONTROLLER)
 ************************************************************/

// =========================================================================
// BAGIAN 1: ROUTER UTAMA (POLISI LALU LINTAS) - SUPER AMAN + SHEET LOG
// =========================================================================
// Nama diubah agar menerima lemparan ID dari klien
function doPostLibrary(e, clientSheetId) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput("NO DATA");
    }

    // TANGKAP PARAMETER URL (Misal: ?device=6285386018255)
    let urlDevice = null;
    if (e.parameter && e.parameter.device) {
        urlDevice = e.parameter.device;
    }

    // 1. Parse data di awal
    let data;
    try {
      data = JSON.parse(e.postData.contents);
    } catch (err) {
      return ContentService.createTextOutput("INVALID JSON");
    }

    if (!data || typeof data !== "object") {
      return ContentService.createTextOutput("DATA NULL");
    }

    // =================================================================================
    // 2. PRE-FILTER & CUSTOM AUTOREPLY JID GRUP (DENGAN SHEET LOG)
    // =================================================================================
    // Toleransi deteksi grup diperluas agar mendeteksi Onesender (isGroup / chat berakhiran @g.us)
    let isGroup = data.is_group === true || data.isGroup === true || String(data.chat_type).toLowerCase() === "group" || String(data.chat || "").includes("@g.us") || String(data.from || "").includes("@g.us");
    
    // Toleransi pembacaan teks pesan (Onesender terkadang menggunakan data.text)
    let pesanTeks = String(data.message || data.message_text || data.text || "").trim().toLowerCase();
    
    // Kunci Command: Hanya sah jika pesannya persis "/group_id"
    let isGroupIdCommand = (pesanTeks === "/group_id");

    // --- FITUR MENDAPATKAN JID GRUP ---
    // Gunakan isGroupIdCommand sebagai syarat pembuka gerbang
    if (isGroup && isGroupIdCommand) {
        // 1. Ekstrak JID Grup secara Universal (Support Starsender v2 group_jid)
        let groupId = data.group_jid || data.chat || data.group_id || (data.is_me ? data.to : data.from);
        
        if (groupId && !String(groupId).includes("@g.us")) {
            groupId = String(groupId) + "@g.us";
        }

        // 2. Ekstrak Nama Grup secara Universal
        let groupName = data.to_group_name || data.group_name || data.push_name || data.participant_name || "Grup Sistem";

        // A. CATAT KE SHEET LOG SECARA PAKSA
        try {
            const ss = SpreadsheetApp.openById(clientSheetId);
            let sheetLog = ss.getSheetByName("LOG");
            if (!sheetLog) {
                sheetLog = ss.insertSheet("LOG");
                sheetLog.appendRow(["Timestamp", "Activity Title", "Detail Pipeline"]);
            }
            sheetLog.appendRow([new Date(), "[INFO GRUP] /group_id Ditekan", "Nama Grup: " + groupName + "\nJID Grup: " + groupId]);
        } catch (logErr) {
            console.error("Gagal nulis sheet LOG: " + logErr.message);
        }

        // B. BALAS KE WHATSAPP (VIA API UNIVERSAL)
        let nomorBot = String(data.to || data.receiver || "").replace(/\D/g, "");
        if (!nomorBot || nomorBot.length < 5) {
            nomorBot = String(data.device || "").split("-").pop().replace(/\D/g, "");
        }

        try {
            const config = getInitializationData(nomorBot, false, clientSheetId);
            if (config.device.waKey && config.device.waUrl) {
                let replyMessage = "WhatsApp ID:\n`" + groupId + "`";
                let universalPayload = {
                    "tujuan": groupId, 
                    "pesan": replyMessage, 
                    "phone": groupId, 
                    "message": replyMessage, 
                    "to": groupId, 
                    "text": replyMessage, 
                    "isGroup": true 
                };

                UrlFetchApp.fetch(config.device.waUrl, {
                    "method": "post",
                    "headers": {
                        "Authorization": config.device.waKey,
                        "Content-Type": "application/json"
                    },
                    "payload": JSON.stringify(universalPayload),
                    "muteHttpExceptions": true
                });
            }
        } catch (e) {
            console.error("Gagal membalas API: " + e.message);
        }
    }

    // ---------------------------------------------------------------------------------
    let isNewsletter = String(data.sender_lid || data.sender || "").includes("@newsletter");
    let isSticker = data.message_type === "sticker" || (data.message && data.message.message_type === "sticker");
    let isBroadcast = String(data.chat || data.to_id || data.from_id || "").includes("@broadcast");
    let isDuplicatePayload = (data.apiUrl === undefined && data.message_id === undefined);

    // 2. MODIFIKASI FILTER: Bypass pesan dari grup HANYA JIKA perintah valid
    let isFinanceCommand = pesanTeks.startsWith("/finance");
    
    // Perbaikan: Tolak pesan grup jika bukan isGroupIdCommand (Admin) dan bukan isFinanceCommand
    if ((isGroup && !isGroupIdCommand && !isFinanceCommand) || isNewsletter || isSticker || isBroadcast || isDuplicatePayload) {
        return ContentService.createTextOutput("IGNORED");
    }

    // 3. Kunci eksekusi
    const lock = LockService.getScriptLock();
    try {
      lock.tryLock(10000);
      if (data.apiUrl !== undefined && data.payload !== undefined) {
        return processUIPost(data, clientSheetId); 
      } else {
        // LEMPAR urlDevice KE FUNGSI processWebhook
        return processWebhook(data, clientSheetId, urlDevice); 
      }
    } catch (err) {
      console.error("ROUTER ERROR: " + err.message);
      logWebhookError("ROUTER ERROR: " + err.message, e.postData.contents, clientSheetId); // <-- Lempar ID
      return ContentService.createTextOutput("ERROR");
      console.error("ROUTER ERROR: " + err.message);
      return ContentService.createTextOutput("ERROR");
    } finally {
      try { lock.releaseLock(); } catch (e) {}
    }

  } catch (fatalError) {
    console.error("FATAL CRASH: " + fatalError.message);
    return ContentService.createTextOutput("FATAL ERROR CAUGHT");
  }
}

// =========================================================================
// BAGIAN 2: FUNGSI APLIKASI UI (FRONTEND & BACKEND OMNICHANNEL)
// =========================================================================
function doGetLibrary(e, clientSheetUrl, clientGasUrl) {
  // Buka file Index.html dari dalam Library Master
  var template = HtmlService.createTemplateFromFile('OmniUI');
  
  // Suntikkan Variabel Klien (URL Sheet & URL GAS) ke dalam HTML
  template.SHEET_URL = clientSheetUrl;
  template.GAS_URL = clientGasUrl;
  
  return template.evaluate()
      .setTitle('Omnichannel Chat Viewer')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1.0, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function processUIPost(data, clientSheetId) { // <-- 1. Tambahkan parameter clientSheetId
  try {
    const { apiUrl, apiKey, payload, fileData, uniqueId, sheetInput } = data;
    let driveLink = "";

    /// 1. PROSES UPLOAD KE DRIVE (Jika ada file dari UI)
    if (fileData && fileData.base64) {
      // 2. Ganti getActiveSpreadsheet menjadi openById(clientSheetId)
      const ssSetting = SpreadsheetApp.openById(clientSheetId).getSheetByName("Setting");
      const FOLDER_ID = String(ssSetting.getRange("C2").getValue()).trim(); 
      
      const folder = DriveApp.getFolderById(FOLDER_ID);
      const base64String = fileData.base64.split(',')[1];
      const decodedData = Utilities.base64Decode(base64String);
      const blob = Utilities.newBlob(decodedData, fileData.mimeType, fileData.name);
      const file = folder.createFile(blob);
      
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      driveLink = "https://drive.google.com/uc?export=view&id=" + file.getId();
      
      if (payload.file) { payload.file = driveLink; } 
      else if (payload.image) { payload.image.link = driveLink; } 
      else if (payload.document) { payload.document.link = driveLink; }
    }

    // 2. KIRIM KE API WHATSAPP
    const response = UrlFetchApp.fetch(apiUrl, {
      method: "post",
      headers: { "Authorization": apiKey, "Content-Type": "application/json" },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    // 3. LOGIKA PENCATATAN KE KOLOM T (HISTORY) JIKA BERHASIL & ADA GAMBAR
    if (response.getResponseCode() == 200 && driveLink !== "" && uniqueId) {
      const ss = SpreadsheetApp.openById(sheetInput.match(/\/d\/([a-zA-Z0-9-_]+)/)[1]);
      const sheet = ss.getSheetByName("Networking");
      const values = sheet.getDataRange().getValues();
      
      for (let i = 1; i < values.length; i++) {
        if (values[i][1] == uniqueId) {
          const now = new Date();
          const timeStr = Utilities.formatDate(now, "GMT+7", "dd/MM HH:mm:ss");
          const newHistoryLog = "\n_\n" + timeStr + "\nadmin :\n[Gambar Disimpan: " + driveLink + "]";
          const currentHistory = values[i][19] || ""; 
          sheet.getRange(i + 1, 20).setValue(currentHistory + newHistoryLog);
          break;
        }
      }
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      data: response.getContentText()
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function getSheetData(sheetInput) { 
  if (typeof OmniBot === 'undefined') return null;
  return OmniBot.getSheetData(sheetInput); 
}

function getAvailableLabels(sheetInput) {
  try {
    var sheetId = sheetInput.match(/\/d\/([a-zA-Z0-9-_]+)/) ? sheetInput.match(/\/d\/([a-zA-Z0-9-_]+)/)[1] : sheetInput;
    var ss = SpreadsheetApp.openById(sheetId);
    var sheetSetting = ss.getSheetByName("Setting");
    if (sheetSetting) {
      return sheetSetting.getRange("F13").getValue().toString(); 
    }
    return "";
  } catch (e) { return ""; }
}

function updateContactData(sheetInput, uniqueId, newName, newLabels, newBiodata) {
  try {
    var sheetId = sheetInput.match(/\/d\/([a-zA-Z0-9-_]+)/) ? sheetInput.match(/\/d\/([a-zA-Z0-9-_]+)/)[1] : sheetInput;
    var ss = SpreadsheetApp.openById(sheetId);
    var sheetInbox = ss.getSheetByName("Networking");
    var data = sheetInbox.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] == uniqueId) { 
        sheetInbox.getRange(i + 1, 8).setValue(newName); 
        sheetInbox.getRange(i + 1, 3).setValue(newLabels);
        sheetInbox.getRange(i + 1, 10).setValue(newBiodata);
        return "success";
      }
    }
    return "not_found";
  } catch (e) { return "error: " + e.message; }
}

function markAsRead(sheetInput, uniqueId) {
  try {
    var sheetId = sheetInput.match(/\/d\/([a-zA-Z0-9-_]+)/) ? sheetInput.match(/\/d\/([a-zA-Z0-9-_]+)/)[1] : sheetInput;
    var ss = SpreadsheetApp.openById(sheetId);
    var sheetInbox = ss.getSheetByName("Networking");
    var data = sheetInbox.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] == uniqueId) {
        sheetInbox.getRange(i + 1, 26).setValue(false);
        return "success";
      }
    }
    return "not_found";
  } catch (e) { return "error: " + e.message; }
}

function pancingIzin() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var folderId = String(ss.getSheetByName("Setting").getRange("C2").getValue()).trim();
  var folder = DriveApp.getFolderById(folderId);
  var file = folder.createFile("test_izin.txt", "ini cuma pancingan");
  file.setTrashed(true); 
  Logger.log("Izin FULL berhasil didapatkan!");
}

// =========================================================================
// BAGIAN 3: FUNGSI INTI WEBHOOK (V6.0 LIGHTWEIGHT)
// =========================================================================
// Tambahkan parameter urlDevice (default null jika tidak diisi)
function processWebhook(data, clientSheetId, urlDevice = null) { 
  try {
    let mappedData = {}, bsuid = "";

    // Deteksi grup universal (Disamakan ketahanannya dengan Bagian 1)
    let isPesanGrup = data.is_group === true || data.isGroup === true || String(data.chat_type).toLowerCase() === "group" || String(data.chat || "").includes("@g.us") || String(data.to_id || "").includes("@g.us") || String(data.from || "").includes("@g.us");

    // 1. PARSING DATA (Dukung format V1 & V2 termasuk Starsender)
    if (data.version || data.message_timestamp) {
      // JIKA urlDevice ADA, UTAMAKAN ITU. JIKA TIDAK, PAKAI DATA DARI JSON.
      let myNumNorm = urlDevice ? normalizePhone(urlDevice) : normalizePhone(data.is_from_me ? data.sender_phone : data.to_id);
      
      bsuid = String(data.sender_lid || "");
      let lampiranUrl = data.attachment_url || data.file || "";
      
      mappedData = {
        nomor_device: myNumNorm, from: normalizePhone(data.sender_phone || data.from || ""),
        is_group: isPesanGrup, is_me: data.is_from_me || false,
        message: data.message_text || data.message || "", message_id: data.message_id || "",
        push_name: !data.is_from_me ? String(normalizePushName(data?.sender_push_name) || ""): "",
        received_at: formatTime(data.message_timestamp), tglformat: getOnlyDate(data.message_timestamp),
        to: normalizePhone(data.to_id || ""), 
        file_url: lampiranUrl,
        message_type: data.message_type || (lampiranUrl ? "media" : "text")
      };
    } else {
      // LAKUKAN HAL YANG SAMA UNTUK FORMAT LAMA
      let myNumNorm = urlDevice ? normalizePhone(urlDevice) : normalizePhone(data.is_me ? data.from : data.to);
      
      bsuid = String(data.sender || ""); 
      let lampiranUrl = data.file || data.attachment_url || "";
      
      mappedData = {
        nomor_device: myNumNorm, from: normalizePhone(data.from || ""),
        is_group: isPesanGrup, is_me: data.is_me || false,
        message: data.message || data.text || "", message_id: data.message_id || "",
        push_name: !data.is_me ? String(normalizePushName(data?.push_name) || ""): "",
        received_at: formatTime(data.received_at), tglformat: getOnlyDate(data.received_at),
        to: normalizePhone(data.to || ""), 
        file_url: lampiranUrl,
        message_type: data.message_type || (lampiranUrl ? "media" : "text")
      };
    }

    // =========================================================================
    // MODIFIKASI PENTING: Izinkan pesan grup masuk Networking JIKA "/group_id"
    // =========================================================================
    let isiPesanFilter = String(mappedData.message).trim().toLowerCase();
    
    let isFinanceProcess = isiPesanFilter.startsWith("/finance");

    // KEMBALIKAN @g.us YANG TERPOTONG KHUSUS UNTUK /group_id DAN /finance
    if (mappedData.is_group && (isiPesanFilter === "/group_id" || isFinanceProcess)) {
        let fullGroupId = data.group_jid || data.chat || data.group_id || (mappedData.is_me ? data.to : data.from);
        if (fullGroupId && !String(fullGroupId).includes("@g.us")) {
            fullGroupId = String(fullGroupId) + "@g.us";
        }
        // Timpa value to/from yang terpotong dengan ID grup utuh
        if (mappedData.is_me) {
            mappedData.to = fullGroupId;
        } else {
            mappedData.from = fullGroupId;
        }
    }

    // Tolak pesan grup KECUALI isinya "/group_id" ATAU diawali "/finance"
    let tolakGrup = mappedData.is_group && isiPesanFilter !== "/group_id" && !isFinanceProcess;

    // Filter Penjaga Kedua
    if (bsuid.toLowerCase().includes("@newsletter") || tolakGrup || (mappedData.message_type || "").toLowerCase() === "sticker") {
      return ContentService.createTextOutput("IGNORED");
    }

    // 2. INIT & SETTINGS
    const isBsuidLookup = (bsuid.toLowerCase().includes("@lid") && mappedData.is_me);
    // Masukkan clientSheetId ke fungsi pengambil data
    const config = getInitializationData(isBsuidLookup ? bsuid : mappedData.nomor_device, isBsuidLookup, clientSheetId);
    
    const ss = SpreadsheetApp.openById(config.targetSpreadsheetId);
    const targetSheet = ss.getSheetByName("Networking");

    let finalColD = config.device.deviceName;
    let finalColE = config.device.deviceNum || mappedData.nomor_device;
    let namaAkun = config.device.akun || "";

    // 3. LOOKUP DATA CUSTOMER
    const allDataBtoG = targetSheet.getRange("B:G").getValues();
    let rowIndex = -1;
    const nomorWA = mappedData.is_me ? mappedData.to : mappedData.from;
    const hasWA = nomorWA && !String(nomorWA).includes("@lid") && String(nomorWA).trim() !== "";
    const uniqueIdB = hasWA ? nomorWA + "." + finalColD : null;

    if (hasWA && uniqueIdB) {
      for (let i = 0; i < allDataBtoG.length; i++) {
        if (String(allDataBtoG[i][0]) === uniqueIdB) { rowIndex = i + 1; break; }
      }
    }
    if (rowIndex === -1 && !mappedData.is_me && bsuid) {
      for (let i = 0; i < allDataBtoG.length; i++) {
        if (String(allDataBtoG[i][5]) === bsuid && String(allDataBtoG[i][2]) === finalColD) { rowIndex = i + 1; break; }
      }
    }

    // BLOK KODE BARU: Ubah const menjadi let agar bisa diperbarui oleh Eager Insertion
    let oldData = getCustomerDataRow(targetSheet, rowIndex); 
    
    if (oldData) {
      const currentMsgId = String(mappedData.message_id);
      const shortId = currentMsgId.slice(-4); 
      if (currentMsgId === String(oldData.msgId) || String(oldData.historyChat || "").includes("[" + shortId + "]")) {
        return ContentService.createTextOutput("DUPLICATE_ID");
      }
    }

   // 4. LABEL & DATA PREP
    let cleanMsg = String(mappedData.message || "").trim().toUpperCase();
    let keywords = String(config.labelwa || "").split(/[,|;]/).map(k => k.trim().toUpperCase()).filter(Boolean);
    let foundKeywordsArr = [];
    
    for (let kw of keywords) { 
        if (new RegExp("\\b" + kw + "\\b").test(cleanMsg)) {
            
            // --- BLOK KODE BARU: PENGAMAN AUTO-LABEL ---
            // Cegah sistem memberikan label "FINANCE" secara otomatis 
            // akibat membaca kata dari perintah komando /finance
            if (kw === "FINANCE" && cleanMsg.startsWith("/FINANCE")) {
                continue; // Lewati siklus ini (Jangan ditambahkan ke array)
            }
            
            foundKeywordsArr.push(kw); 
        } 
    }

    let finalC = "", finalH = "", finalI = oldData ? (oldData.biodata || "") : "";

    if (!oldData) {
      let namaPanggilanWakif = "";
      try {
        let sheetWakif = ss.getSheetByName("🧑🏻 WAKIF");
        if (sheetWakif) {
          let lastRowWakif = sheetWakif.getLastRow();
          if (lastRowWakif >= 2) {
            let dataWakif = sheetWakif.getRange("A2:B" + lastRowWakif).getValues();
            let searchNum = normalizePhone(mappedData.from); 
            for (let i = 0; i < dataWakif.length; i++) {
              if (normalizePhone(dataWakif[i][0]) === searchNum && dataWakif[i][1]) {
                namaPanggilanWakif = String(dataWakif[i][1]).trim(); break; 
              }
            }
          }
        }
      } catch (e) {}

      finalH = namaPanggilanWakif ? namaPanggilanWakif : (mappedData.push_name ? "CEK_" + mappedData.push_name : "");
      let initialLabels = [config.device.autoLabel || "NEW"];
      for(let kw of foundKeywordsArr) { if(!initialLabels.includes(kw)) initialLabels.push(kw); }
      finalC = initialLabels.join(", ");
    } else {
      finalH = oldData.pushName || (mappedData.push_name ? "CEK_" + mappedData.push_name : "");
      let arrayLabelLama = oldData.labelwa ? String(oldData.labelwa).split(",").map(t => t.trim()).filter(Boolean) : [];
      if (foundKeywordsArr.length > 0 && mappedData.is_me) {
        for (let kw of foundKeywordsArr) { if (!arrayLabelLama.includes(kw)) arrayLabelLama.push(kw); }
      }
      finalC = arrayLabelLama.join(", ");
    }


    // ====================================================================
    // FIX: OVERRIDE NAMA & LABEL UNTUK COMMAND /GROUP_ID DARI ADMIN
    // ====================================================================
    if (mappedData.is_group && cleanMsg === "/GROUP_ID") {
        // 1. Override Nama (Kolom H)
        let namaGrupAtauPengirim = data.to_group_name || data.push_name || data.participant_name || data.group_name || "unknown";
        finalH = "Group : " + namaGrupAtauPengirim;

        // 2. Tambahkan Label "GROUP" (Kolom C) jika belum ada
        let arrayLabelGroup = finalC ? finalC.split(",").map(s => s.trim()) : [];
        
        // Mengecek apakah "GROUP" sudah ada di dalam array (Case Insensitive)
        let hasGroupLabel = arrayLabelGroup.some(label => label.toUpperCase() === "GROUP");
        
        if (!hasGroupLabel) {
            arrayLabelGroup.push("GROUP");
        }
        
        // Gabungkan kembali menjadi string dengan koma
        finalC = arrayLabelGroup.join(", ");
    }
    
    // ====================================================================
    // FITUR BARU: OVERRIDE NAMA DARI TANDA KUTIP TUNGGAL ('nama')
    // ====================================================================
    // SYARAT 1: Fitur ini HANYA aktif jika yang mengetik adalah Admin (is_me = true)
    if (mappedData.is_me) {
        let nameMatch = String(mappedData.message || "").match(/'([^']+)'/);
        
        if (nameMatch && nameMatch[1]) {
            let extractedName = nameMatch[1].trim();
            
            // SYARAT 2: Filter Kebersihan (Hanya izinkan huruf, spasi, dan titik)
            // Jika ada Emoji (🟢), Angka, atau Simbol aneh, langsung ditolak!
            let isOnlyLetters = /^[a-zA-Z\s\.]+$/.test(extractedName);
            
            if (isOnlyLetters) {
                // SYARAT 3: Daftar Blacklist (Kata sapaan/tugas yang harus dibuang)
                let blacklistWords = ['hai', 'halo', 'ok', 'oke', 'siap', 'yes', 'done', 'ping', 'do', 'overtime', 'overdue'];
                
                let pureName = extractedName;
                
                // SYARAT 4: Logika Word Boundary (\b) & Case Insensitive (gi)
                // Menghapus kata blacklist dari kalimat, TANPA merusak kata lain (Misal: 'Doke' tidak terpotong)
                blacklistWords.forEach(word => {
                    let regex = new RegExp('\\b' + word + '\\b', 'gi');
                    pureName = pureName.replace(regex, '');
                });
                
                // Bersihkan kelebihan spasi di tengah/awal/akhir akibat kata yang dihapus
                pureName = pureName.replace(/\s+/g, ' ').trim();
                
                // SYARAT 5: Eksekusi penulisan jika sisa namanya valid (2 s/d 25 karakter)
                if (pureName.length >= 2 && pureName.length <= 25) {
                    
                    // Auto-Kapitalisasi (Title Case) -> 'pak galih' jadi 'Pak Galih'
                    let cleanName = pureName.toLowerCase().split(' ').map(w => w.charAt(0).toUpperCase() + w.slice(1)).join(' ');
                    
                    finalH = cleanName; // Timpa Kolom Nama (H) dengan nama yang sudah diekstrak murni
                }
            }
        }
    }

    let qtyInbox = mappedData.is_me ? 0 : 1;
    let qtyOutbox = mappedData.is_me ? 1 : 0;
    if (oldData && mappedData.tglformat === String(oldData.existingM)) {
      qtyInbox += (parseInt(oldData.qtyInbox) || 0);
      qtyOutbox += (parseInt(oldData.qtyOutbox) || 0);
    }

    // ====================================================================
    // BLOK KODE BARU: EAGER INSERTION (ANTI-DUPLIKAT KHUSUS NOMOR BARU)
    // ====================================================================
    // HANYA tereksekusi jika ini nomor baru yang belum ada di Networking
    if (rowIndex === -1) {
        let earlyWaktuMasuk = mappedData.is_me ? "" : mappedData.received_at;
        let earlyWaktuKeluar = mappedData.is_me ? mappedData.received_at : "";
        let earlyBsuid = mappedData.is_me ? "" : bsuid;
        
        // Susun struktur "Baris Draft" 26 Kolom
        let earlyRowData = [
            namaAkun, uniqueIdB, finalC, finalColD, finalColE, (mappedData.is_me ? mappedData.to : mappedData.from), earlyBsuid, finalH, finalI, 
            "", "", "", mappedData.tglformat, mappedData.is_me, mappedData.message_id, 
            earlyWaktuMasuk, earlyWaktuKeluar, qtyInbox, qtyOutbox, 
            "", // History dikosongkan dulu
            JSON.stringify(data), "Queue", "", "", "", true, ""
        ];
        
        // Langsung catat ke database dan kunci paksa dengan flush
        targetSheet.appendRow(earlyRowData);
        SpreadsheetApp.flush(); 
        
        // Ubah rowIndex menjadi baris yang baru saja dibuat
        rowIndex = targetSheet.getLastRow(); 
        
        // Refresh variabel oldData agar sisa script di bawah tidak crash
        oldData = getCustomerDataRow(targetSheet, rowIndex);
    }
    // ====================================================================

    // 5. HISTORY & MEDIA HANDLING
    let historyLama = oldData ? (oldData.historyChat || "") : "";
    let shortTime = mappedData.received_at.substring(8, 10) + "/" + mappedData.received_at.substring(5, 7) + " " + mappedData.received_at.substring(11, 19);
    let shortId = String(mappedData.message_id).slice(-4);
    let pesanFinal = mappedData.message || "";
    let driveLink = "";
    let namaFile = "";

    // =========================================================
// OPTIMASI GAMBAR: BLOB DAN WADAH LAPORAN (LOG)
// =========================================================
let imageBlobGlobal = null;
let imagePipelineLogs = [];

if (mappedData.file_url) {
    const FOLDER_ID = config.folderId;
    let fullUrl = "";
    
    /* PERBAIKAN URL STARSENDER */
    let rawFileUrl = String(mappedData.file_url).trim();
    if (rawFileUrl.startsWith("http://") || rawFileUrl.startsWith("https://")) {
        fullUrl = encodeURI(rawFileUrl);
    } else {
        let prefix = config.device.prefixUrl || "";
        if (prefix.endsWith('/') && rawFileUrl.startsWith('/')) {
            fullUrl = prefix + rawFileUrl.substring(1);
        } else {
            fullUrl = prefix + rawFileUrl;
        }
    }

    let ext = fullUrl.split('.').pop().split(/\#|\?/)[0];
    if (ext.length > 4 || ext.length < 2) ext = "jpg";
    
    // --- BLOK KODE BARU: ANTI DUPLIKAT NAMA FILE ---
    // Membuat stempel waktu absolut berformat YYMMDDHHmmss (Contoh: 260824124739)
    let timeStampUnik = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyMMddHHmmss");
    
    // Format baru: idTanggal_NomorWADevice_Timestamp.jpg
    // Contoh output: id260824_6285386018255.Server_260824124739.jpg
    namaFile = mappedData.tglformat + "_" + uniqueIdB + "_" + timeStampUnik + "." + ext; 
      
      imagePipelineLogs.push("URL TARGET: " + fullUrl);

      // TAHAP 1: FETCH GAMBAR DARI STARSENDER / SERVER WA
      try {
        let responseFetch = UrlFetchApp.fetch(fullUrl, { muteHttpExceptions: true });
        let kodeRespon = responseFetch.getResponseCode();
        
        if (kodeRespon === 200) {
            imageBlobGlobal = responseFetch.getBlob(); 
            let fSize = imageBlobGlobal.getBytes().length;
            let originalType = imageBlobGlobal.getContentType();
            
            // ===============================================================
            // PERBAIKAN FATAL UNTUK STARSENDER (OCTET-STREAM FIX)
            // ===============================================================
            // Paksa semua file yang didownload menjadi tipe MIME gambar JPG
            // Jika tidak, Gemini AI dan Google Drive akan menolak file tersebut
            imageBlobGlobal.setContentType("image/jpeg");
            
            let forcedType = imageBlobGlobal.getContentType();
            
            imagePipelineLogs.push("[SUCCESS] 1. Download dari Server WA. Ukuran: " + fSize + " bytes. Tipe Asli: " + originalType + " -> Dipaksa menjadi: " + forcedType);
        } else {
            let errorText = responseFetch.getContentText().substring(0, 100);
            imagePipelineLogs.push("[FAILED] 1. Download dari Server WA. Kode: " + kodeRespon + ". Response: " + errorText);
        }
      } catch(e) {
        imagePipelineLogs.push("[CRASH] 1. Proses Fetch URL Gagal: " + e.message);
      }

      // TAHAP 2: UPLOAD KE GOOGLE DRIVE
      driveLink = uploadImageToDrive(imageBlobGlobal, namaFile, FOLDER_ID, clientSheetId); // <-- Tambahkan parameter
      
      if (driveLink === "GAGAL_DOWNLOAD") {
        pesanFinal += "\n[Gambar Error: Server Asli (Starsender) Menolak/Gagal Diunduh]";
        imagePipelineLogs.push("[SKIP] 2. Upload GDrive Dibatalkan karena tahap 1 gagal.");
      } else if (driveLink === "ERROR_UPLOAD") {
        pesanFinal += "\n[Gambar Error: Gagal Membuat File di GDrive (Limit/Timeout)]";
        imagePipelineLogs.push("[FAILED] 2. Gagal Create File di GDrive. (Mungkin Rate Limit)");
      } else {
        pesanFinal += "\n[Gambar Disimpan: " + driveLink + "]";
        imagePipelineLogs.push("[SUCCESS] 2. Upload GDrive Berhasil. Link: " + driveLink);
      }
    }

    let newHistoryEntry = "";
    if (oldData && String(oldData.aiLabel).toUpperCase() === "QUEUE" && !mappedData.is_me) {
      newHistoryEntry = "\n" + pesanFinal;
    } else {
      newHistoryEntry = "_\n" + shortTime + " [" + shortId + "]\n" + (mappedData.is_me ? "Admin" : "User") + " : " + pesanFinal;
    }
    
    let historyGabungan = historyLama ? historyLama + (oldData && String(oldData.aiLabel).toUpperCase() === "QUEUE" && !mappedData.is_me ? "" : "\n") + newHistoryEntry : newHistoryEntry;
    let historyUpdatefix = limitHistory(historyGabungan, parseInt(config.limitHistoryValue) || 10);

    // 6. AI LOGIC & URUTAN PENENTUAN STATUS (KOLOM V)
    let finalAiLabel = "Queue", finalAiPrompt = "", finalAiLog = "";
    let finalResumeAi = oldData ? (oldData.resumeAiX || "") : ""; 
    let updateKolomY = false; 

    if (mappedData.is_me && mappedData.message) {
      let arrayLabelAkhir = finalC ? finalC.split(",").map(s => s.trim()).filter(Boolean) : [];
      let firstChar = Array.from(mappedData.message.trim())[0]; 
      if (firstChar === "🙏") {
        if (!arrayLabelAkhir.includes("UNSUBAI")) arrayLabelAkhir.push("UNSUBAI");
      } else if (firstChar === "😊") {
        arrayLabelAkhir = arrayLabelAkhir.filter(label => label.toUpperCase() !== "UNSUBAI");
      }
      finalC = arrayLabelAkhir.join(", ");
    }

    let checkLabel = finalC.toUpperCase();

// ====================================================================
// FITUR BARU: MANUAL INPUT FINANCE (/finance)
// ====================================================================
// 1. Tangkap teks secara agresif (mendukung caption gambar dan teks biasa)
let textForFinance = "";
if (data) textForFinance = data.message_text || data.caption || data.text || data.message || "";
if (!textForFinance) textForFinance = mappedData.message || "";
textForFinance = String(textForFinance).trim();

// 1. Deteksi Command secara terpisah
    let isFinancePure = textForFinance.toUpperCase().startsWith("/FINANCE") && !textForFinance.toUpperCase().startsWith("/FINANCEAI");
    let isFinanceAI = textForFinance.toUpperCase().startsWith("/FINANCEAI") && mappedData.file_url; // Wajib ada gambar

    // 2. SECURITY GATE GLOBAL (Hanya izinkan ID yang memiliki label FINANCE)
    if (isFinancePure || isFinanceAI) {
        let hasFinanceLabel = String(finalC).toUpperCase().includes("FINANCE");
        if (!oldData || !hasFinanceLabel) {
            return ContentService.createTextOutput("IGNORED_UNAUTHORIZED_FINANCE");
        }
    }

    // ====================================================================
    // SKENARIO A: MURNI MANUAL (/finance tanpa AI)
    // ====================================================================
    if (isFinancePure) {
        let getVal = (key) => {
            let regex = new RegExp(key + "\\s*[:=]\\s*(.*)", "i");
            let match = textForFinance.match(regex);
            return match ? match[1].trim() : "";
        };

        let valStatus       = getVal("Status") || "Pending";
        let valNamaKonsumen = getVal("Nama Konsumen") || getVal("Pelanggan"); 
        let valNoWA         = getVal("No WA") || getVal("WA") || getVal("Nomor WA");
        let valAkun         = getVal("Akun");
        let valAgen         = getVal("Agen");
        
        let valJenis    = getVal("Jenis");
        let valKategori = getVal("Kategori");
        let valTgl      = getVal("Tgl Transaksi") || getVal("Tanggal");
        let valNamaItem = getVal("Nama Item");
        let valQty      = getVal("Qty") || getVal("Quantity");
        let valSatuan   = getVal("Satuan");
        let valHarga    = getVal("Harga Satuan") || getVal("Harga");
        let valTotal    = getVal("Total");
        let valNomorStruk = getVal("Nomor Struk") || getVal("Referensi") || getVal("Invoice");
        
        if (valQty)   valQty   = valQty.replace(/\D/g, "");
        if (valHarga) valHarga = valHarga.replace(/\D/g, "");
        if (valTotal) valTotal = valTotal.replace(/\D/g, "");
        
        let finalNamaKonsumen = valNamaKonsumen ? valNamaKonsumen : finalH; 
        let finalUniqueIdB = uniqueIdB; 
        if (valNoWA) {
            let cleanWA = valNoWA.replace(/\D/g, "");
            if (cleanWA.startsWith("0")) cleanWA = "62" + cleanWA.substring(1);
            else if (cleanWA.startsWith("8")) cleanWA = "62" + cleanWA;
            finalUniqueIdB = cleanWA + "." + finalColD;
        }

        let finalAkunFinance = valAkun ? valAkun : namaAkun;
        let finalAgenFinance = valAgen ? valAgen : finalColD;

        let manualDataCSV = valJenis + ";;" + valKategori + ";;" + valTgl + ";;" + valNamaItem + ";;" + valQty + ";;" + valSatuan + ";;" + valHarga + ";;" + valTotal + ";;" + valNomorStruk + ";;" + valStatus;
        let logTambahan = "Input Manual via /finance\n" + imagePipelineLogs.join("\n");

        logImageActivity(finalUniqueIdB, logTambahan, driveLink, manualDataCSV, finalAgenFinance, finalAkunFinance, finalNamaKonsumen, clientSheetId, mappedData, namaFile, textForFinance);

        finalAiLabel = mappedData.is_me ? "Admin Reply" : "Stop"; 
        finalAiLog = "Processed Manual Finance";
        updateKolomY = true;
    }

    // ====================================================================
    // SKENARIO B: ADA GAMBAR (Customer Kirim Gambar ATAU Admin pakai /financeai)
    // ====================================================================
    else if (mappedData.file_url) {
        let aiExtractedData = ""; 
        
        // Eksekusi AI JIKA pengirim adalah Pelanggan (User) ATAU Admin menggunakan /financeai
        if (!mappedData.is_me || isFinanceAI) {
            
            let rawVision = config.device.promptrawVision || "";
            let promptVision = "";

            // LOGIKA PROMPT HYBRID UNTUK MULTI-ITEM
            if (isFinanceAI) {
                promptVision = "### [TUGAS KHUSUS: FINANCE AI HYBRID MULTI-ITEM]\n"+
                "Ekstrak bukti transaksi/struk belanja. Pisahkan nilai menggunakan Double Titik Koma (;;) persis seperti pola berikut:\n"+
                "Status|Jenis;;Kategori;;Tgl Transaksi;;Nama Item (Garis Besar);;Total Qty;;Satuan;;Harga Satuan;;Total;;NomorStrukatauReferensi\n"+
                "ATURAN KHUSUS MULTI-ITEM: Rincikan tiap barang dari struk ke dalam kolom 'Satuan' WAJIB menggunakan format {Nama Barang, Qty, Harga Satuan, Subtotal}. " +
                "Contoh struk dengan 2 barang: {Besi Beton Ulir 13, 30, 142000, 4260000}{Besi Beton Polos 8, 30, 54000, 1620000}. " +
                "PENTING: Jangan gunakan spasi atau koma di antara kurung kurawal penutup dan pembuka }{. Hilangkan tanda titik/koma pada nilai angka harga. " +
                "Kolom 'Total Qty' diisi kuantitas jenis item misal ada ada list besi beton ulir dan besi beton polos maka Qty isi 2, Kolom Harga satuan diisi dengan Rata-rata=Total/Qty, dan Kolom 'Total' diisi grand total tagihan.\n\n"+ rawVision;
            } else {
                promptVision = "### [TUGAS KHUSUS ANALISIS GAMBAR]\n"+
                "Ekstrak data bukti mutasi/transfer/gambar dengan format presisi. Pisahkan setiap nilai menggunakan Double Titik Koma (;;) persis seperti pola berikut:\n"+
                "Status|Jenis;;Kategori;;Tgl Transaksi;;Nama Item;;Qty;;Satuan;;Harga Satuan;;Total;;NomorStruk\n\n" + rawVision;
            }
            
            finalAiPrompt = promptVision;
            
            imagePipelineLogs.push("[INFO] 3. Mengirim Data ke AI Gemini Vision...");
            imagePipelineLogs.push("[INFO] Isi Prompt Vision:\n" + promptVision);
            
            let rawResponse = callGeminiVision(config.geminiApiKey, config.geminiModel, config.fallbackModels, promptVision, mappedData.message, imageBlobGlobal, config.tempGemini, imagePipelineLogs);
            
            let deskripsiGambar = "";
            let aiLabel = "Stop";

            if (rawResponse && (rawResponse.includes("|") || rawResponse.includes(";;"))) {
                deskripsiGambar = rawResponse; 
                if (rawResponse.includes("|")) {
                    let parts = rawResponse.split("|");
                    aiLabel = parts[0].trim();
                    deskripsiGambar = parts.slice(1).join("|").trim(); 
                }
                finalAiLog = rawResponse;
                imagePipelineLogs.push("[SUCCESS] 5. AI Vision Berhasil. Deskripsi: " + deskripsiGambar);
            } else {
                finalAiLog = "Vision Error: " + rawResponse;
                imagePipelineLogs.push("[FAILED] 5. AI Vision Gagal / Respon Tidak Sesuai. Output: " + rawResponse);
            }
            
            updateKolomY = true; 
            
            // --- PECAH JALUR EKSEKUSI PENYIMPANAN ---
            if (isFinanceAI) {
                // JALUR 1: ADMIN HYBRID (Gabungkan AI + Manual Override)
                let getVal = (key) => {
                    let regex = new RegExp(key + "\\s*[:=]\\s*(.*)", "i");
                    let match = textForFinance.match(regex);
                    return match ? match[1].trim() : "";
                };

                // Pecah hasil AI ke Array, pastikan jumlahnya cukup
                let csvParts = deskripsiGambar.split(";;");
                while(csvParts.length < 9) csvParts.push("");

                // Timpa hasil AI jika Admin mengetik nilainya secara manual
                if(getVal("Jenis")) csvParts[0] = getVal("Jenis");
                if(getVal("Kategori")) csvParts[1] = getVal("Kategori");
                if(getVal("Tanggal") || getVal("Tgl Transaksi")) csvParts[2] = getVal("Tanggal") || getVal("Tgl Transaksi");
                if(getVal("Nama Item")) csvParts[3] = getVal("Nama Item");
                if(getVal("Qty")) csvParts[4] = getVal("Qty").replace(/\D/g, "");
                if(getVal("Satuan")) csvParts[5] = getVal("Satuan");
                if(getVal("Harga Satuan") || getVal("Harga")) csvParts[6] = (getVal("Harga Satuan") || getVal("Harga")).replace(/\D/g, "");
                if(getVal("Total")) csvParts[7] = getVal("Total").replace(/\D/g, "");
                if(getVal("Nomor Struk") || getVal("Referensi")) csvParts[8] = getVal("Nomor Struk") || getVal("Referensi");

                // Menentukan Status (Prioritas: Manual -> AI -> Default)
                let hybridStatus = getVal("Status") || aiLabel || "Pending";

                // Gabungkan kembali Array CSV dan tanamkan Status di indeks ke-9
                aiExtractedData = csvParts.slice(0,9).join(";;") + ";;" + hybridStatus;

                // Override Identifier
                let valAkun = getVal("Akun");
                let valAgen = getVal("Agen");
                let valNoWA = getVal("No WA") || getVal("WA") || getVal("Nomor WA");
                let valNamaKonsumen = getVal("Nama Konsumen") || getVal("Pelanggan");

                let finalNamaKonsumen = valNamaKonsumen ? valNamaKonsumen : finalH;
                let finalAkunFinance = valAkun ? valAkun : namaAkun;
                let finalAgenFinance = valAgen ? valAgen : finalColD;
                let finalUniqueIdB = uniqueIdB;

                if (valNoWA) {
                    let cleanWA = valNoWA.replace(/\D/g, "");
                    if (cleanWA.startsWith("0")) cleanWA = "62" + cleanWA.substring(1);
                    else if (cleanWA.startsWith("8")) cleanWA = "62" + cleanWA;
                    finalUniqueIdB = cleanWA + "." + finalColD;
                }

                // Simpan Log ke Finance
                logImageActivity(finalUniqueIdB, "Input Hybrid via /financeai\n" + imagePipelineLogs.join("\n"), driveLink, aiExtractedData, finalAgenFinance, finalAkunFinance, finalNamaKonsumen, clientSheetId, mappedData, namaFile, textForFinance);

                // Hentikan agar AI tidak membalas
                finalAiLabel = "Admin Reply";
                historyUpdatefix += "\n[Admin Eksekusi /financeai]";
            } else {
                // JALUR 2: USER BIASA (Murni AI)
                // Sisipkan AI Label di belakang array (Index ke-9) agar logImageActivity mendeteksinya
                aiExtractedData = deskripsiGambar + ";;" + aiLabel; 
                
                finalAiLabel = (aiLabel.toUpperCase() === "QUEUE") ? "Queue" : "Stop"; 
                historyUpdatefix += "\n[User Kirim Gambar : " + deskripsiGambar + "]";

                let extractedCaption = data.message_text || data.caption || data.text || data.message || "";
                logImageActivity(uniqueIdB, imagePipelineLogs.join("\n"), driveLink, aiExtractedData, finalColD, namaAkun, finalH, clientSheetId, mappedData, namaFile, extractedCaption);
            }

        } else {
            // JALUR 3: ADMIN BYPASS GAMBAR BIASA (Promo/Brosur)
            finalAiLog = "Admin Bypass: Gambar broadcast/promosi tidak masuk Finance.";
            finalAiLabel = "Admin Reply"; 
            historyUpdatefix += "\n[Admin Kirim Gambar]";
            imagePipelineLogs.push("[INFO] 3-5. Gambar dikirim oleh Admin. Eksekusi Vision dan penulisan Finance dilewati.");
            
            updateKolomY = true; 
        }
    }
    
    // URUTAN KONDISI UNTUK TEKS (SKENARIO 2 & 3)
    else {
      // URUTAN 3: Jika is_me = False (Pelanggan yang chat)
      if (!mappedData.is_me) {
        finalAiLabel = "Queue";
        finalAiLog = "Trigger Queue (Text)"; // Nilai disimpan di memori, tapi updateKolomY tetap false

        // Gunakan Logika Stop (UNSUBAI / XX / WARMER) di sini
        let limitWarmer = config.limitWarmerValue;
        let isWarmTarget = (checkLabel.includes("AE") || checkLabel.includes("AER") || checkLabel.includes("AEr"));

        if (checkLabel.includes("UNSUBAI") || checkLabel.includes("XX")) {
            finalAiLabel = "Stop";
            finalAiLog = "Trigger Stop by Label UNSUBAI/XX"; 
        } else if (isWarmTarget && qtyOutbox > limitWarmer) {
            finalAiLabel = "Stop"; 
            finalAiLog = "Limit Warmer " + qtyOutbox + "/" + limitWarmer;
        }
      } 
      // Jika is_me = True (Admin / Bot yang chat keluar)
      else {
        // URUTAN 4: Jika status sebelumnya "AI Reply" -> tulis "AI Reply"
        if (oldData && String(oldData.aiLabel).toLowerCase() === "ai reply") {
            finalAiLabel = "AI Reply";
            finalAiLog = "Processed by AI"; // Nilai disimpan di memori, tapi updateKolomY tetap false
        } 
        // URUTAN 5: Jika is_me = True dan BUKAN "AI Reply" -> "Admin Reply"
        else {
            finalAiLabel = "Admin Reply";
            finalAiLog = "Manual Processed"; // Nilai disimpan di memori, tapi updateKolomY tetap false
        }
      }
    }

    let finalZ = (finalAiLabel === "Admin Reply") ? false : true;

    // ====================================================================
    // FITUR TICKET (UPDATE & CREATE) UNTUK AGEN
    // ====================================================================
    let isTicketUpdateSuccess = false; 

    // PERBAIKAN: Deklarasikan textPesan di sini agar terbaca oleh semua kode di bawahnya
    let textPesan = String(mappedData.message || "").trim();

    // Cek SATU KALI apakah label (Kolom C) mengandung 'AGEN'
    if (String(finalC).toUpperCase().includes("AGEN")) {
        const sheetTicket = ss.getSheetByName("Ticket"); // Panggil sheet 1 kali saja
        
        // Pastikan Sheet Ticket benar-benar ada untuk mencegah eror 'null'
        if (sheetTicket) {
            
            // A. LOGIKA UPDATE PROGRESS TICKET
            // Format Baru: ID - Progress (opsional) - Deskripsi (opsional)
            // Contoh 1: 2605171030 - 20% - Material sudah tiba
            // Contoh 2: 2605171030 - Tukang mulai kerja (Progress otomatis +0.01)
            // Contoh 3: 2605171030 - 50% (Hanya update angka)
            // ----------------------------------------------------------------
            let parts = textPesan.split('-').map(p => p.trim());
            
            if (parts.length >= 1) { // Ubah menjadi minimal 1 karena hanya butuh ID untuk di cek
                let ticketId = parts[0]; 
                let isValidId = /^\d{10}$/.test(ticketId); // Harus pas 10 digit angka
                
                if (isValidId) {
                    const lastRowTicket = sheetTicket.getLastRow();
                    
                    if (lastRowTicket > 1) {
                        // Kolom A = 1 (ID Ticket)
                        const dataIdTicket = sheetTicket.getRange(1, 1, lastRowTicket, 1).getValues();
                        let rowToUpdate = -1;

                        for (let i = 0; i < dataIdTicket.length; i++) {
                            if (String(dataIdTicket[i][0]).trim() === ticketId) {
                                rowToUpdate = i + 1;
                                break;
                            }
                        }

                        if (rowToUpdate !== -1) {
                            // Ambil progress lama terlebih dahulu (Kolom K = 11)
                            let existingProgress = sheetTicket.getRange(rowToUpdate, 11).getValue();
                            let progressDecimal = parseFloat(existingProgress);
                            if (isNaN(progressDecimal)) progressDecimal = 0;
                            
                            let isProgressUpdatedManually = false;
                            let deskripsiBaru = "";
                            let deskripsiStartIndex = 1;

                            // Jika ada bagian kedua setelah ID
                            if (parts.length > 1) {
                                let part1 = parts[1];
                                // Cek apakah part1 adalah persentase/angka progress
                                let isPart1Progress = /^[\d]+%?$/.test(part1.replace(/\s/g, ''));

                                if (isPart1Progress) {
                                    isProgressUpdatedManually = true;
                                    let rawProgress = part1.replace(/[^0-9]/g, ''); 
                                    progressDecimal = parseInt(rawProgress) / 100;
                                    deskripsiStartIndex = 2; // Deskripsi mulai dari part ke-3
                                } else {
                                    // Jika part1 BUKAN angka, berarti itu adalah Deskripsi.
                                    // Progress lama otomatis ditambah 0.01
                                    deskripsiStartIndex = 1;
                                    progressDecimal += 0.01;
                                }
                            }
                            
                            // Gabungkan sisa parts menjadi deskripsi
                            if (parts.length > deskripsiStartIndex) {
                                deskripsiBaru = parts.slice(deskripsiStartIndex).join(' - ').trim();
                            }

                            // Pastikan progress tidak melebihi 1 (100%)
                            if (progressDecimal > 1) progressDecimal = 1;

                            // Tentukan Status Baru
                            let statusBaru = "";
                            if (progressDecimal >= 1) statusBaru = "Done";
                            else if (progressDecimal > 0) statusBaru = "In Progress";
                            else statusBaru = "Plan";

                            // UPDATE 1: Update Angka Progress (Kolom 11)
                            sheetTicket.getRange(rowToUpdate, 11).setValue(progressDecimal);
                            
                            // UPDATE 2: Update Status Ticket (Kolom 9)
                            sheetTicket.getRange(rowToUpdate, 9).setValue(statusBaru);

                            // UPDATE 3: Jika Done, isi waktu selesai (Kolom 19)
                            if (statusBaru === "Done") {
                                let tglDone = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
                                sheetTicket.getRange(rowToUpdate, 19).setValue(tglDone); 
                            }

                            // UPDATE 4: Tambahkan Deskripsi Baru (Kolom 8)
                            if (deskripsiBaru !== "") {
                                let cellDeskripsi = sheetTicket.getRange(rowToUpdate, 8);
                                let deskripsiLama = String(cellDeskripsi.getValue()).trim();
                                let ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");

                                if (deskripsiLama !== "") {
                                    // Deskripsi baru ditaruh paling depan
                                    let updateTeks = "[" + ts + "]\n" + deskripsiBaru + "\n\n" + deskripsiLama;
                                    cellDeskripsi.setValue(updateTeks);
                                } else {
                                    cellDeskripsi.setValue("[" + ts + "]\n" + deskripsiBaru);
                                }
                            }

                            // UPDATE 5: Atur Tanggal (Kolom 16) & Jam (Kolom 17) otomatis ke waktu sekarang + 5 Menit
                            let currentTime = new Date();
                            currentTime.setMinutes(currentTime.getMinutes() + 5); // Tambah 180 menit
                            
                            let timeZone = Session.getScriptTimeZone();
                            let tglPlus180 = Utilities.formatDate(currentTime, timeZone, "yyyy-MM-dd");
                            let jamPlus180 = Utilities.formatDate(currentTime, timeZone, "HH:mm:ss");

                            sheetTicket.getRange(rowToUpdate, 16).setValue(tglPlus180);
                            sheetTicket.getRange(rowToUpdate, 17).setValue(jamPlus180);
                            sheetTicket.getRange(rowToUpdate, 10).setValue(true);
                            
                            isTicketUpdateSuccess = true; // Tandai update sukses agar tidak menjalankan fungsi Create
                        }
                    }
                }
            }

            // ----------------------------------------------------------------
            // B. LOGIKA CREATE TICKET BARU (Hanya jalan jika BUKAN update)
            // ----------------------------------------------------------------
            if (!isTicketUpdateSuccess) {
                let regexCreateTicket = /^(🆕|✅)(?:(\d{4}))?(?:-(\d{2}:\d{2}))?-(.+?)(?:-(.+))?$/is;
                let matchTicket = textPesan.match(regexCreateTicket);

                if (matchTicket) {
                    let emojiStart = matchTicket[1];      
                    let tanggalInput = matchTicket[2]; 
                    let waktuInput = matchTicket[3];
                    let deskripsiProject = matchTicket[4] ? matchTicket[4].trim() : ""; 
                    let catatanTicket = matchTicket[5] ? matchTicket[5].trim() : ""; 
                    
                    let sekarang = new Date();
                    let timeZone = Session.getScriptTimeZone();
                    
                    let randomTicketId = Utilities.formatDate(sekarang, timeZone, "yyMMddHHmm");
                    let akunTicket = (emojiStart === "🆕" || emojiStart === "✅") ? "👷 NusaArtCon" : "Personal";
                    let tglCreateAt = Utilities.formatDate(sekarang, timeZone, "yyyy-MM-dd HH:mm:ss");
                    
                    // PERBAIKAN: Panjang array diubah menjadi 21 (index 0 - 20) agar muat finalH
                    let newTicketRow = new Array(21).fill("");
                    newTicketRow[0] = randomTicketId;     
                    newTicketRow[1] = akunTicket;
                    newTicketRow[4] = "Task";                   
                    newTicketRow[6] = deskripsiProject;   
                    newTicketRow[7] = catatanTicket;      
                    newTicketRow[13] = tglCreateAt;       
                    
                    // ========================================================
                    // PERBAIKAN DISINI: Kolom T diisi Value Kolom B Sheet Networking
                    // ========================================================
                    // Jika data customer sudah ada di Networking, ambil nilai uniknya. 
                    // Jika belum ada, gunakan uniqueIdB pembentukan awal.
                    let valueKolomB = oldData ? oldData.uniqueId : uniqueIdB;
                    
                    newTicketRow[19] = valueKolomB; 
                    newTicketRow[20] = finalH;           
                    
                    if (tanggalInput && tanggalInput.length === 4) {
                        let day = tanggalInput.substring(0, 2);
                        let month = tanggalInput.substring(2, 4);
                        let year = sekarang.getFullYear();
                        
                        newTicketRow[15] = `${year}-${month}-${day}`;
                        newTicketRow[16] = waktuInput ? waktuInput + ":00" : "09:00:00";

                        if (emojiStart === "✅") {
                            newTicketRow[8] = "Done";
                            newTicketRow[10] = 1;
                        } else {
                            newTicketRow[8] = "In Progress";
                            newTicketRow[10] = 0.01;
                        }
                    } else {
                        newTicketRow[15] = Utilities.formatDate(sekarang, timeZone, "yyyy-MM-dd");
                        newTicketRow[16] = Utilities.formatDate(sekarang, timeZone, "HH:mm:ss");

                    if (emojiStart === "✅") {
                            newTicketRow[8] = "Done";
                            newTicketRow[10] = 1;
                        } else {
                            newTicketRow[8] = "In Progress";
                            newTicketRow[10] = 0.01;
                        }
                    }
                    
                    newTicketRow[9] = true;
                    
                    if (deskripsiProject !== "") {
                        sheetTicket.appendRow(newTicketRow);
                    }
                }
            }

            // ========================================================
            // BERHASIL DIGABUNGKAN: LANGSUNG PICU SYNC GOOGLE TASKS
            // ========================================================
            try {
              if (typeof syncSDMToTasks === 'function') {
                syncSDMToTasks();
                Logger.log("Webhook sukses memicu syncSDMToTasks secara langsung.");
              } else if (typeof syncSheetToTasks === 'function') {
                syncSheetToTasks();
                Logger.log("Webhook sukses memicu syncSheetToTasks secara langsung.");
              } else {
                Logger.log("Fungsi syncSDMToTasks tidak ditemukan di project ini.");
              }
            } catch(err) {
              Logger.log("Gagal menjalankan eksekusi langsung syncTasks: " + err.message);
            }
            // ========================================================
            
        }
    }
    // ====================================================================
    // 7. WRITING DATA
    // ====================================================================

    // ====================================================================
    // MODIFIKASI: PISAHKAN WAKTU CHAT MASUK & KELUAR
    // ====================================================================
    let waktuMasukTerakhir = "";
    let waktuKeluarTerakhir = "";

    if (oldData) {
        waktuMasukTerakhir = oldData.message || "";   
        waktuKeluarTerakhir = oldData.receivedAt || ""; 
    }

    if (mappedData.is_me) {
        waktuKeluarTerakhir = mappedData.received_at;
    } else {
        waktuMasukTerakhir = mappedData.received_at;
    }

    /* ======================================================= */
    /* PENGAMAN BSUID: Jangan hapus jika data barunya kosong   */
    /* ======================================================= */
    let bsuidBaru = mappedData.is_me ? "" : bsuid;
    let finalBsuid = bsuidBaru;
    
    // Jika BSUID baru kosong, cek apakah sebelumnya sudah ada BSUID di database (oldData)
    if (!finalBsuid && oldData && oldData.bsuid) {
        finalBsuid = oldData.bsuid; 
    }

    let rowDataUpdate = [
        uniqueIdB, 
        finalC, 
        finalColD, 
        finalColE, 
        (mappedData.is_me ? mappedData.to : mappedData.from), 
        finalBsuid,             // <--- Gunakan variabel finalBsuid yang sudah diamankan (KOLOM G)
        finalH, 
        finalI, 
        "", 
        "", 
        "", 
        mappedData.tglformat, 
        mappedData.is_me, 
        mappedData.message_id, 
        waktuMasukTerakhir,     
        waktuKeluarTerakhir,    
        qtyInbox, 
        qtyOutbox, 
        historyUpdatefix 
    ];

    if (rowIndex === -1) {
        /* Eksekusi baris baru (Data pertama kali masuk): Semua kolom ditulis */
        
        // Cek jika baris baru adalah teks (updateKolomY = false), maka kolom Y dibiarkan kosong ("")
        let nilaiKolomYBaru = updateKolomY ? finalAiLog : ""; 
        
        let aiDataFull = [
            finalAiLabel, 
            finalAiPrompt, 
            finalResumeAi, 
            nilaiKolomYBaru, 
            finalZ, 
            oldData?.jadwalKirim || "", 
        ];
        targetSheet.appendRow([namaAkun, ...rowDataUpdate, JSON.stringify(data), ...aiDataFull]);
    } else {
        /* Eksekusi update baris (Data sudah ada) */
        /* TAMBAHAN: Pastikan Kolom A (indeks 1) juga selalu ter-update dengan Akun */
        targetSheet.getRange(rowIndex, 1).setValue(namaAkun);
        
        /* 1. Update Kolom B sampai T (rowDataUpdate) */
        targetSheet.getRange(rowIndex, 2, 1, rowDataUpdate.length).setValues([rowDataUpdate]);
        
        /* ======================================================= */
        /* TAMBAHAN: Update Kolom U (21) dengan JSON Data Terbaru  */
        /* ======================================================= */
        targetSheet.getRange(rowIndex, 21).setValue(JSON.stringify(data));
        
        /* 2. Update Kolom V (22) - finalAiLabel */
        targetSheet.getRange(rowIndex, 22).setValue(finalAiLabel);

        /* 3. Update Kolom Y (25) - HANYA JIKA updateKolomY bernilai TRUE */
        if (updateKolomY) {
            targetSheet.getRange(rowIndex, 25).setValue(finalAiLog);
        }
        
        /* 4. Update Kolom Z (26) - finalZ */
        targetSheet.getRange(rowIndex, 26).setValue(finalZ);
    }

    return ContentService.createTextOutput("OK");

  } catch (err) {
    // Kirim err.stack beserta 'data' json yang memicu kegagalan
    logWebhookError("WEBHOOK PROCESS ERROR: " + err.stack, data, clientSheetId); // <-- Tambahkan clientSheetId
    return ContentService.createTextOutput("ERROR");
  }
}

// =========================================================================
// BAGIAN 4: FUNGSI PEMBANTU (HELPERS)
// =========================================================================
function getInitializationData(searchKey, isBsuidSearch = false, clientSheetId) { // <-- Tambah Parameter
  // HAPUS getActiveSpreadsheet, langsung pakai clientSheetId
  const result = {
    targetSpreadsheetId: "", folderId: "", geminiApiKey: "", geminiModel: "", 
    tempGemini: "", fallbackModels: "", labelwa: "", limitHistoryValue: 10, limitWarmerValue: 0,
    device: { deviceName: searchKey, deviceNum: "", waKey: "", waUrl: "", prefixUrl: "", aiStatus: "OFF", autoLabel: "" }
  };
  try {
    const sheet = SpreadsheetApp.openById(clientSheetId).getSheetByName("Setting"); // <-- Target Spesifik
    const settingMatrix = sheet.getRange("A1:M30").getValues();
    result.targetSpreadsheetId = String(settingMatrix[1][1]).trim(); 
    result.folderId            = String(settingMatrix[1][2]).trim(); // BACA C2 DISINI
    result.labelwa             = settingMatrix[12][5]; 
    result.tempGemini          = settingMatrix[15][5];  
    result.geminiApiKey        = settingMatrix[17][5]; 
    result.geminiModel         = settingMatrix[17][6]; 
    result.fallbackModels      = settingMatrix[17][7]; 
    result.limitHistoryValue   = settingMatrix[19][4]; 
    result.limitWarmerValue = parseInt(settingMatrix[19][6]) || 0;

    const lastRow = sheet.getLastRow();
    if (lastRow >= 34) {
      const data = sheet.getRange("A34:R" + lastRow).getValues();
      const cleanSearch = String(searchKey).replace(/\D/g, "");
      for (let i = 0; i < data.length; i++) {
        let isMatch = isBsuidSearch ? String(data[i][7]).trim() === searchKey : String(data[i][6]).replace(/\D/g, "") === cleanSearch;
        if (isMatch) {
          result.device = {
            deviceName: data[i][0], 
            akun: String(data[i][1]).trim(),
            deviceNum: String(data[i][6]).replace(/\D/g, ""), 
            waKey: data[i][8], 
            waUrl: data[i][9], 
            prefixUrl: data[i][10], 
            aiStatus: String(data[i][11]).toUpperCase(), 
            promptrawVision: data[i][13], //config.device.promptrawVision
            autoLabel: data[i][4]
          };
          break; 
        }
      }
    }
  } catch (e) { throw e; }
  return result;
}

function getCustomerDataRow(sheet, rowIndex) {
  if (rowIndex === -1) return null;
  const values = sheet.getRange(rowIndex, 1, 1, 28).getValues()[0];
  return {
    timestamp: values[0], uniqueId: values[1], labelwa: values[2], deviceName: values[3], 
    deviceNum: values[4], custNum: values[5], bsuid: values[6], pushName: values[7], 
    biodata: values[8], existingM: values[12], isMe : values[13], msgId: values[14], 
    message: values[15], receivedAt: values[16], qtyInbox: values[17], qtyOutbox: values[18], 
    historyChat: values[19], rawJson: values[20], aiLabel: values[21], aiPrompt: values[22], 
    resumeAiX: values[23], aiLog: values[24], finalZ: values[25], jadwalKirim: values[26], serverKirim: values[27]
  };
}

function normalizePhone(number) {
  if (!number) return "";
  let cleanNumber = String(number).split('@')[0].replace(/\D/g, "");
  return cleanNumber.startsWith("08") ? "628" + cleanNumber.substring(2) : cleanNumber;
}

function normalizePushName(rawName) {
  if (!rawName) return "Kak";
  let nama = String(rawName).replace(/^~/, "").split("|")[0].trim().replace(/[0-9]+$/g, "").replace(/[_\.]/g, " ").replace(/\s+/g, " ").trim();
  if (nama.length < 3 || !/[a-zA-Z]/.test(nama) || /^[^a-zA-Z]+$/.test(nama)) return "Kak";
  return nama.toLowerCase().split(" ").map(w => w.charAt(0).toUpperCase() + w.slice(1)).join(" ");
}

function formatTime(timeStr) {
  try { return Utilities.formatDate(new Date(timeStr), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"); } catch (e) { return timeStr; }
}

function getOnlyDate(timeStr) {
  if (!timeStr || timeStr.length < 10) return "";
  return "id" + timeStr.substring(2, 4) + timeStr.substring(5, 7) + timeStr.substring(8, 10);
}

function limitHistory(historyText, limit) {
  if (!historyText) return "";
  let lines = historyText.split(/\n(?=_)/); 
  return lines.length > limit ? lines.slice(lines.length - limit).join('\n') : historyText;
}

function logWebhookError(message, rawData, clientSheetId) { // <-- Tambah Parameter
  try {
    const ss = SpreadsheetApp.openById(clientSheetId); // <-- Target Spesifik
    
    // Cari sheet bernama "LOG", jika tidak ada maka buat baru
    let sheet = ss.getSheetByName("LOG");
    if (!sheet) {
      sheet = ss.insertSheet("LOG");
      sheet.appendRow(["Timestamp", "Error Message", "Raw JSON Data"]); // Buat header
      sheet.getRange("A1:C1").setFontWeight("bold");
    }

    // Ubah rawData menjadi string agar bisa ditulis ke Spreadsheet
    let dataString = "";
    if (rawData) {
      dataString = typeof rawData === 'object' ? JSON.stringify(rawData) : String(rawData);
    }
    
    // Tulis baris eror ke sheet LOG
    sheet.appendRow([new Date(), message, dataString]);
  } catch (e) {
    // Abaikan jika gagal menulis log (mencegah infinite loop error)
  }
}

// Tambahkan parameter imagePipelineLogs di ujung
function callGeminiVision(apiKey, primaryModel, fallbackStr, systemPrompt, userMessage, imageBlob, temp, imagePipelineLogs) {
  
  // BLOK KODE BARU: Memecah fallback berdasarkan Enter (\n), koma (,), atau pipe (|)
  var cleanPrimary = primaryModel ? String(primaryModel).trim().split(/[\n|,]/)[0].trim() : "";
  var fallbackArr = fallbackStr ? String(fallbackStr).split(/[\n|,]/).map(m => m.trim()) : [];
  
  var modelsToTry = [cleanPrimary].concat(fallbackArr).filter((item, pos, self) => self.indexOf(item) == pos && item !== "");

  // Sensor sebagian API Key untuk keamanan Log
  var maskedKey = apiKey ? apiKey.substring(0, 6) + "..." + apiKey.slice(-4) : "KOSONG";

  try {
    // Jika Blob kosong dari awal, tidak usah proses ke AI
    if (!imageBlob) {
        if(imagePipelineLogs) imagePipelineLogs.push("[ERROR] 4. Eksekusi AI dibatalkan: File gambar gagal diunduh.");
        return "AI No Response: Gambar gagal diunduh dari server WA";
    }

    const payload = {
      "contents": [{ "parts": [
          { "text": "System Instructions: " + systemPrompt + "\n\nUser Message: Analisa gambar ini. " + userMessage },
          // Tidak perlu fetch url lagi, langsung tembak mimeType dan byte Blob
          { "inline_data": { "mime_type": imageBlob.getContentType(), "data": Utilities.base64Encode(imageBlob.getBytes()) } }
      ]}],
      "generationConfig": { "temperature": parseFloat(temp.toString().replace(",", ".")) || 0.2 }
    };
    var options = { method: "POST", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true };
    
    for (var i = 0; i < modelsToTry.length; i++) {
      try {
        var currentModel = modelsToTry[i];
        var res = UrlFetchApp.fetch("https://generativelanguage.googleapis.com/v1beta/models/" + currentModel + ":generateContent?key=" + apiKey, options);
        var resCode = res.getResponseCode();
        var resText = res.getContentText();
        var json = JSON.parse(resText);
        
        if (resCode === 200 && json.candidates) {
            // BERHASIL
            if (imagePipelineLogs) {
                let statusModel = (i === 0) ? "Utama" : "Fallback";
                imagePipelineLogs.push(`[INFO] 4. AI Process (Token: ${maskedKey} | Model ${statusModel}: ${currentModel}) -> BERHASIL`);
            }
            return json.candidates[0].content.parts[0].text;
        } else {
            // GAGAL RESPONSE API
            if (imagePipelineLogs) {
                let errorShort = resText.length > 150 ? resText.substring(0, 150) + "..." : resText;
                imagePipelineLogs.push(`[WARNING] 4. AI Process Gagal menggunakan Model: ${currentModel} (Token: ${maskedKey}). Kode: ${resCode}. Alasan: ${errorShort}`);
            }
        }
      } catch (e) {
         // GAGAL KONEKSI/TIMEOUT
         if (imagePipelineLogs) {
             imagePipelineLogs.push(`[ERROR] 4. Koneksi terputus ke Model: ${modelsToTry[i]}. Alasan: ${e.message}`);
         }
      }
    }
    
    if (imagePipelineLogs) imagePipelineLogs.push("[FAILED] 4. Semua model AI (termasuk fallback) gagal dieksekusi.");
    return "AI No Response";
  } catch (e) { 
      if(imagePipelineLogs) imagePipelineLogs.push(`[CRASH] 4. Error fungsi Vision: ${e.message}`);
      return "Error Vision: Quota/Bandwidth Limit"; 
  }
}

function uploadImageToDrive(imageBlob, fileName, folderId, clientSheetId) { // <-- Tambah parameter
  try {
    // Jika Blob kosong karena gagal download di awal
    if (!imageBlob) return "GAGAL_DOWNLOAD";
    
    // 1. Set nama pada blob terlebih dahulu
    imageBlob.setName(fileName);
    
    // 2. Barulah gunakan blob tersebut untuk membuat file di Drive
    const folder = DriveApp.getFolderById(folderId);
    const file = folder.createFile(imageBlob);
    
    // 3. Set izin akses agar bisa dilihat publik
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return file.getUrl();
  // Weebhook.gs — perbaiki jumlah & urutan argumen agar clientSheetId sampai ke posisi ke-8
  } catch (e) { 
  logImageActivity("Drive Upload Exception", e.message, "", "", "", "", "", clientSheetId); 
  return "ERROR_UPLOAD"; 
  }
}

// Fungsi untuk merekam hasil AI Vision beserta data pelengkap ke Spreadsheet
// BLOK KODE BARU: Penambahan parameter captionText
function logImageActivity(uniqueIdB, pipelineLog, driveLink, aiExtractedData, agen, akun, namaKonsumen, clientSheetId, mappedData, namaFile, captionText) {
  try {
    const ss = SpreadsheetApp.openById(clientSheetId);
    // Asumsi penyimpanan difokuskan pada Sheet Finance sesuai pola URL "Finance_Images"
    const sheetFinance = ss.getSheetByName("Finance"); 
    if (!sheetFinance) return;

   /// 1. Parsing Data dari hasil Gemini AI Vision (Dioptimasi)
    let statusAI = "Pending";
    let jenis = "", kategori = "", namaItem = "", qty = "", satuan = "", harga = "", total = "", tglTransaksiAI = "";
    let statusFinance = ""; 
    let nomorStrukAI = ""; // BLOK KODE BARU: Wadah untuk Nomor Struk AI
    
    if (aiExtractedData) {
      let cleanData = aiExtractedData.split(/\n/)[0].trim();
      let csvParts = cleanData.split(";;");
      
      jenis = csvParts[0] ? csvParts[0].trim() : "";
      kategori = csvParts[1] ? csvParts[1].trim() : "";
      tglTransaksiAI = csvParts[2] ? csvParts[2].trim() : ""; 
      namaItem = csvParts[3] ? csvParts[3].trim() : "";       
      qty = csvParts[4] ? csvParts[4].trim() : "";
      satuan = csvParts[5] ? csvParts[5].trim() : "";
      harga = csvParts[6] ? csvParts[6].trim() : "";
      total = csvParts[7] ? csvParts[7].trim() : "";
      
      // BLOK KODE BARU: Pergeseran Index Pembacaan AI
      nomorStrukAI = csvParts[8] ? csvParts[8].trim() : ""; 
      statusFinance = (csvParts[9] || "").trim() || "Pending"; 
    } else {
      statusFinance = (mappedData && mappedData.is_me) ? "Admin Broadcast" : "Pending";
    }

    // 2. Generate Format Tanggal, Bulan & ID Unik Inv
    const d = new Date();
    const tz = Session.getScriptTimeZone();
    
    const months = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    const yy = String(d.getFullYear()).slice(-2);
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const dd = String(d.getDate()).padStart(2, '0');
    const hh = String(d.getHours()).padStart(2, '0');
    const mnt = String(d.getMinutes()).padStart(2, '0');
    const sec = String(d.getSeconds()).padStart(2, '0');
    
    // BLOK KODE BARU: Format bulan teks murni tanpa tanda strip
    // Output Kolom Q: "2608 Agustus"
    const formatBulan = `${yy}${mm} ${months[d.getMonth()]}`;
    const idUniqInv = `${yy}${mm}${dd}${hh}${mnt}${sec}`; 
    
    // ==============================================================
    // BLOK KODE BARU: VALIDASI KETAT NOMOR STRUK AI (Anti-Halusinasi)
    // ==============================================================
    let finalNomorInvoice = idUniqInv; // Default bawaan saat ini (Backup)
    
    // Syarat 1: AI harus membalas teks (tidak kosong)
    // Syarat 2: Minimal 3 karakter (mencegah balas "-" atau ".")
    // Syarat 3: Harus mengandung Huruf ATAU Angka (Bukan cuma simbol)
    if (nomorStrukAI && nomorStrukAI.length > 2 && /[a-zA-Z0-9]/.test(nomorStrukAI)) {
        
        let lowerStruk = nomorStrukAI.toLowerCase();
        // Syarat 4: Coret kata halusinasi AI yang sering muncul
        let isHalusinasi = lowerStruk.includes("kosong") || lowerStruk.includes("tidak ada") || lowerStruk.includes("null") || lowerStruk.includes("tidak tertera");
        
        if (!isHalusinasi) {
            finalNomorInvoice = nomorStrukAI; // Gunakan hasil AI
        }
    }

    // 3. Mapping Data Payload Lanjutan
    // BLOK KODE BARU: Gunakan captionText yang sudah diekstrak kuat
    const notePesan = captionText ? String(captionText).trim() : "";
    
    // Perbaikan ID Group jika dari mappedData kurang akurat
   let idGroup = "";
   let namaGrup = ""; // Variabel penampung Nama Grup
   
   if (mappedData && mappedData.is_group) {
       idGroup = mappedData.from || mappedData.to || "";
       
       // --- PENCARIAN NAMA GRUP (VLOOKUP INSTAN) ---
       if (idGroup) {
           const sheetNet = ss.getSheetByName("Networking");
           if (sheetNet) {
               // createTextFinder bekerja secepat kilat (0ms) tanpa perlu memuat seluruh isi sheet ke RAM
               const found = sheetNet.createTextFinder(idGroup).matchEntireCell(false).findNext();
               if (found) {
                   // Jika ketemu, ambil nilai dari Kolom H (Indeks 8 = Nama Lengkap/Panggilan)
                   namaGrup = sheetNet.getRange(found.getRow(), 8).getValue();
               }
           }
       }
   }
   
   // --- BLOK KODE BARU: ANTI NAMA FILE PALSU ---
   const namaFileFull = driveLink ? "Finance_Images/" + (namaFile || "Unknown.jpg") : "";

   // 4. Susun Array untuk satu kali penulisan massal (appendRow)
   let newRow = [];
   newRow[0] = Utilities.formatDate(d, tz, "dd/MM/yyyy HH:mm:ss"); // Kolom A
   newRow[1] = akun; // Kolom B
   newRow[2] = agen;  // Kolom C
   newRow[3] = uniqueIdB; // Kolom D
   newRow[4] = namaKonsumen; // Kolom E
   newRow[5] = statusFinance;  // Kolom F (Status)
   newRow[6] = jenis;  // Kolom G
   newRow[7] = kategori; // Kolom H
   newRow[8] = tglTransaksiAI; // Kolom I 
   newRow[9] = namaItem; // Kolom J
   newRow[10] = qty; // Kolom K
   newRow[11] = satuan; // Kolom L
   newRow[12] = harga; // Kolom M
   newRow[13] = total; // Kolom N
  
   // --- PERGESERAN ARRAY KOLOM O SAMPAI Z ---
   newRow[14] = notePesan;    // Kolom O (Note)
   newRow[15] = idGroup;      // Kolom P (ID Grup)
   newRow[16] = namaGrup;     // Kolom Q (Nama Grup) <--- BARU MASUK SINI
   newRow[17] = finalNomorInvoice;    // Kolom R (No Inv Unik)
   newRow[18] = driveLink;    // Kolom S (Link Drive)
   newRow[19] = namaFileFull; // Kolom T (Nama File)
   newRow[20] = pipelineLog;  // Kolom U (Proses AI) 
   newRow[21] = formatBulan;  // Kolom V (Bulan)

   // --- PERSIAPAN DATABASE UNTUK UI OMNICHANNEL ---
   // newRow[22] = "";    // Kolom W (Metode Pembayaran)
   // newRow[23] = "";    // Kolom X (Tgl Verifikasi)
   // newRow[24] = "";    // Kolom Y (Verifikator)
   // newRow[25] = false; // Kolom Z (Is_Deleted) default false

   // 5. Eksekusi Tulis ke Sheet klien
   sheetFinance.appendRow(newRow);

  } catch (err) {
    console.error("Gagal saat memproses logImageActivity: " + err.message);
  }
}

// =========================================================================
// ADAPTOR JALUR DALAM UNTUK UI OMNICHANNEL
// =========================================================================
function processUIPostNative(data) {
  // Ekstrak ID dari URL sheetInput yang dikirim dari UI HTML
  let clientSheetId = data.sheetInput;
  let match = data.sheetInput.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (match) { clientSheetId = match[1]; }

  // 1. Eksekusi fungsi pengirim yang sudah ada dengan menyertakan ID
  let resultObject = processUIPost(data, clientSheetId);
  
  // 2. Ubah tipe datanya menjadi teks biasa agar bisa ditangkap oleh UI HTML
  return resultObject.getContent();
}
