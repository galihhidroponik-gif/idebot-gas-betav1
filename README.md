---
name: omnichannel-gas-guidelines
description: Behavioral guidelines and technical architecture for developing, refactoring, and maintaining the Omnichannel AI WhatsApp SaaS system. Integrates Karpathy's core principles with GAS technical guardrails and explicit execution confirmation gates, including a standard Professional Confirmation Modal pattern.
license: MIT
---

# Omnichannel AI (GAS) Guidelines - Karpathy Framework Edition

Behavioral guidelines and technical architecture rules to reduce LLM coding mistakes, overengineering, and quota/performance issues when working on the **Omnichannel AI WhatsApp SaaS (Google Apps Script & Modular Front-End UI)** project.

**Tradeoff:** These guidelines bias toward caution, quota efficiency, and architectural stability over speed. For trivial fixes, use judgment.

Spesialisasi tingkat lanjut kamu meliputi:
- **Google Apps Script (GAS)** & Arsitektur Library Master vs Client Script (`code.gs`).
- **Optimasi Database Google Sheets** (Batch matrix `getValues`/`setValues`, `CacheService`, `PropertiesService`, `LockService`).
- **Pengembangan UI Web App** (Vanilla JS Modular, Asinkronus `google.script.run`, Fast-DOM Rendering).
- **Integrasi API WhatsApp & Webhook Engine** (Starsender, Fonnte, Wablas, OneSender, WA Business API, Handling Traffic Heavy Batch & Retry Queue).
- **Workspace Tooling**: Keuangan/Invoice, Task/KPI/Area Ticket, Bank Konten, & Management Device.

---

## 🗺️ System Architecture & File Map (`/home/galihprm/agy/project/omnichanel`)

1. **Backend Engine (Library Master):**
   - `webhook.gs`: Webhook receiver, message router, auto-reply logic, AI handler.
   - `actionkirim.gs`: Minute trigger/cron engine, batch message queue processor, rate-limiting, auto-retry.
2. **Backend Bridge / Client Script:**
   - `code.gs`: Web App entry point (`doGet`/`doPost`), HTML partial inclusions, native bridges to Spreadsheet DB & Library Master.
3. **Modular Front-End UI Web App (`omniUI`):**
   - `omniui.html`: Frame master, CDN inclusion, global scope variables (`SHEET_URL`, `GAS_URL`).
   - `omniui_css.html`: Layout 100dvh, responsive mobile drawer, chat bubbles, print styling.
   - `omniui_sidebar.html`: Contact list, search bar, device/label filters.
   - `omniui_chat.html`: Chat header, message container, quick actions, input box.
   - `omniui_modals.html`: Popups (Login, Profile, Finance, Task/KPI, Bank Konten, Settings).
   - `omniui_js_app.html`: Core JS state, timer sync (15s/60s), chat/sidebar renderers, dispatcher.
   - `omniui_js_modules.html`: Workspace modules (Finance & Invoice, Task & KPI, Bank Konten).

---

## ⚡ Quick Lookup Table (File Modification Map)

| Feature / Task Focus | Primary File to Modify | Secondary File |
| :--- | :--- | :--- |
| **Sidebar / Contact List View** | `omniui_sidebar.html` | `omniui_css.html` |
| **Chat Area & Message Input UI** | `omniui_chat.html` | `omniui_css.html` |
| **Chat Logic, Sending & Rendering** | `omniui_js_app.html` (`renderChat`, `triggerSend`) | `code.gs` |
| **Login Auth & Auto-Sync Timers** | `omniui_js_app.html` (`doLogin`, `startSyncTimers`) | `code.gs` |
| **Finance & Invoice Generator** | `omniui_modals.html` & `omniui_js_modules.html` | `code.gs` |
| **Task, KPI & Area Ticket** | `omniui_modals.html` & `omniui_js_modules.html` | `code.gs` |
| **Bank Konten / Static Replies** | `omniui_modals.html` & `omniui_js_modules.html` | `code.gs` |
| **Device Config & WhatsApp API** | `omniui_modals.html` & `omniui_js_app.html` | `code.gs` / `webhook.gs` |
| **Incoming Webhook & AI Routing** | `webhook.gs` | `code.gs` |
| **Batch Message Sender / Cron** | `actionkirim.gs` | `code.gs` |
| **Backend Bridge & Sheet DB Call** | `code.gs` | Google Sheets DB |

---

## 🧠 1. Think Before Coding (Evaluasi & Permukaan Asumsi)

**Jangan berasumsi sepihak. Sampaikan limitasi Google Apps Script (GAS) & evaluasi kriteria sebelum menulis kode.**

- **Evaluasi Quota & Timeout GAS:**
  - Kuota `UrlFetchApp`: Maksimal 20.000 panggilan/hari.
  - Batas waktu eksekusi skrip: Maksimal 6 menit per running script.
  - Sebelum mengimplementasikan fitur baru, periksa dampak terhadap batas kuota dan waktu eksekusi.
- **Sampaikan Asumsi & Tradeoff:**
  - Jika instruksi pengguna memiliki beberapa opsi implementasi (misal: client-side filtering vs backend sheet query), sajikan kelebihan & kekurangannya.
  - Jangan secara implisit memilih cara berat jika cara yang lebih ringan dan hemat resource tersedia.
- **Push Back Jika Perlu:**
  - Jika permintaan dapat membekukan UI WebApp (misal: polling berlebihan < 5 detik) atau berpotensi membuat *race condition* di Spreadsheet, sampaikan risikonya dan berikan saran solusi alternatif.
- **Stop Saat Ragu:**
  - Jika struktur data sheet atau alur fungsi tidak pasti, jangan menebak. Tanyakan atau periksa file rujukan terlebih dahulu.

---

## ⚡ 2. Simplicity First (Utamakan Kesederhanaan Kode)

**Tulis kode minimal yang menyelesaikan masalah. Tanpa overengineering atau fitur spekulatif.**

- **Vanilla JS & Native GAS:** Gunakan Vanilla JS murni dan API native Apps Script tanpa menambah framework/library eksternal berlebih.
- **Tanpa Abstraksi Berlebihan:** Jangan membuat fungsi helper / wrapper bertingkat untuk panggilan `google.script.run` yang hanya digunakan sekali.
- **Cegah Fitur Spekulatif:** Jangan menambahkan opsi konfigurasi atau fleksibilitas ekstra yang tidak diminta pengguna.
- **Uji Kesederhanaan:** Jika fungsi di `code.gs` atau `omniui_js_app.html` mencapai 150 baris padahal bisa diselesaikan dalam 40 baris dengan struktur lebih bersih, lakukan penyederhanaan (*refactor*).

---

## ✂️ 3. Surgical Changes & Technical Guardrails (Perubahan Presisi & Aturan GAS)

**Sentuh hanya file yang relevan. Jaga stabilitas arsitektur & bersihkan sampah sendiri.**

- **Rujukan Presisi:** Hanya edit file yang terdaftar dalam *Quick Lookup Table* sesuai skop tugas.
- **Single Source of Truth Variabel Global:**
  - Deklarasi `SHEET_URL` dan `GAS_URL` **HANYA di `omniui.html`**.
  - Dilarang mendefinisikan ulang variabel tersebut dengan `let` / `const` di file partial (`omniui_js_app.html`, dll) karena akan menyebabkan `Uncaught SyntaxError`.
- **Batch Matrix Reading/Writing (Cegah Rate Limit Sheet):**
  - Dilarang melakukan `getValue()` / `setValue()` di dalam loop (`for`/`forEach`).
  - Gunakan selalu batch reading `sheet.getRange(...).getValues()` dan batch writing `sheet.getRange(...).setValues(...)`.
- **Kompatibilitas Client Script (`code.gs`):**
  - **Dilarang mengubah nama fungsi atau signature parameter** pada fungsi di `code.gs` atau Library Master yang sedang aktif dipanggil oleh akun klien (*No breaking changes*).
- **Kebersihan Kode (Clean Up Your Own Mess):**
  - Hapus variabel, fungsi, atau import yang tidak lagi terpakai akibat perubahan yang *Anda* buat.
  - Jangan menghapus atau mengubah kode mati bawaan proyek (*pre-existing dead code*) kecuali diminta eksplisit.

---

## 🛑 4. Interactive Confirmation & Execution Approval (Protokol Persetujuan "Eksekusi")

**Dilarang mengeksekusi perubahan kode secara langsung sebelum mendapatkan konfirmasi eksplisit dari pengguna.**

- **Jelaskan Rencana Perubahan Dahulu (Resume Plan):**
  Sebelum melakukan edit file atau mengubah kode, AI Agent **wajib** menjelaskan rencana perubahan secara terstruktur:
  1. File target yang akan diubah & alasan perubahan.
  2. Rangkuman/resume logika & perbandingan kode lama vs kode baru.
  3. Evaluasi risiko & dampaknya terhadap kuota/performa GAS.
- **Tanyakan Konfirmasi Eksplisit di Akhir Respon:**
  Di bagian paling akhir dari respon rencana, AI Agent **wajib** mencantumkan kalimat konfirmasi berikut:
  > *"Jika Anda menyetujui rencana perubahan skrip di atas, silakan ketik **`Eksekusi`** untuk melanjutkan ke proses pengubahan kode."*
- **Gerbang Penahanan (Hold Gate):**
  - Jika pengguna **BELUM** memberikan pesan/command eksplisit **"Eksekusi"**, AI Agent **TIDAK BOLEH** mengeksekusi pengubahan file (`edit_file` / `write_file`).
  - AI Agent harus tetap hanya menampilkan/menyempurnakan resume rencana perubahan skrip sampai perintah **"Eksekusi"** diberikan oleh pengguna.

---

## 🎯 5. Goal-Driven Execution & Protocol (Siklus Eksekusi Setelah Disetujui)

**Setelah pengguna memberikan command "Eksekusi", jalankan siklus eksekusi terkontrol:**

```
1. [Auto-Backup Snapshot]   → verify: Folder Backup/<YYYYMMDD_HHMMSS> terbuat & file tersimpan.
2. [Inspect Target Lines]   → verify: Teks target dicocokkan 100% via read_file.
3. [Surgical Code Edit]    → verify: Perubahan presisi tanpa merusak fungsi di sekitarnya.
4. [Guardrail Verification] → verify: Tidak ada redeclaration variabel global & signature code.gs aman.
5. [Report & Rollback Cmd] → verify: Sertakan tautan baris persis & perintah restore bash 1-baris.
```

### Format Laporan Hasil Pekerjaan (Setelah Eksekusi)

Setelah eksekusi selesai, berikan laporan terstruktur:
1. 🔍 **Analisis & Akar Masalah:** Penjelasan logis perubahan & dampaknya terhadap kuota/performa GAS.
2. 🛠️ **File & Snapshot Backup:** Path file yang diubah dan lokasi backup snapshot.
3. 💻 **Ringkasan Kode & Line Tracking:** Tautan file + rentang baris persis (contoh: [`code.gs#L45-L60`](file:///home/galihprm/agy/project/omnichanel/code.gs#L45-L60)).
4. 🔄 **Perintah Rollback Instan:** Perintah `cp` bash 1-baris untuk me-restore file backup jika diperlukan.


---

## 🏗️ 6. Standardized UI Confirmation Protocol (Professional Modal)

To ensure consistency in UI/UX across all confirmation actions (Logout, Delete, Duplicate Warn), **always** use the standardized `showConfirm` modal pattern instead of native `confirm()` or `alert()`.

### 📋 Modal Template (`omniui_modals.html`)
```html
<div id="global-confirm-modal" class="fixed inset-0 bg-black/60 flex items-center justify-center p-4 z-[999999] hidden backdrop-blur-sm">
    <div class="bg-white rounded-2xl w-full max-w-sm shadow-2xl p-6 transform transition-all scale-100 border border-gray-100">
        <div class="flex items-center gap-3 mb-3">
            <div id="confirm-icon-bg" class="w-10 h-10 rounded-full bg-blue-50 flex items-center justify-center shrink-0">
                <span id="confirm-icon" class="text-xl">⚠️</span>
            </div>
            <h3 id="confirm-title" class="text-base font-extrabold text-gray-800 leading-snug">Konfirmasi</h3>
        </div>
        <p id="confirm-message" class="text-xs text-gray-600 mb-6 leading-relaxed"></p>
        <div class="flex gap-2.5">
            <button id="btn-confirm-cancel" class="flex-1 py-2.5 rounded-xl text-xs font-bold bg-gray-100 text-gray-600 hover:bg-gray-200 transition-colors">Batal</button>
            <button id="btn-confirm-ok" class="flex-1 py-2.5 rounded-xl text-xs font-bold bg-[#0055ff] text-white hover:bg-[#0044cc] shadow-md transition-colors">Ya, Lanjutkan</button>
        </div>
    </div>
</div>
```

### ⚡ Global JavaScript Trigger (`omniui_js_app.html`)
```javascript
function showConfirm(title, message, onConfirm, icon = '⚠️') {
    const modal = document.getElementById('global-confirm-modal');
    if (!modal) {
        if (confirm(message)) onConfirm();
        return;
    }

    document.getElementById('confirm-title').innerText = title || "Konfirmasi";
    document.getElementById('confirm-message').innerText = message || "";
    const iconEl = document.getElementById('confirm-icon');
    if (iconEl) iconEl.innerText = icon;

    const okBtn = document.getElementById('btn-confirm-ok');
    const cancelBtn = document.getElementById('btn-confirm-cancel');

    const newOkBtn = okBtn.cloneNode(true);
    const newCancelBtn = cancelBtn.cloneNode(true);
    okBtn.parentNode.replaceChild(newOkBtn, okBtn);
    cancelBtn.parentNode.replaceChild(newCancelBtn, cancelBtn);

    newOkBtn.onclick = () => {
        modal.classList.add('hidden');
        if (typeof onConfirm === 'function') onConfirm();
    };
    newCancelBtn.onclick = () => {
        modal.classList.add('hidden');
    };

    modal.classList.remove('hidden');
}
```

### 🎯 Usage Pattern
```javascript
showConfirm("Hapus Data", "Yakin ingin menghapus item ini secara permanen?", () => {
    // Eksekusi aksi penghapusan di sini
}, '🗑️');
```
