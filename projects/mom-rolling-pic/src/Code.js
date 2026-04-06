// ============================================================
// STANDUP MOM ROLLER v3
// Fitur: Random Rolling, Google Chat, Email, WhatsApp (Fonnte), MOM Docs
// Pre-Meeting Reminders: WhatsApp notification before meeting
// Jadwal: Senin, Rabu, Jumat
// Update: Add WhatsApp reminder + Update Harian tracking sheet
// ============================================================

var SHEET_TIM     = "Tim";
var SHEET_JADWAL  = "Jadwal";
var SHEET_CONFIG  = "Config";
var SHEET_UPDATE  = "Update Harian";


// ============================================================
// MENU CUSTOM
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Standup Roller")
    .addItem("1. Setup Awal (Jalankan Pertama Kali)", "setupSheets")
    .addItem("2. Aktifkan Trigger Otomatis", "setupTrigger")
    .addSeparator()
    .addItem("Test Rolling Sekarang", "rollAssignment")
    .addItem("Test WhatsApp Reminder", "testWhatsAppReminder")
    .addItem("Test Buat MOM Doc Saja", "testCreateDoc")
    .addItem("Cek Folder ID", "cekFolderId")
    .addSeparator()
    .addItem("Lihat History Jadwal", "openJadwal")
    .addToUi();
}


// ============================================================
// SETUP SHEETS
// ============================================================
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  // Sheet: Tim
  var sheetTim = ss.getSheetByName(SHEET_TIM);
  if (!sheetTim) { sheetTim = ss.insertSheet(SHEET_TIM); }
  sheetTim.clearContents();
  sheetTim.clearFormats();

  var timHeader = sheetTim.getRange("A1:B1");
  timHeader.setValues([["Nama", "Email"]]);
  timHeader.setFontWeight("bold")
           .setBackground("#1a73e8")
           .setFontColor("white")
           .setHorizontalAlignment("center");

  sheetTim.getRange("A2:B4").setValues([
    ["Budi Santoso",  "budi@email.com"],
    ["Sari Dewi",     "sari@email.com"],
    ["Ahmad Rizky",   "ahmad@email.com"]
  ]);
  sheetTim.autoResizeColumns(1, 2);

  // Sheet: Jadwal
  var sheetJadwal = ss.getSheetByName(SHEET_JADWAL);
  if (!sheetJadwal) { sheetJadwal = ss.insertSheet(SHEET_JADWAL); }
  sheetJadwal.clearContents();
  sheetJadwal.clearFormats();

  var jadwalHeader = sheetJadwal.getRange("A1:F1");
  jadwalHeader.setValues([["Tanggal", "Hari", "Nama PIC", "Email PIC", "Link MOM Doc", "Timestamp"]]);
  jadwalHeader.setFontWeight("bold")
              .setBackground("#34a853")
              .setFontColor("white")
              .setHorizontalAlignment("center");
  sheetJadwal.autoResizeColumns(1, 6);

  // Sheet: Update Harian (NEW)
  var sheetUpdate = ss.getSheetByName(SHEET_UPDATE);
  if (!sheetUpdate) { sheetUpdate = ss.insertSheet(SHEET_UPDATE); }
  sheetUpdate.clearContents();
  sheetUpdate.clearFormats();

  var updateHeader = sheetUpdate.getRange("A1:I1");
  updateHeader.setValues([["Nama", "Email", "Selesai Kemarin", "Akan Dikerjakan", "Bug Report", "Status Bug (Prod/Stg/Dev)", "Test Exec Status", "Update Jira", "Last Update"]]);
  updateHeader.setFontWeight("bold")
              .setBackground("#f9ab00")
              .setFontColor("white")
              .setHorizontalAlignment("center")
              .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  sheetUpdate.setColumnWidth(1, 150);  // Nama
  sheetUpdate.setColumnWidth(2, 200);  // Email
  sheetUpdate.setColumnWidth(3, 180);  // Selesai Kemarin
  sheetUpdate.setColumnWidth(4, 180);  // Akan Dikerjakan
  sheetUpdate.setColumnWidth(5, 150);  // Bug Report
  sheetUpdate.setColumnWidth(6, 120);  // Status Bug
  sheetUpdate.setColumnWidth(7, 120);  // Test Exec
  sheetUpdate.setColumnWidth(8, 100);  // Update Jira
  sheetUpdate.setColumnWidth(9, 140);  // Last Update

  // Sheet: Config
  var sheetConfig = ss.getSheetByName(SHEET_CONFIG);
  if (!sheetConfig) { sheetConfig = ss.insertSheet(SHEET_CONFIG); }
  sheetConfig.clearContents();
  sheetConfig.clearFormats();

  var configData = [
    ["Google Chat Webhook URL",            "PASTE_WEBHOOK_URL_DISINI"],
    ["Email CC (pisahkan dengan koma)",    "manager@email.com"],
    ["Google Meet Link Recurring",         "https://meet.google.com/xxx-xxxx-xxx"],
    ["Nama Tim",                           "QE INADigital"],
    ["Jam Notifikasi (angka saja)",        "8"],
    ["Google Drive Folder ID untuk MOM",   "PASTE_FOLDER_ID_DISINI"],
    ["WhatsApp Group ID (atau kosongkan)", ""],
    ["Fonnte Token",                       ""],
    ["Reminder (menit sebelum meeting)",   "30"]
  ];

  sheetConfig.getRange("A1:B9").setValues(configData);
  sheetConfig.getRange("A1:A9").setFontWeight("bold").setBackground("#f8f9fa");
  sheetConfig.getRange("A1:B9").setBorder(
    true, true, true, true, true, true,
    "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID
  );
  sheetConfig.setColumnWidth(1, 280);
  sheetConfig.setColumnWidth(2, 350);

  // Add notes for new config
  sheetConfig.getRange("A7").setNote("WhatsApp Group ID (format: 120363xxx@g.us)\nKosongkan jika kirim individual ke setiap anggota tim");
  sheetConfig.getRange("A8").setNote("Token API dari Fonnte.com");
  sheetConfig.getRange("A9").setNote("Berapa menit sebelum meeting untuk kirim reminder WhatsApp\nContoh: 30 = reminder 30 menit sebelum");

  // Hapus sheet WhatsApp Log jika ada
  var sheetWA = ss.getSheetByName("WhatsApp Log");
  if (sheetWA) { ss.deleteSheet(sheetWA); }

  ui.alert(
    "Setup Selesai!",
    "Langkah selanjutnya:\n\n" +
    "1. Isi data tim di sheet 'Tim' (Nama & Email)\n" +
    "2. Lengkapi sheet 'Config' (termasuk WhatsApp config baru)\n" +
    "3. Klik menu > Aktifkan Trigger Otomatis\n\n" +
    "Fitur Baru v3:\n" +
    "- WhatsApp reminder sebelum meeting\n" +
    "- Sheet 'Update Harian' untuk tracking pre-meeting\n" +
    "- MOM Doc auto-filled dari sheet Update Harian\n\n" +
    "Tips:\n" +
    "- Folder ID: Buka folder di Drive > lihat URL > salin ID setelah /folders/\n" +
    "- WhatsApp Group ID: Format 120363xxx@g.us\n" +
    "- Fonnte Token: Daftar di fonnte.com",
    ui.ButtonSet.OK
  );
}


// ============================================================
// SETUP TRIGGER - 2 TRIGGERS (Pre-Meeting + Meeting Time)
// ============================================================
function setupTrigger() {
  var ui = SpreadsheetApp.getUi();

  // Delete all existing triggers
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetConfig = ss.getSheetByName(SHEET_CONFIG);
  var jamNotif = parseInt(sheetConfig.getRange("B5").getValue()) || 8;
  var reminderMinutes = parseInt(sheetConfig.getRange("B9").getValue()) || 30;

  // Calculate reminder time (meeting time - reminder minutes)
  // Example: Meeting at 11:00 (hour 11), reminder 30 min before = trigger at 10:30
  var reminderHour = jamNotif;
  var reminderMinute = 0;

  if (reminderMinutes >= 60) {
    reminderHour = jamNotif - Math.floor(reminderMinutes / 60);
    reminderMinute = 60 - (reminderMinutes % 60);
  } else {
    if (reminderMinutes > 0) {
      reminderHour = jamNotif - 1;
      reminderMinute = 60 - reminderMinutes;
    }
  }

  // Normalize if reminder hour is negative (e.g., meeting at 8:00, reminder 120 min = 6:00)
  if (reminderHour < 0) {
    reminderHour = 23 + reminderHour;
  }

  // TRIGGER 1: Pre-Meeting Reminder (WhatsApp to all team)
  ScriptApp.newTrigger("sendPreMeetingReminder")
    .timeBased()
    .everyDays(1)
    .atHour(reminderHour)
    .nearMinute(reminderMinute)
    .create();

  // TRIGGER 2: Rolling at meeting time (create MOM doc + notify)
  ScriptApp.newTrigger("rollAssignment")
    .timeBased()
    .everyDays(1)
    .atHour(jamNotif)
    .create();

  var reminderTime = (reminderHour < 10 ? "0" : "") + reminderHour + ":" + (reminderMinute < 10 ? "0" : "") + reminderMinute;
  var meetingTime = (jamNotif < 10 ? "0" : "") + jamNotif + ":00";

  ui.alert(
    "Trigger Aktif!",
    "Rolling otomatis berjalan setiap:\n" +
    "- Senin, Rabu, Jumat\n\n" +
    "JADWAL TRIGGER:\n" +
    "1. Pre-Meeting Reminder (WhatsApp): " + reminderTime + "\n" +
    "   - Kirim reminder ke tim\n" +
    "   - Tim isi 'Update Harian' sheet\n\n" +
    "2. Meeting Time (Rolling): " + meetingTime + "\n" +
    "   - Random pick PIC\n" +
    "   - Buat MOM Doc (auto-filled)\n" +
    "   - Kirim Google Chat + Email\n\n" +
    "(Hari lain otomatis di-skip)",
    ui.ButtonSet.OK
  );
}


// ============================================================
// PRE-MEETING REMINDER (WhatsApp)
// ============================================================
function sendPreMeetingReminder() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var today = new Date();
  var day   = today.getDay();

  // Hanya Senin(1), Rabu(3), Jumat(5)
  if (day !== 1 && day !== 3 && day !== 5) {
    Logger.log("Bukan hari standup (" + day + "), skip pre-meeting reminder.");
    return;
  }

  // Ambil Config
  var sheetConfig  = ss.getSheetByName(SHEET_CONFIG);
  var configValues = sheetConfig.getRange("B1:B9").getValues();
  var meetLink     = configValues[2][0].toString();
  var timName      = configValues[3][0].toString();
  var jamNotif     = parseInt(configValues[4][0]) || 8;
  var waGroupId    = configValues[6][0].toString().trim();
  var fontteToken  = configValues[7][0].toString().trim();
  var reminderMin  = parseInt(configValues[8][0]) || 30;

  // Check WhatsApp config
  if (!fontteToken || fontteToken === "") {
    Logger.log("Fonnte Token tidak diset, skip WhatsApp reminder.");
    return;
  }

  // Ambil Data Tim
  var sheetTim = ss.getSheetByName(SHEET_TIM);
  var lastRow  = sheetTim.getLastRow();

  if (lastRow < 2) {
    Logger.log("Data tim kosong.");
    return;
  }

  var timData  = sheetTim.getRange(2, 1, lastRow - 1, 2).getValues();
  var timAktif = [];
  for (var i = 0; i < timData.length; i++) {
    if (timData[i][0].toString().trim() !== "" && timData[i][1].toString().trim() !== "") {
      timAktif.push(timData[i]);
    }
  }

  if (timAktif.length === 0) {
    Logger.log("Tidak ada anggota aktif.");
    return;
  }

  // Prepare Update Harian sheet - clear and populate with team data
  var sheetUpdate = ss.getSheetByName(SHEET_UPDATE);
  if (!sheetUpdate) {
    Logger.log("Sheet 'Update Harian' tidak ditemukan.");
    return;
  }

  // Clear data (keep header)
  if (sheetUpdate.getLastRow() > 1) {
    sheetUpdate.getRange(2, 1, sheetUpdate.getLastRow() - 1, 9).clear();
  }

  // Populate with team members
  var updateData = [];
  for (var i = 0; i < timAktif.length; i++) {
    updateData.push([timAktif[i][0], timAktif[i][1], "", "", "", "", "", "", ""]);
  }
  sheetUpdate.getRange(2, 1, updateData.length, 9).setValues(updateData);

  // Format Tanggal
  var hariArr   = ["Minggu","Senin","Selasa","Rabu","Kamis","Jumat","Sabtu"];
  var hariNama  = hariArr[day];
  var tanggal   = Utilities.formatDate(today, "Asia/Jakarta", "dd/MM/yyyy");
  var waktu     = (jamNotif < 10 ? "0" : "") + jamNotif + ":00";

  // Build WhatsApp message
  var sheetUrl = ss.getUrl() + "#gid=" + sheetUpdate.getSheetId();

  var message = "*REMINDER: Standup Meeting in " + reminderMin + " minutes*\n\n";
  message += "*Tim:* " + timName + "\n";
  message += "*Meeting:* " + hariNama + ", " + tanggal + " at " + waktu + " WIB\n";
  message += "*Link:* " + meetLink + "\n\n";
  message += "━━━━━━━━━━━━━━━\n\n";
  message += "*ACTION REQUIRED: Please update your status NOW*\n\n";
  message += "Update 'Update Harian' sheet dengan 6 items:\n\n";
  message += "1. Selesai Kemarin (What was completed)\n";
  message += "2. Akan Dikerjakan (What will be done today)\n";
  message += "3. Bug Report (Bug IDs or summary)\n";
  message += "4. Status Bug (Prod/Stg/Dev)\n";
  message += "5. Test Exec Status (% complete, blocked)\n";
  message += "6. Update Jira (Ticket IDs updated)\n\n";
  message += "Link Sheet: " + sheetUrl + "\n\n";
  message += "_Ini akan mempersingkat waktu meeting. Terima kasih!_";

  // Send WhatsApp
  if (waGroupId && waGroupId.includes("@g.us")) {
    // Send to group
    sendWhatsApp_(fontteToken, waGroupId, message);
    Logger.log("WhatsApp reminder sent to group: " + waGroupId);
  } else {
    // Send individual messages to each team member
    // Extract phone numbers from email (or use separate phone column if available)
    // For now, send to all via personal message (need phone numbers)
    Logger.log("WhatsApp Group ID tidak diset. Untuk individual messages, perlu tambah kolom nomor HP di sheet Tim.");
    Logger.log("Message yang akan dikirim:\n" + message);
  }
}


// ============================================================
// MAIN FUNCTION - ROLL ASSIGNMENT (Enhanced with Update Harian data)
// ============================================================
function rollAssignment() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var today = new Date();
  var day   = today.getDay();

  // Hanya Senin(1), Rabu(3), Jumat(5)
  if (day !== 1 && day !== 3 && day !== 5) {
    Logger.log("Bukan hari standup (" + day + "), skip.");
    return;
  }

  // Ambil Config
  var sheetConfig  = ss.getSheetByName(SHEET_CONFIG);
  var configValues = sheetConfig.getRange("B1:B9").getValues();
  var gchatWebhook = configValues[0][0].toString();
  var emailCC      = configValues[1][0].toString();
  var meetLink     = configValues[2][0].toString();
  var timName      = configValues[3][0].toString();
  var folderId     = configValues[5][0].toString();

  // Ambil Data Tim
  var sheetTim = ss.getSheetByName(SHEET_TIM);
  var lastRow  = sheetTim.getLastRow();

  if (lastRow < 2) {
    Logger.log("Data tim kosong.");
    return;
  }

  var timData  = sheetTim.getRange(2, 1, lastRow - 1, 2).getValues();
  var timAktif = [];
  for (var i = 0; i < timData.length; i++) {
    if (timData[i][0].toString().trim() !== "") {
      timAktif.push(timData[i]);
    }
  }

  if (timAktif.length === 0) {
    Logger.log("Tidak ada anggota aktif.");
    return;
  }

  // Get yesterday's PIC to exclude from today's pool
  var sheetJadwal = ss.getSheetByName(SHEET_JADWAL);
  var lastJadwalRow = sheetJadwal.getLastRow();
  var yesterdayPIC = null;

  if (lastJadwalRow >= 2) {
    var lastPIC = sheetJadwal.getRange(lastJadwalRow, 3).getValue().toString().trim();
    if (lastPIC) {
      yesterdayPIC = lastPIC;
      Logger.log("Yesterday's PIC: " + yesterdayPIC);
    }
  }

  // Filter tim aktif - exclude yesterday's PIC
  var availablePool = [];
  for (var i = 0; i < timAktif.length; i++) {
    var nama = timAktif[i][0].toString().trim();
    if (nama !== yesterdayPIC) {
      availablePool.push(timAktif[i]);
    }
  }

  // If all team members have been PIC recently, reset pool
  if (availablePool.length === 0) {
    Logger.log("All team members already PIC recently - resetting pool");
    availablePool = timAktif;
  }

  // Random Pick from available pool
  var randomIdx = Math.floor(Math.random() * availablePool.length);
  var picData   = availablePool[randomIdx];
  var namaPIC   = picData[0].toString().trim();
  var emailPIC  = picData[1].toString().trim();

  Logger.log("Selected PIC: " + namaPIC + " (from pool of " + availablePool.length + " candidates)");

  // Format Tanggal
  var hariArr   = ["Minggu","Senin","Selasa","Rabu","Kamis","Jumat","Sabtu"];
  var hariNama  = hariArr[day];
  var tanggal   = Utilities.formatDate(today, "Asia/Jakarta", "dd/MM/yyyy");
  var sesiLabel = hariNama + ", " + tanggal;

  // Read Update Harian data
  var updateData = readUpdateHarian_();

  // Get ALL team members for MOM doc (not just who filled Update Harian)
  var allTeamMembers = [];
  for (var i = 0; i < timAktif.length; i++) {
    allTeamMembers.push({
      nama: timAktif[i][0].toString().trim(),
      email: timAktif[i][1].toString().trim()
    });
  }

  // Generate MOM Doc with all team members + update data
  var docUrl = createMOMDoc(namaPIC, sesiLabel, tanggal, hariNama, meetLink, timName, folderId, allTeamMembers, updateData);

  if (!docUrl) {
    Logger.log("PERINGATAN: MOM Doc gagal dibuat. Notifikasi tetap dikirim tanpa link doc.");
  }

  // Simpan ke History (termasuk link doc)
  sheetJadwal.appendRow([tanggal, hariNama, namaPIC, emailPIC, docUrl || "Gagal dibuat", new Date()]);

  // Kirim Notifikasi
  if (gchatWebhook && gchatWebhook !== "PASTE_WEBHOOK_URL_DISINI") {
    var pesanChat = buildChatMessage(namaPIC, sesiLabel, meetLink, timName, docUrl);
    sendGoogleChat(gchatWebhook, pesanChat);
  } else {
    Logger.log("Google Chat webhook belum diset, skip.");
  }

  if (emailPIC) {
    sendEmail(emailPIC, emailCC, namaPIC, sesiLabel, meetLink, timName, docUrl);
  }

  Logger.log("Rolling selesai! PIC: " + namaPIC + " | Doc: " + (docUrl || "gagal"));
}


// ============================================================
// READ UPDATE HARIAN DATA
// ============================================================
function readUpdateHarian_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetUpdate = ss.getSheetByName(SHEET_UPDATE);

  if (!sheetUpdate) {
    Logger.log("Sheet Update Harian tidak ditemukan.");
    return [];
  }

  var lastRow = sheetUpdate.getLastRow();
  if (lastRow < 2) {
    Logger.log("Sheet Update Harian kosong.");
    return [];
  }

  // Read all data (skip header row 1)
  var data = sheetUpdate.getRange(2, 1, lastRow - 1, 9).getValues();
  var result = [];

  for (var i = 0; i < data.length; i++) {
    var nama = data[i][0].toString().trim();
    if (nama === "") continue;

    result.push({
      nama:          nama,
      email:         data[i][1].toString().trim(),
      selesai:       data[i][2].toString().trim(),
      akanDikerjakan: data[i][3].toString().trim(),
      bugReport:     data[i][4].toString().trim(),
      statusBug:     data[i][5].toString().trim(),
      testExec:      data[i][6].toString().trim(),
      updateJira:    data[i][7].toString().trim()
    });
  }

  return result;
}


// ============================================================
// BUILD PESAN - Google Chat (with mention)
// ============================================================
function buildChatMessage(namaPIC, sesiLabel, meetLink, timName, docUrl) {
  var lines = [
    "ROLLING MOM STANDUP - " + timName,
    "",
    "Sesi    : " + sesiLabel,
    "PIC MOM : <users/all> " + namaPIC,  // Mention in Google Chat
    "Meet    : " + meetLink
  ];

  if (docUrl) {
    lines.push("Template MOM : " + docUrl);
  } else {
    lines.push("Template MOM : (gagal dibuat, cek Apps Script log)");
  }

  lines.push("");
  lines.push("Tugas PIC MOM:");
  lines.push("- Buka link template MOM di atas");
  lines.push("- Catat update, blocker & action item selama standup");
  lines.push("- Setelah selesai, share link MOM ke tim");
  lines.push("");
  lines.push("Pesan ini dikirim otomatis oleh Standup Roller");

  return lines.join("\n");
}


// ============================================================
// KIRIM - Google Chat
// ============================================================
function sendGoogleChat(webhookUrl, pesan) {
  try {
    var options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ text: pesan })
    };
    UrlFetchApp.fetch(webhookUrl, options);
    Logger.log("Google Chat terkirim");
  } catch (e) {
    Logger.log("Gagal kirim Google Chat: " + e.message);
  }
}


// ============================================================
// KIRIM - WhatsApp (Fonnte API)
// ============================================================
function sendWhatsApp_(fontteToken, target, message) {
  try {
    var url = 'https://api.fonnte.com/send';

    var payload = {
      target: target,
      message: message
    };

    var options = {
      method: 'post',
      headers: {
        'Authorization': fontteToken
      },
      payload: payload,
      muteHttpExceptions: true
    };

    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();

    Logger.log('WhatsApp Response Code: ' + responseCode);
    Logger.log('WhatsApp Response: ' + responseText);

    if (responseCode === 200) {
      var data = JSON.parse(responseText);
      if (data.status) {
        Logger.log('WhatsApp sent successfully to: ' + target);
        return true;
      } else {
        Logger.log('WhatsApp send failed: ' + data.reason);
        return false;
      }
    } else {
      Logger.log('WhatsApp API error: ' + responseCode + ' - ' + responseText);
      return false;
    }
  } catch (e) {
    Logger.log('WhatsApp send exception: ' + e.message);
    return false;
  }
}


// ============================================================
// KIRIM - Email
// ============================================================
function sendEmail(emailTo, emailCC, namaPIC, sesiLabel, meetLink, timName, docUrl) {
  try {
    var subject = "[MOM Standup] PIC: " + namaPIC + " - " + sesiLabel;

    var docSection = "";
    if (docUrl) {
      docSection =
        "<tr style='background:#e8f5e9;'>" +
          "<td style='padding:10px;font-weight:bold;'>Template MOM</td>" +
          "<td style='padding:10px;'>" +
            "<a href='" + docUrl + "' style='color:#1a73e8;font-weight:bold;font-size:14px;'>" +
              "Klik di sini untuk buka Template MOM" +
            "</a>" +
          "</td>" +
        "</tr>";
    }

    var htmlBody =
      "<div style='font-family:Arial,sans-serif;max-width:640px;border:1px solid #e0e0e0;border-radius:8px;overflow:hidden;'>" +
        "<div style='background:#1a73e8;padding:20px;text-align:center;'>" +
          "<h2 style='color:white;margin:0;'>Rolling MOM Standup</h2>" +
          "<p style='color:#bbdefb;margin:5px 0 0;'>" + timName + "</p>" +
        "</div>" +
        "<table style='width:100%;border-collapse:collapse;'>" +
          "<tr style='background:#f8f9fa;'>" +
            "<td style='padding:10px;font-weight:bold;width:35%;'>Sesi</td>" +
            "<td style='padding:10px;'>" + sesiLabel + "</td>" +
          "</tr>" +
          "<tr>" +
            "<td style='padding:10px;font-weight:bold;'>PIC MOM</td>" +
            "<td style='padding:10px;font-size:20px;color:#1a73e8;font-weight:bold;'>" + namaPIC + "</td>" +
          "</tr>" +
          "<tr style='background:#f8f9fa;'>" +
            "<td style='padding:10px;font-weight:bold;'>Google Meet</td>" +
            "<td style='padding:10px;'>" +
              "<a href='" + meetLink + "' style='color:#1a73e8;'>" + meetLink + "</a>" +
            "</td>" +
          "</tr>" +
          docSection +
        "</table>" +
        "<div style='padding:16px 20px;background:#fff8e1;border-top:1px solid #ffe082;'>" +
          "<h3 style='color:#f57f17;margin:0 0 10px;'>Tugas Kamu sebagai PIC MOM</h3>" +
          "<ol style='margin:0;padding-left:20px;line-height:1.9;'>" +
            "<li>Buka link <b>Template MOM</b> di atas sebelum standup dimulai</li>" +
            "<li>Catat update harian, blocker dan action item setiap anggota</li>" +
            "<li>Lengkapi bagian <i>Issues &amp; Blocker</i> dan <i>Action Items</i></li>" +
            "<li>Setelah standup selesai, <b>share link MOM</b> ke grup tim</li>" +
          "</ol>" +
        "</div>" +
        "<div style='background:#f1f3f4;padding:12px;text-align:center;'>" +
          "<p style='margin:0;color:#9e9e9e;font-size:11px;'>" +
            "Pesan ini dikirim otomatis oleh Standup Roller | " + sesiLabel +
          "</p>" +
        "</div>" +
      "</div>";

    var mailOptions = { htmlBody: htmlBody };

    if (emailCC && emailCC.trim() !== "") {
      var ccParts    = emailCC.split(",");
      var ccFiltered = [];
      for (var i = 0; i < ccParts.length; i++) {
        var t = ccParts[i].trim();
        if (t) { ccFiltered.push(t); }
      }
      if (ccFiltered.length > 0) {
        mailOptions.cc = ccFiltered.join(",");
      }
    }

    GmailApp.sendEmail(emailTo, subject, "", mailOptions);
    Logger.log("Email terkirim ke " + emailTo);
  } catch (e) {
    Logger.log("Gagal kirim email: " + e.message);
  }
}


// ============================================================
// CREATE MOM GOOGLE DOC (Simple Narrative Format)
// ============================================================
function createMOMDoc(namaPIC, sesiLabel, tanggal, hariNama, meetLink, timName, folderId, allTeamMembers, updateData) {
  try {
    // Buat doc di root dulu (paling aman)
    var docName = "MOM Standup - " + hariNama + " " + tanggal + " - PIC " + namaPIC;
    var doc     = DocumentApp.create(docName);
    var docId   = doc.getId();
    var body    = doc.getBody();

    Logger.log("Doc dibuat dengan ID: " + docId);

    body.setMarginTop(50).setMarginBottom(50).setMarginLeft(60).setMarginRight(60);

    // JUDUL
    var judul = body.appendParagraph("MINUTES OF MEETING (MOM)");
    judul.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    judul.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    judul.editAsText().setForegroundColor("#1a73e8").setBold(true);

    var subJudul = body.appendParagraph("Bi-Daily Standup - " + timName);
    subJudul.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    subJudul.editAsText().setForegroundColor("#5f6368").setItalic(true).setFontSize(11);

    body.appendHorizontalRule();
    body.appendParagraph("");

    // INFO MEETING
    var infoTable = body.appendTable([
      ["Tanggal / Sesi",       sesiLabel],
      ["Tim",                  timName],
      ["PIC MOM",              namaPIC],
      ["Google Meet",          meetLink],
      ["Waktu Mulai",          ""],
      ["Waktu Selesai",        ""],
      ["Peserta Hadir",        ""],
      ["Peserta Tidak Hadir",  ""]
    ]);

    infoTable.setBorderColor("#e0e0e0");
    for (var r = 0; r < infoTable.getNumRows(); r++) {
      infoTable.getCell(r, 0).setBackgroundColor("#f8f9fa")
               .editAsText().setBold(true).setFontSize(10);
      infoTable.getCell(r, 1).editAsText().setFontSize(10);
    }
    // Highlight nama PIC
    infoTable.getCell(2, 1).setBackgroundColor("#e8f0fe")
             .editAsText().setForegroundColor("#1a73e8").setBold(true);

    body.appendParagraph("");

    // SECTION 1: Update Harian (Simple Narrative Format)
    var s1 = body.appendParagraph("1. UPDATE HARIAN PER ANGGOTA");
    s1.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    s1.editAsText().setForegroundColor("#1e8e3e").setBold(true);

    var reminderText = body.appendParagraph(
      "Isi untuk setiap anggota (3 essentials + reminder):\n" +
      "• Kemarin: Apa yang sudah selesai\n" +
      "• Hari Ini: Apa yang akan dikerjakan\n" +
      "• Notes: Bug report, Status bug (Prod/Stg/Dev), Test exec status, Update Jira (jika ada)"
    );
    reminderText.editAsText().setItalic(true).setForegroundColor("#9e9e9e").setFontSize(9);

    body.appendParagraph("");

    // Create update data map for quick lookup
    var updateMap = {};
    if (updateData && updateData.length > 0) {
      for (var u = 0; u < updateData.length; u++) {
        updateMap[updateData[u].nama] = updateData[u];
      }
    }

    // Iterate through ALL team members (from Tab Tim)
    for (var i = 0; i < allTeamMembers.length; i++) {
      var member = allTeamMembers[i];
      var nama = member.nama;
      var prefilledData = updateMap[nama];  // Check if they filled Update Harian

      // Member name (bold)
      var namaPara = body.appendParagraph(nama);
      namaPara.editAsText().setBold(true).setFontSize(11).setForegroundColor("#1a73e8");

      // Kemarin (pre-filled if available)
      var kemarinText = "Kemarin: ";
      if (prefilledData && prefilledData.selesai) {
        kemarinText += prefilledData.selesai;
      }
      body.appendParagraph(kemarinText).editAsText().setFontSize(10);

      // Hari Ini (pre-filled if available)
      var hariIniText = "Hari Ini: ";
      if (prefilledData && prefilledData.akanDikerjakan) {
        hariIniText += prefilledData.akanDikerjakan;
      }
      body.appendParagraph(hariIniText).editAsText().setFontSize(10);

      // Notes (combine all optional fields, or leave blank for PIC to fill)
      var notesText = "Notes: ";
      var notesParts = [];

      if (prefilledData) {
        if (prefilledData.bugReport) notesParts.push("Bug: " + prefilledData.bugReport);
        if (prefilledData.statusBug) notesParts.push("Status: " + prefilledData.statusBug);
        if (prefilledData.testExec) notesParts.push("Test: " + prefilledData.testExec);
        if (prefilledData.updateJira) notesParts.push("Jira: " + prefilledData.updateJira);
      }

      if (notesParts.length > 0) {
        notesText += notesParts.join(" | ");
      }

      body.appendParagraph(notesText).editAsText().setFontSize(10).setForegroundColor("#5f6368");
      body.appendParagraph("");  // Spacing between members
    }

    body.appendParagraph("");

    // SECTION 2: Issues & Blocker
    var s2 = body.appendParagraph("2. ISSUES & BLOCKER");
    s2.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    s2.editAsText().setForegroundColor("#d93025").setBold(true);

    var blockerRows = [
      ["No", "Issue / Blocker", "Dilaporkan Oleh", "Solusi / Action Plan", "PIC", "Target Selesai"],
      ["1", "", "", "", "", ""],
      ["2", "", "", "", "", ""],
      ["3", "", "", "", "", ""],
      ["4", "", "", "", "", ""],
      ["5", "", "", "", "", ""]
    ];
    var blockerTable = body.appendTable(blockerRows);
    styleTableHeader(blockerTable, "#ea4335");
    blockerTable.setBorderColor("#e0e0e0");

    body.appendParagraph("");

    // SECTION 3: Action Items
    var s3 = body.appendParagraph("3. ACTION ITEMS");
    s3.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    s3.editAsText().setForegroundColor("#e37400").setBold(true);

    var actionRows = [
      ["No", "Action Item", "PIC", "Deadline", "Status"],
      ["1", "", "", "", "Open"],
      ["2", "", "", "", "Open"],
      ["3", "", "", "", "Open"],
      ["4", "", "", "", "Open"],
      ["5", "", "", "", "Open"]
    ];
    var actionTable = body.appendTable(actionRows);
    styleTableHeader(actionTable, "#f9ab00");
    actionTable.setBorderColor("#e0e0e0");

    body.appendParagraph("");

    // SECTION 4: Catatan Tambahan
    var s4 = body.appendParagraph("4. CATATAN TAMBAHAN");
    s4.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    s4.editAsText().setForegroundColor("#7627bb").setBold(true);

    body.appendParagraph("(Isi jika ada pengumuman, info sprint, atau hal lain yang perlu dicatat)")
        .editAsText().setItalic(true).setForegroundColor("#9e9e9e").setFontSize(9);

    body.appendParagraph(" ");
    body.appendParagraph(" ");

    // SECTION 5: Dibuat Oleh
    var s5 = body.appendParagraph("5. DIBUAT OLEH");
    s5.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    s5.editAsText().setForegroundColor("#5f6368").setBold(true);

    var signTable = body.appendTable([
      ["Nama PIC MOM",   namaPIC],
      ["Tanggal Dibuat", tanggal],
      ["Tanda Tangan",   " "]
    ]);
    signTable.setBorderColor("#e0e0e0");
    for (var rs = 0; rs < signTable.getNumRows(); rs++) {
      signTable.getCell(rs, 0).setBackgroundColor("#f8f9fa")
               .editAsText().setBold(true).setFontSize(10);
    }

    body.appendParagraph("");
    body.appendHorizontalRule();

    var footerP = body.appendParagraph(
      "Dokumen dibuat otomatis oleh Standup Roller v3 | " + sesiLabel + " | PIC: " + namaPIC
    );
    footerP.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    footerP.editAsText().setForegroundColor("#bdbdbd").setItalic(true).setFontSize(8);

    // Simpan doc sebelum pindah folder
    doc.saveAndClose();

    // Pindah ke folder tujuan (jika Folder ID valid)
    var docFile = DriveApp.getFileById(docId);
    var docUrl  = "https://docs.google.com/document/d/" + docId + "/edit";

    if (folderId && folderId !== "PASTE_FOLDER_ID_DISINI" && folderId.trim() !== "") {
      try {
        var targetFolder = DriveApp.getFolderById(folderId.trim());
        targetFolder.addFile(docFile);
        // Hapus dari root agar tidak duplikat
        var rootFolder = DriveApp.getRootFolder();
        if (rootFolder.getFilesByName(docName).hasNext()) {
          rootFolder.removeFile(docFile);
        }
        Logger.log("Doc berhasil dipindah ke folder: " + folderId);
      } catch (folderErr) {
        Logger.log("Gagal pindah ke folder (doc tetap di root): " + folderErr.message);
      }
    } else {
      Logger.log("Folder ID tidak diset, doc disimpan di root Drive.");
    }

    Logger.log("MOM Doc URL: " + docUrl);
    return docUrl;

  } catch (e) {
    Logger.log("FATAL - Gagal buat MOM Doc: " + e.message);
    return null;
  }
}


// ============================================================
// HELPER - Style header row tabel
// ============================================================
function styleTableHeader(table, bgColor) {
  var headerRow = table.getRow(0);
  for (var c = 0; c < headerRow.getNumCells(); c++) {
    headerRow.getCell(c)
             .setBackgroundColor(bgColor)
             .editAsText()
             .setForegroundColor("#ffffff")
             .setBold(true)
             .setFontSize(10);
  }
}


// ============================================================
// TEST - WhatsApp Reminder
// ============================================================
function testWhatsAppReminder() {
  sendPreMeetingReminder();
  SpreadsheetApp.getUi().alert(
    "Test WhatsApp Reminder",
    "WhatsApp reminder test selesai.\n\nCek execution log untuk detail.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


// ============================================================
// TEST - Buat MOM Doc saja (tanpa tunggu hari standup)
// ============================================================
function testCreateDoc() {
  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var sheetConfig = ss.getSheetByName(SHEET_CONFIG);
  var configVals  = sheetConfig.getRange("B1:B6").getValues();
  var meetLink    = configVals[2][0].toString();
  var timName     = configVals[3][0].toString();
  var folderId    = configVals[5][0].toString();

  var today     = new Date();
  var tanggal   = Utilities.formatDate(today, "Asia/Jakarta", "dd/MM/yyyy");
  var hariArr   = ["Minggu","Senin","Selasa","Rabu","Kamis","Jumat","Sabtu"];
  var hariNama  = hariArr[today.getDay()];
  var sesiLabel = hariNama + ", " + tanggal + " (TEST)";

  // Read test data from Update Harian
  var updateData = readUpdateHarian_();

  var url = createMOMDoc("Test PIC", sesiLabel, tanggal, hariNama, meetLink, timName, folderId, updateData);

  SpreadsheetApp.getUi().alert(
    url ? "MOM Doc Berhasil Dibuat!" : "Gagal Membuat MOM Doc",
    url
      ? "Doc berhasil dibuat dengan data dari 'Update Harian' sheet!\n\nLink: " + url + "\n\nCek juga folder Google Drive kamu."
      : "Gagal membuat doc.\n\nCara debug:\n1. Buka Apps Script\n2. Klik 'Executions' di menu kiri\n3. Lihat error message di log",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


// ============================================================
// CEK FOLDER ID - Validasi apakah Folder ID benar
// ============================================================
function cekFolderId() {
  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var sheetConfig = ss.getSheetByName(SHEET_CONFIG);
  var folderId    = sheetConfig.getRange("B6").getValue().toString().trim();
  var ui          = SpreadsheetApp.getUi();

  if (!folderId || folderId === "PASTE_FOLDER_ID_DISINI" || folderId === "") {
    ui.alert(
      "Folder ID Belum Diisi",
      "Isi Folder ID di sheet Config baris 6.\n\n" +
      "Cara ambil Folder ID:\n" +
      "1. Buka Google Drive\n" +
      "2. Masuk ke folder tujuan MOM\n" +
      "3. Lihat URL: drive.google.com/drive/folders/[INI_FOLDER_ID]\n" +
      "4. Copy dan paste ID tersebut ke Config B6",
      ui.ButtonSet.OK
    );
    return;
  }

  try {
    var folder = DriveApp.getFolderById(folderId);
    ui.alert(
      "Folder ID Valid!",
      "Folder ditemukan: " + folder.getName() + "\n\n" +
      "MOM Doc akan disimpan ke folder ini.",
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert(
      "Folder ID Tidak Valid!",
      "Error: " + e.message + "\n\n" +
      "Kemungkinan penyebab:\n" +
      "1. Folder ID salah (cek lagi URL folder)\n" +
      "2. Folder tidak di-share ke akun Google kamu\n" +
      "3. Folder dimiliki akun Google lain\n\n" +
      "Solusi: Pastikan folder di-share dengan akun yang menjalankan script ini.",
      ui.ButtonSet.OK
    );
  }
}


// ============================================================
// SHORTCUT - Buka sheet Jadwal
// ============================================================
function openJadwal() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(ss.getSheetByName(SHEET_JADWAL));
}
