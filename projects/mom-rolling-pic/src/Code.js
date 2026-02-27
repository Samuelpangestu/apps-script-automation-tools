// ============================================================
// STANDUP MOM ROLLER v2
// Fitur: Random Rolling, Google Chat, Email, MOM Docs
// Jadwal: Senin, Rabu, Jumat
// Update: Hapus WhatsApp, fix MOM Doc creation
// ============================================================

var SHEET_TIM    = "Tim";
var SHEET_JADWAL = "Jadwal";
var SHEET_CONFIG = "Config";


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

  // Sheet: Config
  var sheetConfig = ss.getSheetByName(SHEET_CONFIG);
  if (!sheetConfig) { sheetConfig = ss.insertSheet(SHEET_CONFIG); }
  sheetConfig.clearContents();
  sheetConfig.clearFormats();

  var configData = [
    ["Google Chat Webhook URL",          "PASTE_WEBHOOK_URL_DISINI"],
    ["Email CC (pisahkan dengan koma)",  "manager@email.com"],
    ["Google Meet Link Recurring",       "https://meet.google.com/xxx-xxxx-xxx"],
    ["Nama Tim",                         "QE INADigital"],
    ["Jam Notifikasi (angka saja)",      "8"],
    ["Google Drive Folder ID untuk MOM", "PASTE_FOLDER_ID_DISINI"]
  ];

  sheetConfig.getRange("A1:B6").setValues(configData);
  sheetConfig.getRange("A1:A6").setFontWeight("bold").setBackground("#f8f9fa");
  sheetConfig.getRange("A1:B6").setBorder(
    true, true, true, true, true, true,
    "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID
  );
  sheetConfig.autoResizeColumns(1, 2);

  // Hapus sheet WhatsApp Log jika ada
  var sheetWA = ss.getSheetByName("WhatsApp Log");
  if (sheetWA) { ss.deleteSheet(sheetWA); }

  ui.alert(
    "Setup Selesai!",
    "Langkah selanjutnya:\n\n" +
    "1. Isi data tim di sheet 'Tim' (Nama & Email)\n" +
    "2. Lengkapi sheet 'Config'\n" +
    "3. Klik menu > Aktifkan Trigger Otomatis\n\n" +
    "Tips ambil Folder ID Google Drive:\n" +
    "Buka folder di Drive > lihat URL > salin ID setelah /folders/",
    ui.ButtonSet.OK
  );
}


// ============================================================
// SETUP TRIGGER
// ============================================================
function setupTrigger() {
  var ui = SpreadsheetApp.getUi();

  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetConfig = ss.getSheetByName(SHEET_CONFIG);
  var jamNotif = parseInt(sheetConfig.getRange("B5").getValue()) || 8;

  ScriptApp.newTrigger("rollAssignment")
    .timeBased()
    .everyDays(1)
    .atHour(jamNotif)
    .create();

  ui.alert(
    "Trigger Aktif!",
    "Rolling otomatis berjalan setiap:\n" +
    "- Senin jam " + jamNotif + ":00\n" +
    "- Rabu jam " + jamNotif + ":00\n" +
    "- Jumat jam " + jamNotif + ":00\n\n" +
    "(Hari lain otomatis di-skip)",
    ui.ButtonSet.OK
  );
}


// ============================================================
// MAIN FUNCTION
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
  var configValues = sheetConfig.getRange("B1:B6").getValues();
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

  // Random Pick
  var randomIdx = Math.floor(Math.random() * timAktif.length);
  var picData   = timAktif[randomIdx];
  var namaPIC   = picData[0].toString().trim();
  var emailPIC  = picData[1].toString().trim();

  // Format Tanggal
  var hariArr   = ["Minggu","Senin","Selasa","Rabu","Kamis","Jumat","Sabtu"];
  var hariNama  = hariArr[day];
  var tanggal   = Utilities.formatDate(today, "Asia/Jakarta", "dd/MM/yyyy");
  var sesiLabel = hariNama + ", " + tanggal;

  // Generate MOM Doc DULU sebelum kirim notif
  var docUrl = createMOMDoc(namaPIC, sesiLabel, tanggal, hariNama, meetLink, timName, folderId);

  if (!docUrl) {
    Logger.log("PERINGATAN: MOM Doc gagal dibuat. Notifikasi tetap dikirim tanpa link doc.");
  }

  // Simpan ke History (termasuk link doc)
  var sheetJadwal = ss.getSheetByName(SHEET_JADWAL);
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
// BUILD PESAN - Google Chat
// ============================================================
function buildChatMessage(namaPIC, sesiLabel, meetLink, timName, docUrl) {
  var lines = [
    "ROLLING MOM STANDUP - " + timName,
    "",
    "Sesi    : " + sesiLabel,
    "PIC MOM : " + namaPIC,
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
// CREATE MOM GOOGLE DOC
// ============================================================
function createMOMDoc(namaPIC, sesiLabel, tanggal, hariNama, meetLink, timName, folderId) {
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

    // SECTION 1: Update Harian
    var s1 = body.appendParagraph("1. UPDATE HARIAN PER ANGGOTA");
    s1.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    s1.editAsText().setForegroundColor("#1e8e3e").setBold(true);

    body.appendParagraph("Isi update masing-masing anggota di bawah ini:")
        .editAsText().setItalic(true).setForegroundColor("#9e9e9e").setFontSize(9);

    var updateRows = [["No", "Nama Anggota", "Kemarin (Sudah Dikerjakan)", "Hari Ini (Akan Dikerjakan)", "Blocker / Kendala"]];
    for (var u = 1; u <= 15; u++) {
      updateRows.push([u.toString(), "", "", "", ""]);
    }

    var updateTable = body.appendTable(updateRows);
    styleTableHeader(updateTable, "#1a73e8");
    updateTable.setBorderColor("#e0e0e0");
    updateTable.setColumnWidth(0, 25);
    updateTable.setColumnWidth(1, 100);
    updateTable.setColumnWidth(2, 155);
    updateTable.setColumnWidth(3, 155);
    updateTable.setColumnWidth(4, 120);

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
      "Dokumen dibuat otomatis oleh Standup Roller | " + sesiLabel + " | PIC: " + namaPIC
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

  var url = createMOMDoc("Test PIC", sesiLabel, tanggal, hariNama, meetLink, timName, folderId);

  SpreadsheetApp.getUi().alert(
    url ? "MOM Doc Berhasil Dibuat!" : "Gagal Membuat MOM Doc",
    url
      ? "Doc berhasil dibuat!\n\nLink: " + url + "\n\nCek juga folder Google Drive kamu."
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