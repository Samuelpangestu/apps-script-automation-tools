/**
 * setupGoogleChat.js - Simplified Google Chat setup
 *
 * Alternative untuk jiraSetup() yang hanya fokus menambahkan
 * section Google Chat Notification
 */

function setupGoogleChatOnly() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cfg = ss.getSheetByName('Config');

    if (!cfg) {
      SpreadsheetApp.getUi().alert('Config tab tidak ditemukan!');
      return;
    }

    Logger.log('Spreadsheet ID: ' + ss.getId());
    Logger.log('Config sheet found: ' + cfg.getName());

    // Check if Google Chat section already exists
    const data = cfg.getDataRange().getValues();
    const existing = data.some(row => String(row[0]).toUpperCase().includes('GOOGLE CHAT'));

    if (existing) {
      SpreadsheetApp.getUi().alert(
        'ℹ️  Section GOOGLE CHAT NOTIFICATION sudah ada!\n\n' +
        'Silakan scroll ke bawah di Config tab untuk melihatnya.'
      );
      Logger.log('Google Chat section already exists');
      return;
    }

    // Find last row with data
    const lastRow = cfg.getLastRow();
    Logger.log('Last row: ' + lastRow);

    // Add Google Chat section 3 rows below last data
    const chatSec = lastRow + 3;
    Logger.log('Adding Google Chat section at row: ' + chatSec);

    // Row 1: Header
    cfg.getRange(chatSec, 1, 1, 12).merge()
      .setValue('GOOGLE CHAT NOTIFICATION  —  Notifikasi Blocker Harian ke Google Chat')
      .setBackground('#1565C0')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setFontSize(10)
      .setFontFamily('Arial')
      .setHorizontalAlignment('left');
    cfg.setRowHeight(chatSec, 26);
    Logger.log('Header row created');

    // Row 2: Info
    cfg.getRange(chatSec + 1, 1, 1, 12).merge()
      .setValue('💬  Notifikasi blocker otomatis dikirim ke Google Chat Space setiap hari. Atur webhook URL dan waktu notifikasi di bawah.')
      .setBackground('#E3F2FD')
      .setFontColor('#1565C0')
      .setFontStyle('italic')
      .setFontSize(8);
    cfg.setRowHeight(chatSec + 1, 16);
    Logger.log('Info row created');

    // Row 3: Column headers
    const headers = [
      'Google Chat Webhook URL',
      'Notif Time (Hour)',
      'Enable Notifikasi'
    ];

    headers.forEach((h, i) => {
      cfg.getRange(chatSec + 2, i + 1)
        .setValue(h)
        .setBackground('#1976D2')
        .setFontColor('#FFFFFF')
        .setFontWeight('bold')
        .setFontSize(9)
        .setFontFamily('Arial')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle')
        .setBorder(true, true, true, true, false, false, '#90CAF9', SpreadsheetApp.BorderStyle.SOLID);
    });

    // Set column widths
    cfg.setColumnWidth(1, 450);  // Webhook URL
    cfg.setColumnWidth(2, 120);  // Hour
    cfg.setColumnWidth(3, 140);  // Enable
    cfg.setRowHeight(chatSec + 2, 22);
    Logger.log('Column headers created');

    // Row 4: Data row with defaults
    cfg.getRange(chatSec + 3, 1, 1, 3)
      .setValues([['https://chat.googleapis.com/v1/spaces/...', 15, 'No']])
      .setBackground('#E3F2FD')
      .setFontFamily('Arial')
      .setFontSize(9)
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, false, false, '#90CAF9', SpreadsheetApp.BorderStyle.SOLID);

    cfg.getRange(chatSec + 3, 1).setFontFamily('Courier New').setFontSize(8);

    // Data validation for Enable (Yes/No)
    const dvEnable = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Yes', 'No'], true)
      .build();
    cfg.getRange(chatSec + 3, 3).setDataValidation(dvEnable);

    // Data validation for Hour (0-23)
    const dvHour = SpreadsheetApp.newDataValidation()
      .requireNumberBetween(0, 23)
      .build();
    cfg.getRange(chatSec + 3, 2).setDataValidation(dvHour);

    cfg.setRowHeight(chatSec + 3, 22);
    Logger.log('Data row created with validation');

    // Add notes
    cfg.getRange(chatSec + 2, 1).setNote(
      'Buat webhook di Google Chat Space:\n' +
      'Space Settings > Apps & integrations > Webhooks\n\n' +
      'Format: https://chat.googleapis.com/v1/spaces/.../messages?key=...'
    );

    cfg.getRange(chatSec + 2, 2).setNote(
      'Jam berapa notifikasi dikirim (0-23)\n' +
      'Contoh: 15 = jam 3 sore'
    );

    cfg.getRange(chatSec + 2, 3).setNote(
      'Yes = aktif notifikasi harian\n' +
      'No = nonaktifkan'
    );

    Logger.log('✅ Google Chat section created successfully');

    // Activate Config sheet
    ss.setActiveSheet(cfg);

    SpreadsheetApp.getUi().alert(
      '✅ Section GOOGLE CHAT NOTIFICATION berhasil ditambahkan!\n\n' +
      'Langkah selanjutnya:\n\n' +
      '1. Buat webhook di Google Chat Space:\n' +
      '   Space Settings > Apps & integrations > Webhooks\n\n' +
      '2. Di Config tab, scroll ke bawah\n\n' +
      '3. Isi 3 kolom:\n' +
      '   • Webhook URL (paste dari Google Chat)\n' +
      '   • Notif Time (0-23, default: 15)\n' +
      '   • Enable = No (untuk test dulu)\n\n' +
      '4. Test dengan jalankan: sendBlockerNotification()\n\n' +
      '5. Jika berhasil, set Enable = Yes dan jalankan:\n' +
      '   setupDailyBlockerNotification()'
    );

  } catch (e) {
    Logger.log('❌ Error: ' + e.message);
    Logger.log('Stack: ' + e.stack);

    SpreadsheetApp.getUi().alert(
      '❌ Error saat setup Google Chat section:\n\n' +
      e.message +
      '\n\nDetail error ada di Execution log.\n' +
      'Atau setup manual mengikuti instruksi dari assistant.'
    );
  }
}

/**
 * Manual cleanup - hapus section Google Chat jika perlu redo
 */
function removeGoogleChatSection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName('Config');

  if (!cfg) {
    SpreadsheetApp.getUi().alert('Config tab tidak ditemukan!');
    return;
  }

  const data = cfg.getDataRange().getValues();
  let chatRow = -1;

  // Find Google Chat section header
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).toUpperCase().includes('GOOGLE CHAT NOTIFICATION')) {
      chatRow = i + 1; // Convert to 1-based row number
      break;
    }
  }

  if (chatRow === -1) {
    SpreadsheetApp.getUi().alert('Section GOOGLE CHAT NOTIFICATION tidak ditemukan.');
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Konfirmasi Hapus',
    'Hapus section GOOGLE CHAT NOTIFICATION di row ' + chatRow + '-' + (chatRow + 3) + '?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  // Delete 4 rows (header + info + column headers + data)
  cfg.deleteRows(chatRow, 4);

  SpreadsheetApp.getUi().alert('✅ Section GOOGLE CHAT NOTIFICATION berhasil dihapus.');
  Logger.log('Google Chat section removed from row ' + chatRow);
}
