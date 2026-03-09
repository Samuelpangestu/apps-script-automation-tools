/**
 * forceAuth.js - Force trigger authorization popup
 *
 * Function ini SENGAJA menggunakan UrlFetchApp untuk trigger authorization.
 * Jalankan ini SETELAH clear all authorizations.
 */

function forceAuthorizationPopup() {
  try {
    // Simple fetch to trigger auth popup
    const response = UrlFetchApp.fetch('https://www.google.com', {
      muteHttpExceptions: true
    });

    Logger.log('✅ Authorization granted! Response code: ' + response.getResponseCode());

    SpreadsheetApp.getUi().alert(
      '✅ SUCCESS!\n\n' +
      'UrlFetchApp.fetch() berhasil dipanggil.\n' +
      'Authorization sudah di-grant.\n\n' +
      'Sekarang jalankan jiraActivate() untuk test koneksi Jira.'
    );

  } catch (e) {
    Logger.log('❌ Error: ' + e.message);

    if (e.message.includes('permission')) {
      SpreadsheetApp.getUi().alert(
        '❌ Authorization GAGAL\n\n' +
        'Error: ' + e.message + '\n\n' +
        '── TROUBLESHOOTING ──\n\n' +
        '1. Pastikan appsscript.json berisi:\n' +
        '   "https://www.googleapis.com/auth/script.external_request"\n\n' +
        '2. Clear authorization:\n' +
        '   Run → Clear all authorizations\n\n' +
        '3. Jalankan forceAuthorizationPopup() lagi\n\n' +
        '4. Klik "Review Permissions" dan Allow'
      );
    } else {
      SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
    }
  }
}
