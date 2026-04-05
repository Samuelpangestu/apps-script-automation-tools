/**
 * jira_sync.js  —  Jira Cloud → QA Sheet BugReport Sync  +  Notifikasi
 * ═══════════════════════════════════════════════════════════════════════
 *
 * SETUP (2 langkah):
 *   1. jiraSetup()    → tambah kolom + section credentials di Config tab
 *   2. jiraActivate() → verifikasi koneksi + aktifkan semua trigger
 *
 * MANUAL:
 *   syncJiraToAllSheets()  → sync sekarang (tanpa update Status)
 *   syncJiraStatusAll()    → update Status dari Jira
 *   sendBugNotification()  → kirim notif email pending bugs
 *   showJiraJQL()          → lihat JQL & URL yang digunakan (untuk debugging)
 *
 * CUSTOMIZE JQL PER MODUL:
 *   - JQL otomatis filter per modul menggunakan custom field "modul[dropdown]"
 *   - Nilai modul diambil dari Config kolom "Modul" (col D)
 *   - Edit JIRA_JQL_ (line 32-36) untuk ubah field name atau query global
 *   - Contoh: ganti "modul[dropdown]" → "Modul" atau cf[10001]
 *   - Jalankan showJiraJQL() untuk lihat query aktif per modul
 *
 * DEBUGGING:
 *   - Semua JQL & URL otomatis di-log ke Execution log
 *   - View > Execution log di Apps Script editor
 *   - URL browser juga tersedia untuk test manual di Jira
 * ═══════════════════════════════════════════════════════════════════════
 */

const JIRA_INSTANCES_ = {
  'digitalperuri': 'https://digitalperuri.atlassian.net',
  'bgn-peruri':    'https://bgn-peruri.atlassian.net',
};

// Modul field berbeda per instance
const JIRA_MODUL_FIELD_ = {
  'digitalperuri': 'cf[10097]',  // customfield_10097
  'bgn-peruri':    'cf[10289]',  // customfield_10289
};

// JQL fetch SEMUA bug (termasuk Closed) untuk detect perubahan status
// Cleanup logic akan hapus bug yang Closed/Won't Fix dari sheet
const JIRA_JQL_ =
  'project = "{P}" AND issuetype = Bug ' +
  'AND "{F}" = "{M}" ' +
  'ORDER BY priority ASC, updated DESC';
// Custom fields for Feature/Environment/Steps/Expected/Actual/Submodul/Screenshot (per instance)
const JIRA_CUSTOM_FIELDS_ = {
  'digitalperuri': ',customfield_11090,customfield_10095,customfield_10560,customfield_10561,customfield_10562,customfield_11354,customfield_10179',  // Feature, Environment, Step To Reproduce, Expectation Result, Actual Result, Submodul, Screenshot/Video
  'bgn-peruri': ',customfield_10298,customfield_10291,customfield_10292,customfield_10293,customfield_10294,customfield_10300,customfield_10296'     // Feature, Environment, Step To Reproduce, Expectation Result, Actual Result, Submodul, Screenshot/Video
};

// Custom field mappings for accessing fields by instance
const JIRA_FIELD_MAP_ = {
  'digitalperuri': {
    feature: 'customfield_11090',
    environment: 'customfield_10095',
    steps: 'customfield_10560',
    expected: 'customfield_10561',
    actual: 'customfield_10562',
    submodul: 'customfield_11354',
    screenshot: 'customfield_10179'
  },
  'bgn-peruri': {
    feature: 'customfield_10298',
    environment: 'customfield_10291',
    steps: 'customfield_10292',
    expected: 'customfield_10293',
    actual: 'customfield_10294',
    submodul: 'customfield_10300',
    screenshot: 'customfield_10296'
  }
};

const JIRA_FIELDS_ = 'summary,description,priority,status,assignee,reporter,resolutiondate,key,created,updated,labels,components,environment,attachment';
const BUG_START_   = 5;
const BC_ = {
  BUG_ID:1,TYPE:2,PRIORITY:3,STATUS:4,FEATURE:5,SUBMODUL:6,
  TITLE:7,DESC:8,ENV:9,STEPS:10,EXP:11,ACT:12,TC:13,REPORTED_BY:14,
  ASSIGNED:15,DATE_FOUND:16,DATE_FIXED:17,SPRINT:18,LINK:19,NOTES:20,SCREENSHOT:21,JIRA_KEY:22,SYNCED:23
};
const BUG_COLS_ = 23;


// ═══════════════════════════════════════════════════════════════════════
// SETUP & ACTIVATE  ← 2 fungsi yang perlu dijalankan user
// ═══════════════════════════════════════════════════════════════════════

/**
 * LANGKAH 1: jiraSetup()
 * Jalankan SEKALI dari Apps Script editor.
 * Yang dilakukan:
 *   - Tambah kolom H (Jira Sync), I (Jira Instance) dan J (Jira Project Key) ke Config
 *   - Tambah section JIRA CREDENTIALS di bawah daftar modul
 *   - Tampilkan instruksi apa yang harus diisi
 */
function jiraSetup() {
  try {
    const ss  = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log('Spreadsheet ID: ' + ss.getId());
    Logger.log('Spreadsheet Name: ' + ss.getName());

    const cfg = ss.getSheetByName('Config');
    if (!cfg) {
      safeAlert_('Config tab tidak ditemukan. Jalankan createDashboard() dulu.');
      return;
    }

    const log = [];

  // Continue with original code...

  // ── Tambah col H, I & J jika belum ada ──────────────────────────────
  const hdrH = String(cfg.getRange(3,8).getValue());
  if (!hdrH.includes('Jira')) {
    cfg.insertColumnsAfter(7, 3);

    // H: Jira Sync (Yes/No)
    _hdrCell_(cfg, 3, 8, 'Jira Sync', 80, 'Yes = fetch bugs dari Jira\nNo = skip modul ini');

    // I: Jira Instance
    _hdrCell_(cfg, 3, 9, 'Jira Instance', 140, 'digitalperuri / bgn-peruri / -');

    // J: Jira Project Key
    _hdrCell_(cfg, 3, 10, 'Jira Project Key', 130, 'Contoh: TEST, SQA');

    // K: Link (existing column shifted)
    _hdrCell_(cfg, 3, 11, 'Link', 60, '');
    cfg.getRange(3,12).setBackground('#0D47A1').setFontColor('#FFFFFF')
      .setFontWeight('bold').setFontSize(9).setFontFamily('Arial');

    const lastRow = Math.max(cfg.getLastRow(), 4);

    // Data validation untuk Jira Sync (Yes/No)
    const dvSync = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Yes', 'No', '-'], true).build();
    cfg.getRange(4, 8, lastRow-3, 1).setDataValidation(dvSync).setValue('No');

    // Data validation untuk Jira Instance
    const dvInstance = SpreadsheetApp.newDataValidation()
      .requireValueInList(['digitalperuri','bgn-peruri','-'], true).build();
    cfg.getRange(4, 9, lastRow-3, 1).setDataValidation(dvInstance).setValue('-');

    // Default value untuk Project Key
    cfg.getRange(4, 10, lastRow-3, 1).setValue('-');

    log.push('✅ Kolom H (Jira Sync), I (Jira Instance) & J (Jira Project Key) ditambah');
  } else {
    log.push('ℹ️  Kolom Jira sudah ada, dilewati');
  }

  // ── Tambah section JIRA CREDENTIALS jika belum ada ───────────────
  const existing = cfg.getDataRange().getValues().map(r => String(r[0]));
  if (!existing.some(v => v.toUpperCase().includes('JIRA CREDENTIALS'))) {
    const sec = cfg.getLastRow() + 3;

    cfg.getRange(sec,1,1,12).merge()
      .setValue('JIRA CREDENTIALS  —  Isi email & token per instance Jira')
      .setBackground('#4A148C').setFontColor('#FFFFFF').setFontWeight('bold')
      .setFontSize(10).setFontFamily('Arial').setHorizontalAlignment('left');
    cfg.setRowHeight(sec, 26);

    cfg.getRange(sec+1,1,1,12).merge()
      .setValue('⚠️  Pastikan hanya owner yang bisa lihat tab ini. ' +
                'Jangan share spreadsheet ke publik jika token diisi di sini.')
      .setBackground('#F3E5F5').setFontColor('#6A1B9A').setFontStyle('italic').setFontSize(8);
    cfg.setRowHeight(sec+1, 16);

    const hdrs = [
      ['Instance',          '#6A1B9A', 140, 'Nilai: digitalperuri atau bgn-peruri'],
      ['Email Atlassian',   '#7B1FA2', 220, 'Email login ke id.atlassian.com'],
      ['API Token',         '#7B1FA2', 400,
        'Buat token di:\nhttps://id.atlassian.com/manage-profile/security/api-tokens'],
      ['Notif Email',       '#1565C0', 280,
        'Penerima notif harian jam 07.00.\nBisa multiple, pisah koma.'],
    ];
    hdrs.forEach(([h,bg,w,note], i) => {
      cfg.getRange(sec+2, i+1).setValue(h)
        .setBackground(bg).setFontColor('#FFFFFF').setFontWeight('bold')
        .setFontSize(9).setFontFamily('Arial')
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBorder(true,true,true,true,false,false,'#CE93D8',SpreadsheetApp.BorderStyle.SOLID);
      cfg.setColumnWidth(i+1, w);
      if (note) cfg.getRange(sec+2, i+1).setNote(note);
    });
    cfg.setRowHeight(sec+2, 22);

    [['digitalperuri','email@company.com','ATATT3xFf...(paste token disini)','qa@company.com'],
     ['bgn-peruri',   'email@company.com','ATATT3xFf...(paste token disini)',''],
    ].forEach(([inst,em,tok,notif], i) => {
      const r = sec+3+i;
      cfg.getRange(r,1,1,4).setValues([[inst,em,tok,notif]])
        .setBackground(i%2===0 ? '#F3E5F5' : '#FFFFFF')
        .setFontFamily('Arial').setFontSize(9).setVerticalAlignment('middle')
        .setBorder(true,true,true,true,false,false,'#CE93D8',SpreadsheetApp.BorderStyle.SOLID);
      cfg.getRange(r, 3).setFontFamily('Courier New').setFontSize(8);
      cfg.setRowHeight(r, 22);
    });

    log.push('✅ Section JIRA CREDENTIALS ditambah di bawah daftar modul');
  } else {
    log.push('ℹ️  Section JIRA CREDENTIALS sudah ada, dilewati');
  }

  // ── Tambah section GOOGLE CHAT NOTIFICATION jika belum ada ──────
  const existingChat = cfg.getDataRange().getValues().map(r => String(r[0]));
  if (!existingChat.some(v => v.toUpperCase().includes('GOOGLE CHAT'))) {
    const chatSec = cfg.getLastRow() + 3;

    cfg.getRange(chatSec,1,1,12).merge()
      .setValue('GOOGLE CHAT NOTIFICATION  —  Notifikasi Blocker Harian ke Google Chat')
      .setBackground('#1565C0').setFontColor('#FFFFFF').setFontWeight('bold')
      .setFontSize(10).setFontFamily('Arial').setHorizontalAlignment('left');
    cfg.setRowHeight(chatSec, 26);

    cfg.getRange(chatSec+1,1,1,12).merge()
      .setValue('💬  Notifikasi blocker otomatis dikirim ke Google Chat Space setiap hari. ' +
                'Atur webhook URL dan waktu notifikasi di bawah.')
      .setBackground('#E3F2FD').setFontColor('#1565C0').setFontStyle('italic').setFontSize(8);
    cfg.setRowHeight(chatSec+1, 16);

    const chatHdrs = [
      ['Google Chat Webhook URL',  '#1976D2', 450,
        'Buat webhook di Google Chat Space:\nSpace Settings > Apps & integrations > Webhooks\n\nFormat: https://chat.googleapis.com/v1/spaces/.../messages?key=...'],
      ['Notif Time (Hour)',        '#1976D2', 120,
        'Jam berapa notifikasi dikirim (0-23)\nContoh: 15 = jam 3 sore'],
      ['Enable Notifikasi',        '#1976D2', 140,
        'Yes = aktif notifikasi harian\nNo = nonaktifkan']
    ];
    chatHdrs.forEach(([h,bg,w,note], i) => {
      cfg.getRange(chatSec+2, i+1).setValue(h)
        .setBackground(bg).setFontColor('#FFFFFF').setFontWeight('bold')
        .setFontSize(9).setFontFamily('Arial')
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBorder(true,true,true,true,false,false,'#90CAF9',SpreadsheetApp.BorderStyle.SOLID);
      cfg.setColumnWidth(i+1, w);
      if (note) cfg.getRange(chatSec+2, i+1).setNote(note);
    });
    cfg.setRowHeight(chatSec+2, 22);

    // Data row with defaults
    cfg.getRange(chatSec+3,1,1,3).setValues([['https://chat.googleapis.com/v1/spaces/...', 15, 'No']])
      .setBackground('#E3F2FD')
      .setFontFamily('Arial').setFontSize(9).setVerticalAlignment('middle')
      .setBorder(true,true,true,true,false,false,'#90CAF9',SpreadsheetApp.BorderStyle.SOLID);
    cfg.getRange(chatSec+3, 1).setFontFamily('Courier New').setFontSize(8);

    // Data validation untuk Enable
    const dvEnable = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Yes', 'No'], true).build();
    cfg.getRange(chatSec+3, 3).setDataValidation(dvEnable);

    // Data validation untuk Hour (0-23)
    const dvHour = SpreadsheetApp.newDataValidation()
      .requireNumberBetween(0, 23).build();
    cfg.getRange(chatSec+3, 2).setDataValidation(dvHour);

    cfg.setRowHeight(chatSec+3, 22);

    log.push('✅ Section GOOGLE CHAT NOTIFICATION ditambah');
  } else {
    log.push('ℹ️  Section GOOGLE CHAT NOTIFICATION sudah ada, dilewati');
  }

    ss.setActiveSheet(cfg);
    safeAlert_(
      '✅ jiraSetup() selesai!\n\n' + log.join('\n') +
      '\n\n──────────────────────────────────' +
      '\nYang harus dilakukan sekarang:\n\n' +
      '1. Di tab Config, gulir ke bawah → section JIRA CREDENTIALS\n' +
      '2. Ganti email & token untuk digitalperuri dan bgn-peruri\n' +
      '3. Isi Notif Email (col D) → siapa yang terima notif harian\n' +
      '4. Per modul yang ingin di-sync:\n' +
      '   - Col H (Jira Sync): pilih "Yes"\n' +
      '   - Col I (Jira Instance): pilih instance (digitalperuri/bgn-peruri)\n' +
      '   - Col J (Jira Project Key): isi project key (contoh: SQA, TEST)\n\n' +
      '5. Di section GOOGLE CHAT NOTIFICATION:\n' +
      '   - Paste webhook URL dari Google Chat\n' +
      '   - Set jam notifikasi (0-23)\n' +
      '   - Enable = Yes untuk aktifkan\n\n' +
      'Setelah selesai → jalankan jiraActivate()'
    );

  } catch (e) {
    Logger.log('❌ jiraSetup error: ' + e.message);
    Logger.log('Error stack: ' + e.stack);
    safeAlert_('❌ Error saat setup:\n\n' + e.message +
               '\n\nCek Execution log untuk detail.\n\n' +
               'Pastikan Anda membuka Dashboard spreadsheet yang benar.');
  }
}


/**
 * LANGKAH 2: jiraActivate()
 * Jalankan setelah credentials & config modul sudah diisi.
 * Yang dilakukan:
 *   - Test koneksi ke semua instance Jira yang dikonfigurasi
 *   - Jika semua OK → aktifkan 3 triggers otomatis
 *   - Lakukan sync pertama sekarang
 */
function jiraActivate() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const mods  = _getJiraMods_(ss);
  const creds = _getCreds_(ss);

  if (mods.length === 0) {
    safeAlert_('Tidak ada modul aktif dengan Jira Sync = Yes.\n' +
               'Pastikan col H = "Yes" dan col I & J sudah diisi, lalu coba lagi.');
    return;
  }

  // ── 1. Test koneksi ────────────────────────────────────────────────
  const testResults = [];
  const tested = new Set();
  let allOk = true;

  mods.forEach(mod => {
    const key = mod.inst + '|' + mod.projKey;
    if (tested.has(key)) return;
    tested.add(key);

    const cred = creds[mod.inst];
    if (!cred) {
      testResults.push('❌ ' + mod.inst + ': credentials belum diisi di Config');
      allOk = false; return;
    }
    const jql  = 'project = "'+mod.projKey+'" AND issuetype = Bug ORDER BY updated DESC';
    const base = JIRA_INSTANCES_[mod.inst];
    const url  = base + '/rest/api/3/search/jql' +
      '?jql=' + encodeURIComponent(jql) +
      '&maxResults=1&fields=summary,status,priority';

    // Log JQL & URL untuk debugging
    Logger.log('');
    Logger.log('══════════════════════════════════════════');
    Logger.log('📌 TEST CONNECTION: ' + mod.name);
    Logger.log('══════════════════════════════════════════');
    Logger.log('Instance: ' + mod.inst);
    Logger.log('Project:  ' + mod.projKey);
    Logger.log('');
    Logger.log('JQL (decoded):');
    Logger.log(jql);
    Logger.log('');
    Logger.log('Browser URL (untuk test manual):');
    Logger.log(base + '/issues/?jql=' + encodeURIComponent(jql));
    Logger.log('');
    Logger.log('API Endpoint:');
    Logger.log(url);
    Logger.log('══════════════════════════════════════════');
    Logger.log('');

    try {
      const resp = UrlFetchApp.fetch(url, {
        headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(cred.email+':'+cred.token) },
        muteHttpExceptions: true
      });
      if (resp.getResponseCode() === 200) {
        const data = JSON.parse(resp.getContentText());
        const s = data.issues&&data.issues[0]
          ? data.issues[0].key+': '+(data.issues[0].fields.summary||'').substring(0,50)
          : '(no issues)';
        testResults.push('✅ '+mod.inst+'/'+mod.projKey+' → '+data.total+' bug aktif\n     '+s);
      } else {
        testResults.push('❌ '+mod.inst+'/'+mod.projKey+' HTTP '+resp.getResponseCode()+
          '\n   '+resp.getContentText().substring(0,150));
        allOk = false;
      }
    } catch(e) {
      testResults.push('❌ '+mod.inst+': '+e.message);
      allOk = false;
    }
  });

  if (!allOk) {
    safeAlert_('⚠️  Koneksi gagal di beberapa instance:\n\n' + testResults.join('\n\n') +
               '\n\nPerbaiki credentials di Config tab lalu jalankan jiraActivate() lagi.');
    return;
  }

  // ── 2. Aktifkan triggers ───────────────────────────────────────────
  ScriptApp.getProjectTriggers().forEach(t => {
    if (['syncJiraToAllSheets','syncJiraStatusAll','sendBugNotification'].includes(t.getHandlerFunction()))
      ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncJiraToAllSheets').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('sendBugNotification').timeBased().atHour(7).everyDays(1).create();
  ScriptApp.newTrigger('syncJiraStatusAll').timeBased().atHour(23).everyDays(1).create();

  // ── 3. Sync pertama ────────────────────────────────────────────────
  const syncResults = _runSync_(ss, false);

  safeAlert_(
    '✅ jiraActivate() selesai!\n\n' +
    '── Koneksi ──\n' + testResults.join('\n') +
    '\n\n── Triggers aktif ──\n' +
    '• Setiap 1 jam    : sync Title, Desc, Priority, Assignee\n' +
    '• Setiap 07.00    : kirim notif email pending bugs\n' +
    '• Setiap 23.00    : sync Status dari Jira\n' +
    '\n── Sync pertama ──\n' + syncResults.join('\n') +
    '\n\n── Modul yang di-sync ──\n' +
    mods.length + ' modul dengan Jira Sync = Yes'
  );
}


// ═══════════════════════════════════════════════════════════════════════
// MANUAL ENTRY POINTS
// ═══════════════════════════════════════════════════════════════════════

function syncJiraToAllSheets() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const res = _runSync_(ss, false);
  safeAlert_('Jira Sync Selesai\n\n' + res.join('\n'));
}

function syncJiraStatusAll() {
  _runSync_(SpreadsheetApp.getActiveSpreadsheet(), true);
  Logger.log('syncJiraStatusAll selesai');
}

function sendBugNotification() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const mods  = _getJiraMods_(ss);
  const creds = _getCreds_(ss);
  const bugs  = [];
  mods.forEach(mod => {
    try {
      const bugSh = SpreadsheetApp.openById(mod.id).getSheetByName('BugReport');
      if (bugSh) bugs.push(..._pendingBugs_(bugSh, mod.name));
    } catch(e) { Logger.log('Notif read ['+mod.name+']: '+e.message); }
  });
  if (bugs.length === 0) { safeAlert_('Tidak ada pending bug. 🎉'); return; }
  const emails = _notifEmails_(creds);
  if (!emails.length) { safeAlert_('Notif Email belum diisi di Config section JIRA CREDENTIALS (col D).'); return; }
  const subj = _notifSubj_(bugs);
  const body = _notifBody_(bugs, ss);
  emails.forEach(e => { try { MailApp.sendEmail({to:e, subject:subj, htmlBody:body}); } catch(ex){} });
  safeAlert_('Notifikasi terkirim ke:\n'+emails.join('\n')+'\n\nTotal: '+bugs.length+' pending bug');
}

function removeJiraTriggers() {
  let n=0;
  ScriptApp.getProjectTriggers().forEach(t=>{
    if(['syncJiraToAllSheets','syncJiraStatusAll','sendBugNotification'].includes(t.getHandlerFunction()))
      {ScriptApp.deleteTrigger(t);n++;}
  });
  safeAlert_(n+' trigger dihapus.');
}

/**
 * syncAllJira() - Sync Jira bugs untuk semua modul aktif
 * Called from menu: 🎯 QA Dashboard > 🔄 Jira Sync > Sync All Modules from Jira
 */
function syncAllJira() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Logger.log('═══════════════════════════════════════════');
  Logger.log('🔄 SYNC ALL JIRA - Manual trigger from menu');
  Logger.log('═══════════════════════════════════════════');

  try {
    // Use existing _runSync_ function (inclStatus = true to update status)
    const lines = _runSync_(ss, true);

    if (lines.length === 0) {
      safeAlert_('⚠️ Tidak ada modul dengan Jira Sync aktif.\n\nAktifkan Jira Sync di Config tab (kolom B).');
      return;
    }

    // Show result
    const msg = '✅ Jira sync selesai!\n\n' + lines.join('\n');
    safeAlert_(msg);

    Logger.log('──────────────────────────────────────────');
    Logger.log('✅ Jira sync completed');
    lines.forEach(line => Logger.log(line));
    Logger.log('══════════════════════════════════════════');

  } catch (e) {
    Logger.log('❌ Error during Jira sync: ' + e.message);
    Logger.log('Stack trace: ' + e.stack);
    safeAlert_('❌ Error saat sync Jira:\n\n' + e.message);
  }
}

/**
 * showJiraJQL() - Helper untuk lihat JQL yang digunakan
 * Jalankan ini untuk melihat JQL query untuk setiap modul yang aktif
 * Berguna untuk debugging dan customization
 */
function showJiraJQL() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mods = _getJiraMods_(ss);

  if (mods.length === 0) {
    safeAlert_('Tidak ada modul dengan Jira Sync = Yes');
    return;
  }

  let output = '═══ JIRA JQL CONFIGURATION ═══\n\n';
  output += 'Default JQL Template:\n';
  output += JIRA_JQL_ + '\n\n';
  output += '─────────────────────────────\n\n';

  mods.forEach(mod => {
    const modulField = JIRA_MODUL_FIELD_[mod.inst] || 'cf[10097]';
    const jql = JIRA_JQL_
      .replace('{P}', mod.projKey)
      .replace('{F}', modulField)
      .replace('{M}', mod.module || '');
    const base = JIRA_INSTANCES_[mod.inst];
    const url = base + '/rest/api/3/search/jql?jql=' + encodeURIComponent(jql) + '&fields=' + encodeURIComponent(JIRA_FIELDS_) + '&maxResults=100';

    output += '📌 Module: ' + mod.name + '\n';
    output += 'Instance: ' + mod.inst + '\n';
    output += 'Project: ' + mod.projKey + '\n';
    output += 'Modul: ' + (mod.module || '(not specified)') + '\n\n';
    output += 'JQL:\n' + jql + '\n\n';
    output += 'URL (untuk test di browser - remove auth):\n';
    output += base + '/issues/?jql=' + encodeURIComponent(jql) + '\n\n';
    output += 'API Endpoint:\n' + url + '\n\n';
    output += '─────────────────────────────\n\n';

    // Log ke console juga
    Logger.log('=== ' + mod.name + ' ===');
    Logger.log('Modul: ' + (mod.module || '(not specified)'));
    Logger.log('JQL: ' + jql);
    Logger.log('Browser URL: ' + base + '/issues/?jql=' + encodeURIComponent(jql));
    Logger.log('API URL: ' + url);
    Logger.log('');
  });

  output += '\n💡 TIPS CUSTOMIZATION:\n\n';
  output += '1. Field "modul[dropdown]" di JQL mengacu ke custom field di Jira\n';
  output += '   - Jika field name berbeda, edit JIRA_JQL_ di line 32-36\n';
  output += '   - Contoh alternatif: "Modul", cf[10001], customfield_10001\n\n';
  output += '2. Untuk filter tambahan per modul:\n';
  output += '   - Filter by component: AND component = "ModuleName"\n';
  output += '   - Filter by label: AND labels = "mobile"\n';
  output += '   - Multiple modul: AND "modul[dropdown]" IN (Modul1, Modul2)\n\n';
  output += '3. Lihat execution log untuk detail URL yang dipakai\n';
  output += '   (View > Execution log di Apps Script editor)\n';

  safeAlert_(output);
}


// ═══════════════════════════════════════════════════════════════════════
// CORE
// ═══════════════════════════════════════════════════════════════════════

function _runSync_(ss, inclStatus) {
  const mods  = _getJiraMods_(ss);
  const creds = _getCreds_(ss);
  const lines = [];

  if (mods.length === 0) {
    lines.push('ℹ️  Tidak ada modul dengan Jira Sync = Yes');
    return lines;
  }

  mods.forEach(mod => {
    try {
      const cred = creds[mod.inst];
      if (!cred) throw new Error('Credentials tidak ada untuk: '+mod.inst);
      lines.push('✅ '+mod.name+': '+_syncMod_(mod, cred, inclStatus));
    } catch(e) { lines.push('❌ '+mod.name+': '+e.message); }
    Utilities.sleep(400);
  });
  return lines;
}

function _syncMod_(mod, cred, inclStatus) {
  const src   = SpreadsheetApp.openById(mod.id);
  const bugSh = src.getSheetByName('BugReport');
  if (!bugSh) return 'skipped (no BugReport)';

  // Step 1: Clean all existing data (keep header rows 1-4)
  _cleanBugReportData_(bugSh);

  // Step 2: Fetch bugs from Jira
  const issues = _fetch_(mod.inst, mod.projKey, mod.module, cred);
  if (!issues) return 'failed (fetch error)';

  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const instUrl = JIRA_INSTANCES_[mod.inst];
  let ins = 0;

  // Step 3: Insert all bugs (skip Closed/Won't Fix)
  issues.forEach(iss => {
    const statusName = iss.fields.status && iss.fields.status.name;
    const isClosedStatus = statusName === 'Closed' || statusName === "Won't Fix";

    // Skip closed bugs - we don't want them in the sheet
    if (isClosedStatus) {
      return;
    }

    // Insert all active bugs as new rows
    _ins_(bugSh, iss, instUrl, now, mod, mod.inst);
    ins++;
  });

  return ins + ' bugs synced from Jira';
}

function _fetch_(instKey, projKey, modulName, cred) {
  const base = JIRA_INSTANCES_[instKey];
  const modulField = JIRA_MODUL_FIELD_[instKey] || 'cf[10097]';  // Default to digitalperuri field
  const customFields = JIRA_CUSTOM_FIELDS_[instKey] || '';  // Custom fields per instance

  // Check if modulName is numeric - if so, don't quote it in JQL
  const isNumeric = /^\d+$/.test(modulName);
  const modulValue = isNumeric ? modulName : '"' + modulName + '"';

  let jql  = JIRA_JQL_
    .replace('{P}', projKey)
    .replace('{F}', modulField)
    .replace('"{M}"', modulValue);  // Replace including quotes
  const auth = Utilities.base64Encode(cred.email+':'+cred.token);
  const hdrs = {'Authorization':'Basic '+auth,'Content-Type':'application/json'};
  const all = [];
  let total = 0;
  const fields = JIRA_FIELDS_ + customFields;  // Append custom fields

  // Log JQL & URL untuk debugging (hanya di iterasi pertama)
  Logger.log('');
  Logger.log('══════════════════════════════════════════');
  Logger.log('FETCHING BUGS FROM JIRA');
  Logger.log('══════════════════════════════════════════');
  Logger.log('Instance: ' + instKey);
  Logger.log('Project:  ' + projKey);
  Logger.log('Modul:    ' + (modulName || '(not specified)'));
  Logger.log('');
  Logger.log('JQL (decoded):');
  Logger.log(jql);
  Logger.log('');
  Logger.log('Browser URL (untuk test manual):');
  Logger.log(base + '/issues/?jql=' + encodeURIComponent(jql));
  Logger.log('');

  // New pagination using nextPageToken (not startAt)
  let nextPageToken = null;
  let isFirstPage = true;

  do {
    // Build URL with nextPageToken for pagination (new API format)
    let url = base + '/rest/api/3/search/jql?jql=' + encodeURIComponent(jql) +
      '&fields=' + encodeURIComponent(fields) + '&maxResults=100';

    if (nextPageToken) {
      url += '&nextPageToken=' + encodeURIComponent(nextPageToken);
    }

    if (isFirstPage) {
      Logger.log('API Endpoint:');
      Logger.log(url);
      Logger.log('══════════════════════════════════════════');
      Logger.log('');
    }

    let r;
    try { r = UrlFetchApp.fetch(url, {headers: hdrs, muteHttpExceptions: true}); }
    catch (e) { Logger.log('fetch err:' + e.message); return null; }

    if (r.getResponseCode() !== 200) {
      Logger.log('Jira ' + r.getResponseCode() + ':' + r.getContentText().substring(0, 200));
      return null;
    }

    let d;
    try { d = JSON.parse(r.getContentText()); }
    catch (e) { return null; }

    total = d.total || 0;
    (d.issues || []).forEach(i => all.push(i));

    if (isFirstPage) {
      Logger.log('✅ Found ' + total + ' bug(s) matching JQL');

      // Debug: Log raw response when 0 results to troubleshoot
      if (total === 0) {
        Logger.log('');
        Logger.log('⚠️  ZERO RESULTS - Debug Info:');
        Logger.log('Raw API Response (first 1000 chars):');
        Logger.log(r.getContentText().substring(0, 1000));
        Logger.log('');
        Logger.log('💡 Possible causes:');
        Logger.log('1. Custom field "modul[dropdown]" might need field ID instead (e.g., customfield_10001)');
        Logger.log('2. Module value "' + modulName + '" might not match exactly (check case/spaces)');
        Logger.log('3. Field permissions might have changed');
        Logger.log('');
        Logger.log('To find custom field ID:');
        Logger.log('• Go to: ' + base + '/rest/api/3/field');
        Logger.log('• Search for: modul');
        Logger.log('• Use the "id" value (e.g., customfield_10001) instead of "modul[dropdown]"');
        Logger.log('');
      }
      isFirstPage = false;
    }

    // Get next page token for pagination (new API format)
    nextPageToken = d.nextPageToken || null;

    // Sleep to avoid rate limiting
    if (nextPageToken) {
      Utilities.sleep(300);
    }

  } while (nextPageToken);

  Logger.log('✅ Fetched ' + all.length + ' bug(s) total');
  Logger.log('');
  return all;
}

/**
 * Clean all data rows from BugReport sheet (keep header rows 1-4)
 * Simple and fast - delete everything from row 5 onwards
 */
function _cleanBugReportData_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow >= BUG_START_) {
    sh.deleteRows(BUG_START_, lastRow - BUG_START_ + 1);
    Logger.log('Cleaned ' + (lastRow - BUG_START_ + 1) + ' existing rows from BugReport');
  }
}

function _ins_(sh,iss,instUrl,now,mod,instKey){
  const f=iss.fields,key=iss.key,type=_type_(iss);
  const fieldMap = JIRA_FIELD_MAP_[instKey] || JIRA_FIELD_MAP_['digitalperuri'];  // Default to digitalperuri
  const row=new Array(BUG_COLS_).fill('');
  row[BC_.BUG_ID-1]   = key;  // Use Jira Key as Bug ID (e.g., TEST-3017)
  row[BC_.TYPE-1]     = type;
  row[BC_.PRIORITY-1] = _prio_(f.priority&&f.priority.name)||'';
  row[BC_.STATUS-1]   = _stat_(f.status&&f.status.name)||'Open';
  row[BC_.FEATURE-1]  = f[fieldMap.feature]||'';  // Feature from Jira
  row[BC_.SUBMODUL-1] = _dropdown_(f[fieldMap.submodul]) || mod.submodule || '';  // Submodul from Jira, fallback to Config
  row[BC_.TITLE-1]    = f.summary||'';
  row[BC_.DESC-1]     = _adf_(f.description);

  // Use field map for custom fields with proper extraction
  row[BC_.ENV-1]      = _dropdown_(f[fieldMap.environment]);
  row[BC_.STEPS-1]    = _adf_(f[fieldMap.steps]);
  row[BC_.EXP-1]      = _adf_(f[fieldMap.expected]);
  row[BC_.ACT-1]      = _adf_(f[fieldMap.actual]);

  row[BC_.NOTES-1]    = '';  // Notes - dikosongkan
  row[BC_.ASSIGNED-1] = (f.assignee&&f.assignee.displayName)||'';
  row[BC_.REPORTED_BY-1] = (f.reporter&&f.reporter.displayName)||'';
  row[BC_.DATE_FOUND-1]= f.created?new Date(f.created):'';
  row[BC_.DATE_FIXED-1]= f.resolutiondate?new Date(f.resolutiondate):'';
  row[BC_.LINK-1]       = instUrl+'/browse/'+key;
  // BC_.NOTES already set above (line 806)
  row[BC_.SCREENSHOT-1] = _extractScreenshots_(iss, instKey) || '';  // Extract screenshot URLs from custom field (Col U)
  row[BC_.JIRA_KEY-1]   = key;
  row[BC_.SYNCED-1]     = now;
  const nr=Math.max(sh.getLastRow(),BUG_START_-1)+1;
  sh.getRange(nr,1,1,BUG_COLS_).setValues([row]);
  const bg=(nr-BUG_START_)%2===0?'#F8FBFF':'#FFFFFF';
  sh.getRange(nr,1,1,BUG_COLS_).setBackground(bg).setFontFamily('Arial').setFontSize(9)
    .setVerticalAlignment('middle')
    .setBorder(true,true,true,true,false,false,'#90CAF9',SpreadsheetApp.BorderStyle.SOLID);
  if(f.created)       sh.getRange(nr,BC_.DATE_FOUND).setNumberFormat('yyyy-mm-dd');
  if(f.resolutiondate)sh.getRange(nr,BC_.DATE_FIXED).setNumberFormat('yyyy-mm-dd');

  // Make Bug ID clickable with hyperlink to Jira
  const jiraLink = instUrl + '/browse/' + key;
  sh.getRange(nr,BC_.BUG_ID).setFormula('=HYPERLINK("' + jiraLink + '","' + key + '")');
  sh.getRange(nr,BC_.BUG_ID).setFontColor('#1565C0').setFontWeight('bold');

  sh.getRange(nr,BC_.JIRA_KEY).setFontColor('#1565C0').setFontWeight('bold');
  sh.getRange(nr,BC_.LINK).setFontColor('#1A73E8');
}

/**
 * Extract screenshot/video URLs from Jira custom field "Screenshot/ Video"
 * @param {Object} iss - Jira issue object
 * @param {string} instKey - Jira instance key (digitalperuri or bgn-peruri)
 * @returns {string} - Screenshot/Video URLs from custom field, or empty string
 */
function _extractScreenshots_(iss, instKey) {
  if (!iss || !iss.fields) return '';

  const fieldMap = JIRA_FIELD_MAP_[instKey] || JIRA_FIELD_MAP_['digitalperuri'];
  const screenshotFieldId = fieldMap.screenshot;
  if (!screenshotFieldId) return '';

  const screenshotValue = iss.fields[screenshotFieldId];
  if (!screenshotValue) return '';

  let textValue = '';
  if (typeof screenshotValue === 'string') {
    textValue = screenshotValue.trim();
  } else if (screenshotValue && screenshotValue.content) {
    textValue = _adf_(screenshotValue);
  }

  return textValue || '';
}


// ═══════════════════════════════════════════════════════════════════════
// NOTIFICATION
// ═══════════════════════════════════════════════════════════════════════

function _pendingBugs_(sh,modName){
  const last=sh.getLastRow(); if(last<BUG_START_)return[];
  const rows=sh.getRange(BUG_START_,1,last-BUG_START_+1,BUG_COLS_).getValues();
  const PEND=['open','in progress','reopen'];
  return rows.filter(r=>{
    const st=String(r[BC_.STATUS-1]).trim().toLowerCase();
    const ti=String(r[BC_.TITLE-1]).trim();
    const jk=String(r[BC_.JIRA_KEY-1]).trim();
    return (ti||jk)&&PEND.includes(st);
  }).map(r=>({
    module:modName,
    bugId: String(r[BC_.BUG_ID-1]).trim()||String(r[BC_.JIRA_KEY-1]).trim(),
    priority:String(r[BC_.PRIORITY-1]).trim()||'-',
    status:  String(r[BC_.STATUS-1]).trim()||'-',
    title:   String(r[BC_.TITLE-1]).trim()||'(no title)',
    link:    String(r[BC_.LINK-1]).trim(),
  }));
}

function _notifSubj_(bugs){
  const c=bugs.filter(b=>b.priority==='Critical').length;
  const h=bugs.filter(b=>b.priority==='High').length;
  const f=c>0?'🔴':h>0?'🟠':'🟡';
  return f+' [QA Alert] '+bugs.length+' Pending Bug'+(bugs.length>1?'s':'')+
    (c?' — '+c+' Critical':'')+((c&&h)?',':'')+(h?' '+h+' High':''  );
}

function _notifBody_(bugs,ss){
  const dashUrl=ss?'https://docs.google.com/spreadsheets/d/'+ss.getId():'';
  const now=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'dd MMM yyyy HH:mm');
  const PO={'Critical':0,'High':1,'Medium':2,'Low':3};
  bugs.sort((a,b)=>(PO[a.priority]||9)-(PO[b.priority]||9));
  const PBG={'Critical':'#FFCDD2','High':'#FFE0B2','Medium':'#E3F2FD','Low':'#F1F8E9'};
  const PFG={'Critical':'#B71C1C','High':'#E65100','Medium':'#1565C0','Low':'#388E3C'};
  const SBG={'Open':'#FFCDD2','In Progress':'#E3F2FD','Reopen':'#EDE7F6','Fixed':'#FFF9C4','Verified':'#C8E6C9','In Progress VAPT':'#E1F5FE','Done VAPT':'#B2DFDB'};
  const SFG={'Open':'#C62828','In Progress':'#1565C0','Reopen':'#6A1B9A','Fixed':'#F57F17','Verified':'#2E7D32','In Progress VAPT':'#01579B','Done VAPT':'#004D40'};
  const cnt={}; bugs.forEach(b=>{cnt[b.priority]=(cnt[b.priority]||0)+1;});
  const badges=['Critical','High','Medium','Low'].filter(p=>cnt[p])
    .map(p=>`<span style="background:${PBG[p]};color:${PFG[p]};font-weight:bold;padding:3px 10px;border-radius:12px;font-size:12px;margin-right:6px;">${p}: ${cnt[p]}</span>`).join('');
  const byMod={};
  bugs.forEach(b=>{if(!byMod[b.module])byMod[b.module]=[];byMod[b.module].push(b);});
  let trs='';
  Object.entries(byMod).forEach(([m,mb])=>{
    trs+=`<tr><td colspan="5" style="background:#1565C0;color:#fff;font-weight:bold;font-size:11px;padding:6px 10px;border-bottom:2px solid #0D47A1;">📁 ${m} (${mb.length} bug)</td></tr>`;
    mb.forEach((b,i)=>{
      const bg=i%2===0?'#F8FBFF':'#FFFFFF';
      const lnk=b.link?`<a href="${b.link}" style="color:#1A73E8;font-weight:bold;text-decoration:none;">${b.bugId}</a>`:b.bugId;
      trs+=`<tr style="background:${bg};">
        <td style="padding:7px 10px;border-bottom:1px solid #E0E0E0;font-size:12px;">${lnk}</td>
        <td style="padding:7px 10px;border-bottom:1px solid #E0E0E0;text-align:center;"><span style="background:${PBG[b.priority]||''};color:${PFG[b.priority]||''};font-weight:bold;padding:2px 8px;border-radius:10px;font-size:11px;">${b.priority}</span></td>
        <td style="padding:7px 10px;border-bottom:1px solid #E0E0E0;text-align:center;"><span style="background:${SBG[b.status]||''};color:${SFG[b.status]||''};font-weight:bold;padding:2px 8px;border-radius:10px;font-size:11px;">${b.status}</span></td>
        <td style="padding:7px 10px;border-bottom:1px solid #E0E0E0;font-size:12px;max-width:380px;">${b.title}</td>
        <td style="padding:7px 10px;border-bottom:1px solid #E0E0E0;text-align:center;">${b.link?`<a href="${b.link}" style="color:#1A73E8;font-size:11px;text-decoration:none;">Lihat ↗</a>`:'—'}</td>
      </tr>`;
    });
  });
  return `<!DOCTYPE html><html><body style="margin:0;padding:0;background:#F5F7FA;font-family:Arial,sans-serif;">
<div style="max-width:800px;margin:24px auto;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.12);">
  <div style="background:#0D47A1;padding:20px 24px;">
    <div style="color:#fff;font-size:18px;font-weight:bold;">🐛 QA Bug Notification</div>
    <div style="color:#90CAF9;font-size:12px;margin-top:4px;">${now}  ·  ${bugs.length} pending bug</div>
  </div>
  <div style="padding:16px 24px;background:#E3F2FD;border-bottom:1px solid #BBDEFB;">
    <div style="font-size:11px;color:#1565C0;font-weight:bold;margin-bottom:8px;">RINGKASAN PRIORITAS</div>
    ${badges}
  </div>
  <table style="width:100%;border-collapse:collapse;">
    <thead><tr style="background:#37474F;">
      <th style="padding:8px 10px;color:#fff;font-size:11px;text-align:left;width:110px;">Bug / Key</th>
      <th style="padding:8px 10px;color:#fff;font-size:11px;text-align:center;width:85px;">Priority</th>
      <th style="padding:8px 10px;color:#fff;font-size:11px;text-align:center;width:105px;">Status</th>
      <th style="padding:8px 10px;color:#fff;font-size:11px;text-align:left;">Title</th>
      <th style="padding:8px 10px;color:#fff;font-size:11px;text-align:center;width:65px;">Link</th>
    </tr></thead>
    <tbody>${trs}</tbody>
  </table>
  <div style="padding:12px 24px;background:#F5F7FA;border-top:1px solid #E0E0E0;font-size:11px;color:#78909C;">
    ${dashUrl?`<a href="${dashUrl}" style="color:#1A73E8;text-decoration:none;">📊 Buka QA Dashboard</a>  ·  `:''}
    Dikirim otomatis oleh QA Portfolio Dashboard.
  </div>
</div></body></html>`;
}


// ═══════════════════════════════════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════════════════════════════════

function _getCreds_(ss){
  const cred=ss.getSheetByName('Credentials');
  if(!cred)return{};
  const data=cred.getDataRange().getValues();
  const c={};

  // Data dimulai dari row 4 (row 1=title, 2=warning, 3=header)
  for(let i=3;i<data.length;i++){
    const inst=String(data[i][0]).trim().toLowerCase();
    const em=String(data[i][1]).trim();
    const tok=String(data[i][2]).trim();
    const notif=String(data[i][3]).trim();

    if(!inst||!JIRA_INSTANCES_[inst])continue;
    if(em&&tok){
      c[inst]={email:em,token:tok};
      if(notif&&notif.includes('@')){
        if(!c._notif)c._notif=[];
        notif.split(',').forEach(e=>{const t=e.trim();if(t)c._notif.push(t);});
      }
    }
  }
  return c;
}

function _notifEmails_(c){return c._notif||[];}

/**
 * _getJiraMods_
 * Ambil modul yang:
 *   - Col A (Active) = TRUE
 *   - Col B (Jira Sync) = TRUE
 *   - Col I (Jira Instance) ada dan bukan '-'
 *   - Col J (Jira Project) ada dan bukan '-'
 */
function _getJiraMods_(ss){
  const cfg=ss.getSheetByName('Config'); if(!cfg)return[];
  const data=cfg.getDataRange().getValues(); const mods=[];
  for(let i=3;i<data.length;i++){
    // Cek Active (col A) = TRUE
    if(data[i][0] !== true) continue;

    // Cek Jira Sync (col B) = TRUE
    if(data[i][1] !== true) continue;

    const project   = String(data[i][2]).trim();  // Col C = Project
    const module    = String(data[i][3]).trim();  // Col D = Modul
    const submodule = String(data[i][4]).trim();  // Col E = Submodul
    const team      = String(data[i][5]).trim();  // Col F = PIC QA
    const id        = String(data[i][6]).trim();  // Col G = Spreadsheet ID
    const inst      = String(data[i][8]).trim().toLowerCase();  // Col I = Jira Instance
    const pk        = String(data[i][9]).trim().toUpperCase();  // Col J = Jira Project

    if(!id||id.length<10||id==='PASTE_SPREADSHEET_ID_HERE')continue;
    if(!inst||!pk||inst==='-'||pk==='-')continue;
    if(!JIRA_INSTANCES_[inst]){Logger.log('Unknown instance: '+inst);continue;}

    mods.push({
      name:submodule||module||'Unknown',
      id,
      project,
      module,
      submodule,
      team,
      lead: '',  // QA Lead tidak ada di Config
      inst,
      projKey:pk
    });
  }
  return mods;
}

function _hdrCell_(cfg,r,c,label,w,note){
  cfg.getRange(r,c).setValue(label)
    .setBackground('#0D47A1').setFontColor('#FFFFFF').setFontWeight('bold')
    .setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('center').setVerticalAlignment('middle');
  cfg.setColumnWidth(c,w);
  if(note)cfg.getRange(r,c).setNote(note);
}

function _prio_(p){
  if(!p)return'';
  return({'Highest':'Critical','Critical':'Critical','High':'High',
          'Medium':'Medium','Low':'Low','Lowest':'Low','Minor':'Low','Trivial':'Low'})[p]||p;
}
function _stat_(s){
  if(!s)return''; const sl=s.toLowerCase();
  if(['open','to do','backlog','new'].includes(sl))               return'Open';
  if(['in progress','in review','review','testing'].includes(sl)) return'In Progress';
  if(['fixed','ready for qa','ready for review'].includes(sl))     return'Fixed';
  if(['verified','qa verified'].includes(sl))                       return'Verified';
  if(['closed','done'].includes(sl))                                return'Closed';
  if(["won't fix",'wontfix','not a bug','invalid'].includes(sl))   return"Won't Fix";
  if(['reopened','reopen'].includes(sl))                            return'Reopen';
  return'Open';
}
function _adf_(adf){
  if(!adf)return''; if(typeof adf==='string')return adf;
  function ex(n){
    if(!n)return'';
    if(n.type==='text')return n.text||'';
    if(n.type==='hardBreak')return'\n';
    if(n.type==='paragraph')return(n.content||[]).map(ex).join('')+'\n';
    if(n.type==='bulletList'||n.type==='orderedList')
      return(n.content||[]).map((it,i)=>(n.type==='orderedList'?(i+1)+'. ':' • ')+
        (it.content||[]).map(ex).join('')).join('\n')+'\n';
    // Handle inlineCard (embedded links with URL in attrs)
    if(n.type==='inlineCard'&&n.attrs&&n.attrs.url){
      return n.attrs.url;  // Extract URL from inlineCard
    }
    // Handle media nodes (images, attachments) - skip them or add placeholder
    if(n.type==='mediaSingle'||n.type==='mediaInline'||n.type==='media'){
      return'[Image]';  // Placeholder for images
    }
    // Handle other node types with content
    if(n.content)return n.content.map(ex).join('');
    return'';
  }
  try{
    const result = ex(adf).trim();
    // If result is empty or only contains [Image], return empty string
    if(!result || result==='[Image]')return '';
    return result.substring(0,1000);
  }catch(e){
    Logger.log('Error parsing ADF: '+e.message);
    return'';
  }
}
function _dropdown_(field){
  // Extract value from Jira dropdown field object (e.g., Environment field)
  if(!field)return'';
  if(typeof field==='string')return field;
  if(field.value)return field.value;  // Dropdown field has .value property
  return'';
}
function _type_(iss){
  const lb=(iss.fields.labels||[]).map(l=>l.toLowerCase());
  if(lb.some(l=>['api','backend','be'].includes(l)))       return'API';
  if(lb.some(l=>['mobile','android','ios'].includes(l)))   return'Mobile';
  const cp=(iss.fields.components||[]).map(c=>(c.name||'').toLowerCase());
  if(cp.some(c=>c.includes('api')||c.includes('backend')))   return'API';
  if(cp.some(c=>c.includes('mobile')||c.includes('android')))return'Mobile';
  return'Web';
}
function _bugId_(sh,type){
  const pre=type==='API'?'BUG-API':type==='Mobile'?'BUG-MOB':'BUG-WEB';
  const last=sh.getLastRow(); if(last<BUG_START_)return pre+'-001';
  const nums=sh.getRange(BUG_START_,BC_.BUG_ID,last-BUG_START_+1,1).getValues()
    .map(r=>String(r[0])).filter(id=>id.startsWith(pre))
    .map(id=>parseInt(id.split('-').pop())||0);
  return pre+'-'+String((nums.length?Math.max(...nums):0)+1).padStart(3,'0');
}


// ═══════════════════════════════════════════════════════════════════════
// GOOGLE CHAT NOTIFICATION  —  Notifikasi Blocker Harian
// ═══════════════════════════════════════════════════════════════════════

/**
 * sendBlockerNotification()
 * Manual entry point untuk kirim notifikasi blocker ke Google Chat
 * Akan otomatis dipanggil oleh trigger harian
 */
function sendBlockerNotification() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const chatCfg = _getChatConfig_(ss);
  const emailCfg = _getEmailConfig_(ss);

  // Check if at least one notification method is enabled
  const chatEnabled = chatCfg && chatCfg.enabled && chatCfg.webhookUrl && !chatCfg.webhookUrl.includes('...');
  const emailEnabled = emailCfg && emailCfg.enabled && emailCfg.recipients && emailCfg.recipients.length > 0;

  if (!chatEnabled && !emailEnabled) {
    Logger.log('No notification methods enabled');
    safeAlert_('⚠️  Tidak ada notifikasi aktif.\n\nAktifkan Google Chat atau Email di Config tab.');
    return;
  }

  const mods = _getJiraMods_(ss);
  if (mods.length === 0) {
    Logger.log('No active modules to check blockers');
    return;
  }

  const blockerData = _getBlockerData_(mods);

  // Filter only modules with blockers
  const modulesWithBlockers = blockerData.filter(m => m.totalBlocker > 0);

  Logger.log('Found ' + blockerData.length + ' total modules, ' + modulesWithBlockers.length + ' with blockers');

  let sentCount = 0;
  let errors = [];

  // If no blockers at all, send one summary message
  if (modulesWithBlockers.length === 0) {
    Logger.log('No blockers found - sending all clear message');

    // Send Google Chat
    if (chatEnabled) {
      try {
        const message = _buildChatMessageForModule_(null, ss);
        const response = UrlFetchApp.fetch(chatCfg.webhookUrl, {
          method: 'post',
          contentType: 'application/json',
          payload: JSON.stringify(message),
          muteHttpExceptions: true
        });

        if (response.getResponseCode() === 200) {
          Logger.log('✅ All clear notification sent to Google Chat');
          sentCount++;
        } else {
          errors.push('Google Chat: HTTP ' + response.getResponseCode());
        }
      } catch (e) {
        errors.push('Google Chat: ' + e.message);
      }
    }

    // Send Email
    if (emailEnabled) {
      try {
        const subject = '✅ QA Blocker Status - All Clear';
        const htmlBody = _buildEmailBodyForModule_(null, ss);

        MailApp.sendEmail({
          to: emailCfg.recipients.join(','),
          subject: subject,
          htmlBody: htmlBody
        });

        Logger.log('✅ All clear email sent to ' + emailCfg.recipients.length + ' recipient(s)');
        sentCount++;
      } catch (e) {
        errors.push('Email: ' + e.message);
      }
    }
  } else {
    // Send notification PER MODULE
    Logger.log('Sending ' + modulesWithBlockers.length + ' separate notifications (one per module)');

    modulesWithBlockers.forEach((moduleData, idx) => {
      Logger.log('Processing module ' + (idx + 1) + '/' + modulesWithBlockers.length + ': ' + moduleData.module);

      // Send Google Chat for this module
      if (chatEnabled) {
        try {
          const message = _buildChatMessageForModule_(moduleData, ss);
          const response = UrlFetchApp.fetch(chatCfg.webhookUrl, {
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify(message),
            muteHttpExceptions: true
          });

          if (response.getResponseCode() === 200) {
            Logger.log('✅ Google Chat notification sent for ' + moduleData.module);
            sentCount++;
          } else {
            errors.push(moduleData.module + ' (Chat): HTTP ' + response.getResponseCode());
          }

          // Small delay to avoid rate limiting
          Utilities.sleep(500);
        } catch (e) {
          Logger.log('❌ Error sending Google Chat for ' + moduleData.module + ': ' + e.message);
          errors.push(moduleData.module + ' (Chat): ' + e.message);
        }
      }

      // Send Email for this module
      if (emailEnabled) {
        try {
          const subject = '🚨 QA Blocker Alert - ' + moduleData.project + ' > ' + moduleData.module +
                         ' (' + moduleData.totalBlocker + ' Blocker' + (moduleData.totalBlocker > 1 ? 's' : '') + ')';

          const htmlBody = _buildEmailBodyForModule_(moduleData, ss);

          MailApp.sendEmail({
            to: emailCfg.recipients.join(','),
            subject: subject,
            htmlBody: htmlBody
          });

          Logger.log('✅ Email notification sent for ' + moduleData.module + ' to ' + emailCfg.recipients.length + ' recipient(s)');
          sentCount++;

          // Small delay to avoid rate limiting
          Utilities.sleep(300);
        } catch (e) {
          Logger.log('❌ Error sending email for ' + moduleData.module + ': ' + e.message);
          errors.push(moduleData.module + ' (Email): ' + e.message);
        }
      }
    });
  }

  // Show result
  if (sentCount > 0) {
    const totalBlocker = blockerData.reduce((sum, m) => sum + m.totalBlocker, 0);
    safeAlert_('✅ Notifikasi blocker berhasil dikirim!\n\n' +
               'Total modul dengan blocker: ' + modulesWithBlockers.length + '\n' +
               'Total blocker: ' + totalBlocker + '\n' +
               'Notifikasi terkirim: ' + sentCount +
               (errors.length > 0 ? '\n\n⚠️ Sebagian gagal:\n' + errors.join('\n') : ''));
  } else {
    safeAlert_('❌ Gagal kirim notifikasi!\n\n' + errors.join('\n'));
  }
}

/**
 * setupDailyBlockerNotification()
 * Setup trigger harian untuk notifikasi blocker
 * Waktu diambil dari Config
 */
function setupDailyBlockerNotification() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const chatCfg = _getChatConfig_(ss);

  if (!chatCfg) {
    safeAlert_('⚠️  Google Chat config tidak ditemukan.\n\nJalankan jiraSetup() dulu.');
    return;
  }

  // Delete existing trigger
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendBlockerNotification') {
      ScriptApp.deleteTrigger(t);
    }
  });

  if (!chatCfg.enabled) {
    safeAlert_('ℹ️  Trigger blocker notification dihapus.\n\n' +
               'Set "Enable Notifikasi" = Yes di Config untuk aktifkan kembali.');
    return;
  }

  if (!chatCfg.webhookUrl || chatCfg.webhookUrl.includes('...')) {
    safeAlert_('⚠️  Webhook URL belum dikonfigurasi!\n\n' +
               'Isi webhook URL di Config sebelum setup trigger.');
    return;
  }

  // Create new trigger
  const hour = chatCfg.notifHour;
  ScriptApp.newTrigger('sendBlockerNotification')
    .timeBased()
    .atHour(hour)
    .everyDays(1)
    .create();

  Logger.log('✅ Trigger created for daily blocker notification at ' + hour + ':00');

  safeAlert_('✅ Trigger blocker notification berhasil dibuat!\n\n' +
             'Waktu: Setiap hari jam ' + hour + ':00\n' +
             'Target: Google Chat\n\n' +
             'Untuk test sekarang, jalankan sendBlockerNotification()');
}

/**
 * removeDailyBlockerNotification()
 * Hapus trigger notifikasi blocker
 */
function removeDailyBlockerNotification() {
  let deleted = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendBlockerNotification') {
      ScriptApp.deleteTrigger(t);
      deleted++;
    }
  });
  safeAlert_(deleted + ' trigger blocker notification dihapus.');
}

function _getChatConfig_(ss) {
  const cfg = ss.getSheetByName('Config');
  if (!cfg) return null;

  // Google Chat config is at row 4, columns L-N (12-14)
  // L4 = Webhook URL, M4 = Notif Hour, N4 = Enable
  const chatRow = 4;
  const chatColStart = 12; // Column L

  try {
    const webhookUrl = String(cfg.getRange(chatRow, chatColStart).getValue() || '').trim();
    const notifHour = parseInt(cfg.getRange(chatRow, chatColStart + 1).getValue()) || 15;
    const enabledVal = cfg.getRange(chatRow, chatColStart + 2).getValue();

    // Support both boolean (TRUE/FALSE checkbox) and string ('Yes'/'No')
    const enabled = (typeof enabledVal === 'boolean') ? enabledVal : String(enabledVal).trim().toLowerCase() === 'yes';

    // Validate webhook URL exists
    if (!webhookUrl || webhookUrl.includes('...')) {
      Logger.log('⚠️ Webhook URL belum diisi di Config L4');
      return null;
    }

    return { webhookUrl, notifHour, enabled };
  } catch (e) {
    Logger.log('❌ Error reading Google Chat config: ' + e.message);
    return null;
  }
}

function _getEmailConfig_(ss) {
  const cfg = ss.getSheetByName('Config');
  if (!cfg) return null;

  // Email config is at row 4, columns O-P (15-16)
  // O4 = Email Recipients, P4 = Enable
  const emailRow = 4;
  const emailColStart = 15; // Column O

  try {
    const recipients = String(cfg.getRange(emailRow, emailColStart).getValue() || '').trim();
    const enabledVal = cfg.getRange(emailRow, emailColStart + 1).getValue();

    // Support both boolean and string
    const enabled = (typeof enabledVal === 'boolean') ? enabledVal : String(enabledVal).trim().toLowerCase() === 'yes';

    // Validate recipients
    if (!recipients || recipients.includes('example.com')) {
      Logger.log('⚠️ Email recipients belum diisi di Config O4');
      return null;
    }

    // Parse comma-separated emails
    const emailList = recipients.split(',').map(e => e.trim()).filter(e => e);

    return { recipients: emailList, enabled };
  } catch (e) {
    Logger.log('❌ Error reading Email config: ' + e.message);
    return null;
  }
}

function _buildEmailBody_(blockerData, ss) {
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm');
  const totalCritical = blockerData.reduce((sum, m) => sum + m.critical, 0);
  const totalHigh = blockerData.reduce((sum, m) => sum + m.high, 0);
  const totalMedium = blockerData.reduce((sum, m) => sum + m.medium, 0);
  const totalBlocker = totalCritical + totalHigh + totalMedium;

  const overviewSheet = ss.getSheetByName('Overview');
  const overviewGid = overviewSheet ? overviewSheet.getSheetId() : 0;
  const dashboardUrl = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/edit#gid=' + overviewGid;

  let statusColor, statusIcon, statusMessage;
  if (totalBlocker === 0) {
    statusColor = '#4CAF50';
    statusIcon = '✅';
    statusMessage = 'Tidak ada blocker! Semua modul dalam kondisi baik. Great job team!';
  } else if (totalBlocker <= 3) {
    statusColor = '#FF9800';
    statusIcon = '⚠️';
    statusMessage = totalBlocker + ' blocker terdeteksi — Kondisi masih terkendali, perlu perhatian.';
  } else if (totalCritical >= 5) {
    statusColor = '#D32F2F';
    statusIcon = '🔥';
    statusMessage = totalBlocker + ' blocker (' + totalCritical + ' CRITICAL)! — URGENT: Tim harus fokus resolve blocker critical!';
  } else {
    statusColor = '#F44336';
    statusIcon = '🚨';
    statusMessage = totalBlocker + ' blocker terdeteksi! — Perlu tindakan segera.';
  }

  let html = `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: Arial, sans-serif; margin: 0; padding: 20px; background: #f5f5f5; }
    .container { max-width: 600px; margin: 0 auto; background: white; border-radius: 8px; overflow: hidden; }
    .header { background: ${statusColor}; color: white; padding: 20px; text-align: center; }
    .header h1 { margin: 0; font-size: 24px; }
    .header p { margin: 10px 0 0; opacity: 0.9; }
    .summary { background: #f9f9f9; padding: 15px 20px; border-left: 4px solid ${statusColor}; margin: 20px; }
    .summary p { margin: 5px 0; font-size: 14px; }
    .priority { display: inline-block; padding: 4px 8px; border-radius: 4px; font-size: 12px; font-weight: bold; margin-right: 8px; }
    .priority.critical { background: #ffebee; color: #c62828; }
    .priority.high { background: #fff3e0; color: #e65100; }
    .priority.medium { background: #fffde7; color: #f57f17; }
    .module { margin: 20px; padding: 15px; border: 1px solid #ddd; border-radius: 4px; }
    .module-header { font-weight: bold; font-size: 16px; margin-bottom: 10px; color: #333; }
    .bug { padding: 8px 0; border-bottom: 1px solid #eee; }
    .bug:last-child { border-bottom: none; }
    .bug-id { font-family: monospace; color: #1976D2; font-weight: bold; }
    .footer { text-align: center; padding: 20px; background: #f9f9f9; }
    .button { display: inline-block; padding: 12px 24px; background: #1976D2; color: white; text-decoration: none; border-radius: 4px; margin-top: 10px; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>${statusIcon} QA Blocker Status</h1>
      <p>${now}</p>
    </div>

    <div class="summary">
      <p style="font-weight: bold; font-size: 16px; margin-bottom: 10px;">${statusMessage}</p>
`;

  if (totalBlocker > 0) {
    html += '<p>';
    if (totalCritical > 0) html += '<span class="priority critical">🔴 ' + totalCritical + ' Critical</span>';
    if (totalHigh > 0) html += '<span class="priority high">🟠 ' + totalHigh + ' High</span>';
    if (totalMedium > 0) html += '<span class="priority medium">🟡 ' + totalMedium + ' Medium</span>';
    html += '</p>';
  }

  html += '</div>';

  // Modules with blockers
  const modulesWithBlockers = blockerData.filter(m => m.totalBlocker > 0).sort((a, b) => b.totalBlocker - a.totalBlocker);

  modulesWithBlockers.forEach(mod => {
    html += '<div class="module">';
    html += '<div class="module-header">' + mod.project + ' &gt; ' + mod.module;
    if (mod.submodule) html += ' &gt; ' + mod.submodule;
    html += ' (' + mod.totalBlocker + ' blocker' + (mod.totalBlocker > 1 ? 's' : '') + ')</div>';
    html += '<p style="font-size: 12px; color: #666;">PIC: ' + (mod.team || '—') + '</p>';

    const topBugs = mod.bugs.slice(0, 5);
    topBugs.forEach(bug => {
      const prioClass = bug.priority === 'Critical' ? 'critical' : bug.priority === 'High' ? 'high' : 'medium';
      const prioIcon = bug.priority === 'Critical' ? '🔴' : bug.priority === 'High' ? '🟠' : '🟡';
      html += '<div class="bug">';
      html += '<span class="priority ' + prioClass + '">' + prioIcon + ' ' + bug.priority + '</span> ';

      // Make bug ID clickable if Jira link exists
      if (bug.link && bug.link.trim() && !bug.link.includes('http://example.com')) {
        html += '<a href="' + bug.link + '" style="font-family: monospace; color: #1976D2; font-weight: bold; text-decoration: none;" target="_blank">' + bug.bugId + ' 🔗</a> — ';
      } else {
        html += '<span class="bug-id">' + bug.bugId + '</span> — ';
      }

      html += bug.title.substring(0, 80);
      if (bug.title.length > 80) html += '...';
      html += '</div>';
    });

    if (mod.bugs.length > 5) {
      html += '<p style="font-size: 12px; color: #999; margin-top: 8px;">... dan ' + (mod.bugs.length - 5) + ' blocker lainnya</p>';
    }

    html += '</div>';
  });

  html += `
    <div class="footer">
      <a href="${dashboardUrl}" class="button">📊 Open Dashboard</a>
      <p style="font-size: 12px; color: #999; margin-top: 20px;">QA Test Management System - Automated Blocker Report</p>
    </div>
  </div>
</body>
</html>
`;

  return html;
}

/**
 * Build Google Chat card message for SINGLE MODULE
 * @param {Object} moduleData - Single module blocker data (or null for all clear)
 * @param {Spreadsheet} ss - Dashboard spreadsheet
 */
function _buildChatMessageForModule_(moduleData, ss) {
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm');
  const overviewSheet = ss.getSheetByName('Overview');
  const overviewGid = overviewSheet ? overviewSheet.getSheetId() : 0;
  const dashboardUrl = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/edit#gid=' + overviewGid;

  const widgets = [];

  // ALL CLEAR case (no blockers in any module)
  if (!moduleData) {
    widgets.push({
      decoratedText: {
        topLabel: '✅ ALL CLEAR',
        text: '<b>Tidak ada blocker!</b> 🎉 Semua modul dalam kondisi baik. Great job team!',
        bottomLabel: now
      }
    });

    widgets.push({ divider: {} });
    widgets.push({
      buttonList: {
        buttons: [{
          text: '📊 Open Dashboard',
          onClick: { openLink: { url: dashboardUrl } }
        }]
      }
    });

    return {
      cardsV2: [{
        cardId: 'blocker-all-clear',
        card: {
          header: {
            title: 'QA Blocker Status',
            subtitle: 'Daily Blocker Report',
            imageUrl: 'https://www.gstatic.com/images/branding/product/1x/keep_48dp.png',
            imageType: 'CIRCLE'
          },
          sections: [{ widgets: widgets }]
        }
      }]
    };
  }

  // Single module with blockers
  const totalBlocker = moduleData.totalBlocker;
  const totalCritical = moduleData.critical;
  const totalHigh = moduleData.high;
  const totalMedium = moduleData.medium;
  const totalProdBugs = moduleData.prodBugs || 0;

  // Determine severity
  let alertIcon, alertLabel, alertMessage;
  if (totalBlocker <= 3) {
    alertIcon = '⚠️';
    alertLabel = 'BLOCKER ALERT';
    alertMessage = `<b>${totalBlocker} blocker terdeteksi</b> — Kondisi masih terkendali, perlu perhatian.`;
  } else if (totalCritical >= 5) {
    alertIcon = '🔥';
    alertLabel = 'CRITICAL EMERGENCY';
    alertMessage = `<b>${totalBlocker} blocker (${totalCritical} CRITICAL)!</b> — ⚠️ URGENT: Prioritaskan resolve blocker critical!`;
  } else if (totalBlocker <= 10) {
    alertIcon = '🚨';
    alertLabel = 'BLOCKER ALERT';
    alertMessage = `<b>${totalBlocker} blocker terdeteksi!</b> — Perlu tindakan segera.`;
  } else {
    alertIcon = '🚨';
    alertLabel = 'BLOCKER ALERT';
    alertMessage = `<b>${totalBlocker} blocker terdeteksi!</b> — Situasi memburuk, prioritaskan penyelesaian.`;
  }

  // Header with module name
  widgets.push({
    decoratedText: {
      topLabel: `${alertIcon} ${alertLabel}`,
      text: `<b>${moduleData.project} > ${moduleData.module}</b>${moduleData.submodule ? ' > ' + moduleData.submodule : ''}`,
      bottomLabel: now
    }
  });

  // Alert message
  widgets.push({
    decoratedText: {
      topLabel: 'Status',
      text: alertMessage
    }
  });

  // Priority breakdown
  const summaryParts = [];
  if (totalCritical > 0) summaryParts.push(`🔴 ${totalCritical} Critical`);
  if (totalHigh > 0) summaryParts.push(`🟠 ${totalHigh} High`);
  if (totalMedium > 0) summaryParts.push(`🟡 ${totalMedium} Medium`);

  widgets.push({
    decoratedText: {
      topLabel: 'Breakdown by Priority',
      text: summaryParts.join('  |  ')
    }
  });

  // PROD BUGS alert (separate widget for visibility)
  if (totalProdBugs > 0) {
    widgets.push({
      decoratedText: {
        topLabel: '🚨 PRODUCTION BUGS',
        text: `<b>${totalProdBugs} bug${totalProdBugs > 1 ? 's' : ''} di Production environment!</b> — ⚠️ URGENT: Butuh immediate action!`,
        wrapText: true
      }
    });
  }

  // PIC info
  widgets.push({
    decoratedText: {
      topLabel: 'PIC QA',
      text: moduleData.team || '—',
      icon: { knownIcon: 'PERSON' }
    }
  });

  widgets.push({ divider: {} });

  // Top bugs (show max 5)
  const topBugs = moduleData.bugs.slice(0, 5);
  topBugs.forEach(bug => {
    const prioIcon = bug.priority === 'Critical' ? '🔴' : bug.priority === 'High' ? '🟠' : '🟡';
    const widget = {
      decoratedText: {
        topLabel: `${prioIcon} ${bug.bugId}`,
        text: bug.title.substring(0, 100)
      }
    };

    if (bug.link) {
      widget.decoratedText.button = {
        text: 'View',
        onClick: { openLink: { url: bug.link } }
      };
    }

    widgets.push(widget);
  });

  if (moduleData.bugs.length > 5) {
    widgets.push({
      textParagraph: {
        text: `<i>... dan ${moduleData.bugs.length - 5} blocker lainnya</i>`
      }
    });
  }

  // Dashboard & Bug Report links
  widgets.push({ divider: {} });

  const buttons = [{
    text: '📊 Open Dashboard',
    onClick: { openLink: { url: dashboardUrl } }
  }];

  // Add Bug Report link if there are bugs and URL is available
  if (moduleData.totalBlocker > 0 && moduleData.bugReportUrl) {
    buttons.push({
      text: '🐛 View All Bugs',
      onClick: { openLink: { url: moduleData.bugReportUrl } }
    });
  }

  widgets.push({
    buttonList: { buttons: buttons }
  });

  return {
    cardsV2: [{
      cardId: 'blocker-module-' + moduleData.module.replace(/\s+/g, '-').toLowerCase(),
      card: {
        header: {
          title: moduleData.project + ' > ' + moduleData.module,
          subtitle: totalBlocker + ' Blocker' + (totalBlocker > 1 ? 's' : '') + ' Detected',
          imageUrl: 'https://www.gstatic.com/images/branding/product/1x/keep_48dp.png',
          imageType: 'CIRCLE'
        },
        sections: [{ widgets: widgets }]
      }
    }]
  };
}

/**
 * Build HTML email body for SINGLE MODULE
 * @param {Object} moduleData - Single module blocker data (or null for all clear)
 * @param {Spreadsheet} ss - Dashboard spreadsheet
 */
function _buildEmailBodyForModule_(moduleData, ss) {
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm');
  const overviewSheet = ss.getSheetByName('Overview');
  const overviewGid = overviewSheet ? overviewSheet.getSheetId() : 0;
  const dashboardUrl = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/edit#gid=' + overviewGid;

  // ALL CLEAR case
  if (!moduleData) {
    const html = `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: Arial, sans-serif; margin: 0; padding: 20px; background: #f5f5f5; }
    .container { max-width: 600px; margin: 0 auto; background: white; border-radius: 8px; overflow: hidden; }
    .header { background: #4CAF50; color: white; padding: 20px; text-align: center; }
    .header h1 { margin: 0; font-size: 24px; }
    .header p { margin: 10px 0 0; opacity: 0.9; }
    .summary { background: #f9f9f9; padding: 20px; text-align: center; }
    .summary p { margin: 5px 0; font-size: 16px; }
    .footer { text-align: center; padding: 20px; background: #f9f9f9; }
    .button { display: inline-block; padding: 12px 24px; background: #1976D2; color: white; text-decoration: none; border-radius: 4px; margin-top: 10px; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>✅ QA Blocker Status</h1>
      <p>${now}</p>
    </div>
    <div class="summary">
      <p style="font-weight: bold; font-size: 18px; color: #4CAF50;">Tidak ada blocker! 🎉</p>
      <p>Semua modul dalam kondisi baik. Great job team!</p>
    </div>
    <div class="footer">
      <a href="${dashboardUrl}" class="button">📊 Open Dashboard</a>
      <p style="font-size: 12px; color: #999; margin-top: 20px;">QA Test Management System - Automated Blocker Report</p>
    </div>
  </div>
</body>
</html>`;
    return html;
  }

  // Single module with blockers
  const totalBlocker = moduleData.totalBlocker;
  const totalCritical = moduleData.critical;
  const totalHigh = moduleData.high;
  const totalMedium = moduleData.medium;
  const totalProdBugs = moduleData.prodBugs || 0;

  let statusColor, statusIcon, statusMessage;
  if (totalBlocker <= 3) {
    statusColor = '#FF9800';
    statusIcon = '⚠️';
    statusMessage = totalBlocker + ' blocker terdeteksi — Kondisi masih terkendali, perlu perhatian.';
  } else if (totalCritical >= 5) {
    statusColor = '#D32F2F';
    statusIcon = '🔥';
    statusMessage = totalBlocker + ' blocker (' + totalCritical + ' CRITICAL)! — URGENT: Prioritaskan resolve blocker critical!';
  } else {
    statusColor = '#F44336';
    statusIcon = '🚨';
    statusMessage = totalBlocker + ' blocker terdeteksi! — Perlu tindakan segera.';
  }

  // Prod bugs warning (prepend to status message if exists)
  let prodBugsWarning = '';
  if (totalProdBugs > 0) {
    prodBugsWarning = `
    <div style="background: #FFEBEE; border-left: 4px solid #D32F2F; padding: 15px; margin-bottom: 15px; border-radius: 4px;">
      <p style="margin: 0; font-weight: bold; color: #D32F2F; font-size: 16px;">🚨 PRODUCTION BUGS ALERT</p>
      <p style="margin: 5px 0 0; color: #C62828;">${totalProdBugs} bug${totalProdBugs > 1 ? 's' : ''} di Production environment! — ⚠️ URGENT: Butuh immediate action!</p>
    </div>`;
  }

  let html = `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: Arial, sans-serif; margin: 0; padding: 20px; background: #f5f5f5; }
    .container { max-width: 600px; margin: 0 auto; background: white; border-radius: 8px; overflow: hidden; }
    .header { background: ${statusColor}; color: white; padding: 20px; text-align: center; }
    .header h1 { margin: 0; font-size: 24px; }
    .header p { margin: 10px 0 0; opacity: 0.9; font-size: 14px; }
    .module-header { background: #f9f9f9; padding: 15px 20px; border-left: 4px solid ${statusColor}; margin: 20px; }
    .module-header h2 { margin: 0 0 5px; font-size: 18px; color: #333; }
    .module-header p { margin: 5px 0; font-size: 14px; color: #666; }
    .summary { background: #f9f9f9; padding: 15px 20px; border-left: 4px solid ${statusColor}; margin: 20px; }
    .summary p { margin: 5px 0; font-size: 14px; }
    .priority { display: inline-block; padding: 4px 8px; border-radius: 4px; font-size: 12px; font-weight: bold; margin-right: 8px; }
    .priority.critical { background: #ffebee; color: #c62828; }
    .priority.high { background: #fff3e0; color: #e65100; }
    .priority.medium { background: #fffde7; color: #f57f17; }
    .bugs { margin: 20px; }
    .bug { padding: 12px; border: 1px solid #ddd; border-radius: 4px; margin-bottom: 10px; background: white; }
    .bug-id { font-family: monospace; color: #1976D2; font-weight: bold; font-size: 14px; }
    .bug-title { margin: 5px 0 0; color: #333; }
    .footer { text-align: center; padding: 20px; background: #f9f9f9; }
    .button { display: inline-block; padding: 12px 24px; background: #1976D2; color: white; text-decoration: none; border-radius: 4px; margin-top: 10px; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>${statusIcon} QA Blocker Alert</h1>
      <p>${now}</p>
    </div>

    <div class="module-header">
      <h2>${moduleData.project} &gt; ${moduleData.module}${moduleData.submodule ? ' &gt; ' + moduleData.submodule : ''}</h2>
      <p><strong>PIC QA:</strong> ${moduleData.team || '—'}</p>
    </div>

    ${prodBugsWarning}

    <div class="summary">
      <p style="font-weight: bold; font-size: 16px; margin-bottom: 10px;">${statusMessage}</p>
      <p>`;

  if (totalCritical > 0) html += '<span class="priority critical">🔴 ' + totalCritical + ' Critical</span>';
  if (totalHigh > 0) html += '<span class="priority high">🟠 ' + totalHigh + ' High</span>';
  if (totalMedium > 0) html += '<span class="priority medium">🟡 ' + totalMedium + ' Medium</span>';

  html += '</p></div>';

  // List all bugs (max 10 for email)
  html += '<div class="bugs">';
  const topBugs = moduleData.bugs.slice(0, 10);
  topBugs.forEach(bug => {
    const prioClass = bug.priority === 'Critical' ? 'critical' : bug.priority === 'High' ? 'high' : 'medium';
    const prioIcon = bug.priority === 'Critical' ? '🔴' : bug.priority === 'High' ? '🟠' : '🟡';

    html += '<div class="bug">';
    html += '<span class="priority ' + prioClass + '">' + prioIcon + ' ' + bug.priority + '</span> ';

    // Make bug ID clickable if Jira link exists
    if (bug.link && bug.link.trim() && !bug.link.includes('http://example.com')) {
      html += '<a href="' + bug.link + '" class="bug-id" target="_blank" style="text-decoration: none; color: #1976D2;">' + bug.bugId + ' 🔗</a>';
    } else {
      html += '<span class="bug-id">' + bug.bugId + '</span>';
    }

    html += '<div class="bug-title">' + bug.title + '</div>';
    html += '</div>';
  });

  if (moduleData.bugs.length > 10) {
    html += '<p style="text-align: center; color: #999; font-size: 12px; margin-top: 10px;">... dan ' + (moduleData.bugs.length - 10) + ' blocker lainnya</p>';
  }

  html += '</div>';

  html += `
    <div class="footer">
      <a href="${dashboardUrl}" class="button">📊 Open Dashboard</a>`;

  // Add Bug Report link if available
  if (moduleData.bugReportUrl) {
    html += `
      <a href="${moduleData.bugReportUrl}" class="button" style="margin-left: 10px; background: #F44336;">🐛 View All Bugs in QATCM</a>`;
  }

  html += `
      <p style="font-size: 12px; color: #999; margin-top: 20px;">QA Test Management System - Automated Blocker Report</p>
    </div>
  </div>
</body>
</html>`;

  return html;
}

function _getBlockerData_(mods) {
  const blockerData = [];

  mods.forEach(mod => {
    try {
      const modSs = SpreadsheetApp.openById(mod.id);
      const bugSh = modSs.getSheetByName('BugReport');

      if (!bugSh) {
        blockerData.push({
          name: mod.name,
          project: mod.project || '',
          module: mod.module || '',
          submodule: mod.submodule || '',
          team: mod.team || '',
          totalBlocker: 0,
          critical: 0,
          high: 0,
          medium: 0,
          prodBugs: 0,
          bugs: []
        });
        return;
      }

      const last = bugSh.getLastRow();
      if (last < BUG_START_) {
        blockerData.push({
          name: mod.name,
          project: mod.project || '',
          module: mod.module || '',
          submodule: mod.submodule || '',
          team: mod.team || '',
          totalBlocker: 0,
          critical: 0,
          high: 0,
          medium: 0,
          prodBugs: 0,
          bugs: []
        });
        return;
      }

      const rows = bugSh.getRange(BUG_START_, 1, last - BUG_START_ + 1, BUG_COLS_).getValues();
      const BLOCKER_STATUS = ['open', 'in progress', 'reopen', 'fixed', 'verified', 'in progress vapt', 'done vapt'];
      const bugs = [];
      let critical = 0, high = 0, medium = 0, prodBugs = 0;

      rows.forEach(r => {
        const status = String(r[BC_.STATUS - 1]).trim().toLowerCase();
        const priority = String(r[BC_.PRIORITY - 1]).trim();
        const title = String(r[BC_.TITLE - 1]).trim();
        const bugId = String(r[BC_.BUG_ID - 1]).trim();
        const link = String(r[BC_.LINK - 1]).trim();
        const environment = String(r[BC_.ENVIRONMENT - 1]).trim();

        // Count PROD BUGS (bugs in Production environment, not closed)
        if (BLOCKER_STATUS.includes(status) && environment === 'Production') {
          prodBugs++;
        }

        if (!BLOCKER_STATUS.includes(status)) return;
        if (!['Critical', 'High', 'Medium'].includes(priority)) return;
        if (!title && !bugId) return;

        if (priority === 'Critical') critical++;
        else if (priority === 'High') high++;
        else if (priority === 'Medium') medium++;

        bugs.push({
          bugId: bugId || '—',
          priority,
          status: String(r[BC_.STATUS - 1]).trim(),
          title: title || '(no title)',
          link,
          environment
        });
      });

      // Build Bug Report sheet URL
      const bugSheetGid = bugSh.getSheetId();
      const bugReportUrl = 'https://docs.google.com/spreadsheets/d/' + mod.id + '/edit#gid=' + bugSheetGid;

      blockerData.push({
        name: mod.name,
        project: mod.project || '',
        module: mod.module || '',
        submodule: mod.submodule || '',
        team: mod.team || '',
        bugReportUrl: bugReportUrl,
        totalBlocker: critical + high + medium,
        critical,
        high,
        medium,
        prodBugs,
        bugs
      });

    } catch (e) {
      Logger.log('Error reading blocker from ' + mod.name + ': ' + e.message);
      blockerData.push({
        name: mod.name,
        project: mod.project || '',
        module: mod.module || '',
        submodule: mod.submodule || '',
        team: mod.team || '',
        totalBlocker: 0,
        critical: 0,
        high: 0,
        medium: 0,
        prodBugs: 0,
        bugs: [],
        error: e.message
      });
    }
  });

  return blockerData;
}

function _buildChatMessage_(blockerData, ss) {
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm');
  const totalCritical = blockerData.reduce((sum, m) => sum + m.critical, 0);
  const totalHigh = blockerData.reduce((sum, m) => sum + m.high, 0);
  const totalMedium = blockerData.reduce((sum, m) => sum + m.medium, 0);
  const totalBlocker = totalCritical + totalHigh + totalMedium;

  // Sort by totalBlocker descending
  blockerData.sort((a, b) => b.totalBlocker - a.totalBlocker);

  const widgets = [];

  // Determine severity level and message variation
  let alertIcon, alertLabel, alertMessage, headerColor;

  if (totalBlocker === 0) {
    // No blockers - Celebration!
    alertIcon = '✅';
    alertLabel = 'ALL CLEAR';
    alertMessage = '<b>Tidak ada blocker!</b> 🎉 Semua modul dalam kondisi baik. Great job team!';
    headerColor = '#4CAF50'; // Green
  } else if (totalBlocker <= 3) {
    // Few blockers - Mild alert
    alertIcon = '⚠️';
    alertLabel = 'BLOCKER ALERT - Low Priority';
    alertMessage = `<b>${totalBlocker} blocker terdeteksi</b> — Kondisi masih terkendali, perlu perhatian.`;
    headerColor = '#FF9800'; // Orange
  } else if (totalBlocker <= 10) {
    // Moderate blockers
    alertIcon = '🚨';
    alertLabel = 'BLOCKER ALERT - Moderate';
    alertMessage = `<b>${totalBlocker} blocker terdeteksi!</b> — Perlu tindakan segera untuk menghindari bottleneck.`;
    headerColor = '#FF5722'; // Deep Orange
  } else if (totalCritical >= 5) {
    // Many critical - Emergency!
    alertIcon = '🔥';
    alertLabel = 'CRITICAL EMERGENCY';
    alertMessage = `<b>${totalBlocker} blocker (${totalCritical} CRITICAL)!</b> — ⚠️ URGENT: Tim harus fokus resolve blocker critical!`;
    headerColor = '#D32F2F'; // Red
  } else {
    // Many blockers - High alert
    alertIcon = '🚨';
    alertLabel = 'BLOCKER ALERT - High Priority';
    alertMessage = `<b>${totalBlocker} blocker terdeteksi!</b> — Situasi memburuk, prioritaskan penyelesaian blocker.`;
    headerColor = '#F44336'; // Red
  }

  // Header section with varied message
  widgets.push({
    decoratedText: {
      topLabel: `${alertIcon} ${alertLabel}`,
      text: alertMessage,
      bottomLabel: now
    }
  });

  // Summary section - only show if there are blockers
  if (totalBlocker > 0) {
    const summaryParts = [];
    if (totalCritical > 0) summaryParts.push(`🔴 ${totalCritical} Critical`);
    if (totalHigh > 0) summaryParts.push(`🟠 ${totalHigh} High`);
    if (totalMedium > 0) summaryParts.push(`🟡 ${totalMedium} Medium`);

    widgets.push({
      decoratedText: {
        topLabel: 'Breakdown by Priority',
        text: summaryParts.join('  |  ')
      }
    });
  }

  // Modules with blockers
  const modulesWithBlockers = blockerData.filter(m => m.totalBlocker > 0);

  if (modulesWithBlockers.length > 0) {
    widgets.push({ divider: {} });

    modulesWithBlockers.forEach((mod, idx) => {
      if (idx > 0) widgets.push({ divider: {} });

      widgets.push({
        decoratedText: {
          topLabel: `${mod.project} > ${mod.module}${mod.submodule ? ' > ' + mod.submodule : ''}`,
          text: `<b>${mod.totalBlocker} blocker${mod.totalBlocker > 1 ? 's' : ''}</b>`,
          bottomLabel: `PIC: ${mod.team || '—'}`,
          icon: { knownIcon: 'STAR' }
        }
      });

      const summary = [];
      if (mod.critical > 0) summary.push(`🔴 ${mod.critical} Critical`);
      if (mod.high > 0) summary.push(`🟠 ${mod.high} High`);
      if (mod.medium > 0) summary.push(`🟡 ${mod.medium} Medium`);

      widgets.push({
        textParagraph: {
          text: summary.join('  |  ')
        }
      });

      // Top 3 bugs
      const topBugs = mod.bugs.slice(0, 3);
      topBugs.forEach(bug => {
        const prioIcon = bug.priority === 'Critical' ? '🔴' : bug.priority === 'High' ? '🟠' : '🟡';
        const widget = {
          decoratedText: {
            topLabel: `${prioIcon} ${bug.bugId}`,
            text: bug.title.substring(0, 100)
          }
        };

        // Add button only if link exists
        if (bug.link) {
          widget.decoratedText.button = {
            text: 'View',
            onClick: { openLink: { url: bug.link } }
          };
        }

        widgets.push(widget);
      });

      if (mod.bugs.length > 3) {
        widgets.push({
          textParagraph: {
            text: `<i>... dan ${mod.bugs.length - 3} blocker lainnya</i>`
          }
        });
      }
    });
  }

  // Dashboard link - open Overview tab
  if (ss) {
    widgets.push({ divider: {} });

    // Get Overview sheet ID
    const overviewSheet = ss.getSheetByName('Overview');
    const overviewGid = overviewSheet ? overviewSheet.getSheetId() : 0;
    const dashboardUrl = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/edit#gid=' + overviewGid;

    widgets.push({
      buttonList: {
        buttons: [{
          text: '📊 Open Dashboard',
          onClick: {
            openLink: {
              url: dashboardUrl
            }
          }
        }]
      }
    });
  }

  return {
    cardsV2: [{
      cardId: 'blocker-notification',
      card: {
        header: {
          title: 'QA Blocker Status',
          subtitle: 'Daily Blocker Report',
          imageUrl: 'https://www.gstatic.com/images/branding/product/1x/keep_48dp.png',
          imageType: 'CIRCLE'
        },
        sections: [{
          widgets: widgets
        }]
      }
    }]
  };
}


// safeAlert_ — fallback jika tidak ada di file lain dalam project
// (Jika QA_Portfolio_Dashboard.js ada di project yang sama, fungsi ini tidak akan bentrok)
function safeAlert_(msg) {
  try { SpreadsheetApp.getUi().alert(msg); } catch(e) { Logger.log(msg); }
}
