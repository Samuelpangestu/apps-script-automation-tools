/**
 * Review360.js — 360 Review Form Management
 * ═══════════════════════════════════════════════════════════════════════
 * Enhanced version with Config integration and auto-aggregation
 * ═══════════════════════════════════════════════════════════════════════
 */

// Roles that will be included as reviewees (excluding Team Lead)
const REVIEWEE_ROLES = [
  'QA Lead (Project Dedicated)',
  'PIC Project (QE + Koordinator)',
  'Quality Engineer (QE)'
];

// Reviewer types for the form
const REVIEWER_TYPES = [
  'QA Team Lead',
  'PM / Tribe Lead',
  'QE Peer',
  'Self Assessment',
  'Developer',
  'Product Owner (PO)',
  'Design'
];

// Rating criteria (aligned with KPI 360 Review Score)
const RATING_CRITERIA = [
  {
    title: 'Technical Competency',
    helpText: 'Pemahaman standar QA, test strategy, tools, dan metodologi. ' +
              'Contoh: kemampuan membuat test plan, menggunakan Jira, Playwright, dll.'
  },
  {
    title: 'Delivery & Coordination',
    helpText: 'Ketepatan delivery, koordinasi dengan developer, PM, dan stakeholder. ' +
              'Contoh: konsistensi memenuhi sprint commitment, komunikasi blocker tepat waktu.'
  },
  {
    title: 'Communication',
    helpText: 'Kejelasan laporan, eskalasi masalah, dan update kepada stakeholder. ' +
              'Contoh: kualitas bug report, kejelasan status update, efektivitas meeting.'
  },
  {
    title: 'Team Leadership',
    helpText: 'Membimbing rekan tim, problem solving, dan mendukung pertumbuhan tim. ' +
              'Contoh: membantu QE junior, proaktif berbagi knowledge, inisiatif saat ada blocker.'
  },
  {
    title: 'Quality Mindset',
    helpText: 'Inisiatif improvement, kepatuhan standar CoE, dan quality awareness. ' +
              'Contoh: mengusulkan perbaikan workflow, konsisten menggunakan template CoE.'
  }
];

/**
 * Get reviewees from Config tab (active members only)
 * @returns {Array} Array of reviewee names
 */
function getRevieweesFromConfig_() {
  const members = getTeamMembers();  // From TeamConfig.js
  const reviewees = [];

  members.forEach(member => {
    if (REVIEWEE_ROLES.indexOf(member.role) !== -1 && member.status === 'Aktif') {
      reviewees.push(member.name);
    }
  });

  if (reviewees.length === 0) {
    throw new Error(
      'Tidak ada reviewee ditemukan di Config.\n' +
      'Pastikan ada anggota dengan role: ' + REVIEWEE_ROLES.join(', ') +
      ' dan Status = Aktif'
    );
  }

  return reviewees;
}

/**
 * Create 360 Review Form
 * Run this once to create the form
 */
function create360ReviewForm() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reviewees = getRevieweesFromConfig_();

  // Check if form already exists
  const existingSettings = getFormSettings();  // From TeamConfig.js
  if (existingSettings.formId) {
    const ui = SpreadsheetApp.getUi();
    const res = ui.alert(
      'Form sudah pernah dibuat',
      'Form ID sudah ada di Config (' + existingSettings.formId + ').\n\n' +
      'Apakah Anda ingin membuat form BARU? (Form lama tidak akan dihapus)\n' +
      'Klik "Tidak" untuk membatalkan, atau jalankan update360RevieweeList() ' +
      'untuk update daftar nama saja.',
      ui.ButtonSet.YES_NO
    );
    if (res !== ui.Button.YES) return;
  }

  // Create new form
  const form = FormApp.create('360 Review — QA Department PERURI');

  form.setTitle('360 Review — QA Department PERURI');
  form.setDescription(
    'Form penilaian 360 Review untuk anggota tim QA.\n' +
    'Satu pengisian = satu reviewer untuk satu reviewee.\n\n' +
    'Skala penilaian:\n' +
    '1 = Perlu peningkatan signifikan\n' +
    '2 = Di bawah ekspektasi\n' +
    '3 = Memenuhi sebagian ekspektasi\n' +
    '4 = Memenuhi ekspektasi\n' +
    '5 = Melampaui ekspektasi'
  );
  form.setCollectEmail(false);
  form.setLimitOneResponsePerUser(false);
  form.setShowLinkToRespondAgain(true);

  // Section 1: Identitas
  form.addSectionHeaderItem()
    .setTitle('Identitas')
    .setHelpText('Isi informasi reviewer dan siapa yang Anda nilai.');

  form.addTextItem()
    .setTitle('Periode Review')
    .setHelpText('Contoh: Semester 1 2026 atau Q1 2026')
    .setRequired(true);

  form.addTextItem()
    .setTitle('Nama Reviewer')
    .setHelpText('Nama lengkap Anda')
    .setRequired(true);

  form.addListItem()
    .setTitle('Nama yang Direview')
    .setHelpText('Pilih nama anggota yang Anda nilai')
    .setChoiceValues(reviewees)
    .setRequired(true);

  form.addListItem()
    .setTitle('Tipe Reviewer')
    .setHelpText('Pilih peran Anda sebagai reviewer')
    .setChoiceValues(REVIEWER_TYPES)
    .setRequired(true);

  // Section 2: Penilaian
  form.addPageBreakItem()
    .setTitle('Penilaian')
    .setHelpText(
      '1 = Perlu peningkatan signifikan  |  2 = Di bawah ekspektasi  |  ' +
      '3 = Memenuhi sebagian  |  4 = Memenuhi ekspektasi  |  5 = Melampaui ekspektasi'
    );

  // Add rating criteria from constant
  RATING_CRITERIA.forEach(criteria => {
    form.addScaleItem()
      .setTitle(criteria.title)
      .setHelpText(criteria.helpText)
      .setBounds(1, 5)
      .setLabels('Perlu peningkatan', 'Melampaui ekspektasi')
      .setRequired(true);
  });

  form.addParagraphTextItem()
    .setTitle('Catatan Tambahan')
    .setHelpText('Opsional. Apresiasi atau saran pengembangan untuk orang yang Anda nilai.');

  form.setConfirmationMessage(
    'Terima kasih! Penilaian Anda telah berhasil dikirim. ' +
    'Data akan diproses setelah periode review selesai.'
  );

  // Link responses to this spreadsheet
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  // Save form settings to Config
  const shortUrl = form.shortenFormUrl(form.getPublishedUrl());
  saveFormSettings(form.getId(), shortUrl, form.getEditUrl());  // From TeamConfig.js

  Logger.log('Form ID  : ' + form.getId());
  Logger.log('Form URL : ' + shortUrl);
  Logger.log('Edit URL : ' + form.getEditUrl());
  Logger.log('Reviewees: ' + reviewees.length + ' orang');

  SpreadsheetApp.getUi().alert(
    'Form berhasil dibuat!\n\n' +
    'Reviewees: ' + reviewees.length + ' orang (dibaca dari Config)\n' +
    'Daftar: ' + reviewees.join(', ') + '\n\n' +
    'Link form untuk reviewer:\n' + shortUrl + '\n\n' +
    'Link dan Form ID sudah tersimpan otomatis di tab Config.'
  );
}

/**
 * Update reviewee list in existing form
 * Run this when team members change
 */
function update360RevieweeList() {
  const settings = getFormSettings();  // From TeamConfig.js

  if (!settings.formId) {
    SpreadsheetApp.getUi().alert(
      'Form ID tidak ditemukan di Config.\n' +
      'Jalankan create360ReviewForm() terlebih dahulu.'
    );
    return;
  }

  const reviewees = getRevieweesFromConfig_();
  const form = FormApp.openById(settings.formId);
  const items = form.getItems();
  let updated = false;

  for (let i = 0; i < items.length; i++) {
    if (items[i].getTitle() === 'Nama yang Direview') {
      items[i].asListItem().setChoiceValues(reviewees);
      updated = true;
      break;
    }
  }

  if (updated) {
    Logger.log('Reviewees diperbarui: ' + reviewees.length + ' orang');
    Logger.log(reviewees.join(', '));
    SpreadsheetApp.getUi().alert(
      'Daftar reviewee di form berhasil diperbarui!\n\n' +
      'Total: ' + reviewees.length + ' orang\n' +
      reviewees.join('\n')
    );
  } else {
    SpreadsheetApp.getUi().alert(
      'Gagal — item "Nama yang Direview" tidak ditemukan di form.\n' +
      'Pastikan Form ID di Config sudah benar.'
    );
  }
}

/**
 * Show 360 Review Form info
 */
function show360ReviewFormInfo() {
  const settings = getFormSettings();  // From TeamConfig.js

  if (!settings.formId) {
    SpreadsheetApp.getUi().alert('Belum ada form. Jalankan create360ReviewForm() terlebih dahulu.');
    return;
  }

  try {
    const form = FormApp.openById(settings.formId);
    const items = form.getItems();
    let reviewees = [];

    for (let i = 0; i < items.length; i++) {
      if (items[i].getTitle() === 'Nama yang Direview') {
        reviewees = items[i].asListItem().getChoices().map(c => c.getValue());
        break;
      }
    }

    SpreadsheetApp.getUi().alert(
      'Info Form 360 Review\n\n' +
      'Form ID  : ' + settings.formId + '\n' +
      'Form URL : ' + settings.formUrl + '\n\n' +
      'Reviewees saat ini (' + reviewees.length + ' orang):\n' +
      reviewees.join('\n')
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error membuka form: ' + e.message);
  }
}

/**
 * Calculate weighted 360 review score for a reviewee
 * Weighting: TL 35% + PM 25% + QE Peer 15% + Self 10% + Optional 15%
 *
 * @param {string} revieweeName - Name of person being reviewed
 * @param {string} period - Review period (e.g., "Semester 1 2026")
 * @returns {number} Weighted average score (1-5 scale)
 */
function calculate360Score(revieweeName, period) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const responsesSheet = ss.getSheetByName('Form Responses 1'); // Default name from Google Forms

  if (!responsesSheet) {
    Logger.log('⚠️ No form responses sheet found');
    return null;
  }

  const data = responsesSheet.getDataRange().getValues();
  if (data.length <= 1) {
    Logger.log('⚠️ No responses yet');
    return null;
  }

  // Find column indices (adjust based on actual form structure)
  const headers = data[0];
  let periodeCol = -1, revieweeCol = -1, reviewerTypeCol = -1;
  const ratingCols = [];

  headers.forEach((header, idx) => {
    const h = header.toString().toLowerCase();
    if (h.includes('periode')) periodeCol = idx;
    else if (h.includes('yang direview')) revieweeCol = idx;
    else if (h.includes('tipe reviewer')) reviewerTypeCol = idx;
    else if (RATING_CRITERIA.some(c => header.toString().includes(c.title))) {
      ratingCols.push(idx);
    }
  });

  if (periodeCol === -1 || revieweeCol === -1 || reviewerTypeCol === -1 || ratingCols.length === 0) {
    Logger.log('⚠️ Could not find required columns in form responses');
    return null;
  }

  // Collect scores by reviewer type
  const scoresByType = {
    'QA Team Lead': [],
    'PM / Tribe Lead': [],
    'QE Peer': [],
    'Self Assessment': [],
    'Other': []  // Developer, PO, Design
  };

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowPeriod = row[periodeCol].toString().trim();
    const rowReviewee = row[revieweeCol].toString().trim();
    const rowReviewerType = row[reviewerTypeCol].toString().trim();

    if (rowPeriod === period && rowReviewee === revieweeName) {
      // Calculate average rating for this response
      let sum = 0;
      let count = 0;
      ratingCols.forEach(col => {
        const val = parseFloat(row[col]);
        if (!isNaN(val)) {
          sum += val;
          count++;
        }
      });

      if (count > 0) {
        const avgScore = sum / count;

        if (rowReviewerType === 'QA Team Lead') {
          scoresByType['QA Team Lead'].push(avgScore);
        } else if (rowReviewerType === 'PM / Tribe Lead') {
          scoresByType['PM / Tribe Lead'].push(avgScore);
        } else if (rowReviewerType === 'QE Peer') {
          scoresByType['QE Peer'].push(avgScore);
        } else if (rowReviewerType === 'Self Assessment') {
          scoresByType['Self Assessment'].push(avgScore);
        } else {
          scoresByType['Other'].push(avgScore);
        }
      }
    }
  }

  // Calculate weighted average
  // TL 35% + PM 25% + QE Peer 15% + Self 10% + Optional 15%
  let weightedSum = 0;
  let totalWeight = 0;

  const weights = {
    'QA Team Lead': 0.35,
    'PM / Tribe Lead': 0.25,
    'QE Peer': 0.15,
    'Self Assessment': 0.10,
    'Other': 0.15
  };

  Object.keys(scoresByType).forEach(type => {
    const scores = scoresByType[type];
    if (scores.length > 0) {
      const avg = scores.reduce((sum, val) => sum + val, 0) / scores.length;
      weightedSum += avg * weights[type];
      totalWeight += weights[type];
    }
  });

  if (totalWeight === 0) {
    return null;
  }

  // Normalize to actual weights used
  const finalScore = weightedSum / totalWeight;
  return Math.round(finalScore * 100) / 100;  // Round to 2 decimal places
}
