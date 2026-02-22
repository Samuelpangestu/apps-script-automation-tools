// =============================================
// STEP TEST - jalankan testStep1, 2, 3 satu per satu
// Lihat View > Logs setelah setiap step
// =============================================

function testStep1_SimpleSheet() {
  Logger.log('Step 1: Create simple sheet');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.insertSheet('TEST_SIMPLE');
  if (!ws) { Logger.log('ERROR: ws is undefined'); return; }
  ws.getRange(1,1).setValue('Hello World');
  ws.getRange(1,2).setValue(new Date());
  Logger.log('Step 1 OK - sheet created, ws = ' + ws.getName());
}

function testStep2_DeleteAndRecreate() {
  Logger.log('Step 2: Delete + recreate sheet');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ex = ss.getSheetByName('TEST_SIMPLE');
  if (ex) { ss.deleteSheet(ex); SpreadsheetApp.flush(); Logger.log('Deleted existing'); }
  const ws = ss.insertSheet('TEST_SIMPLE');
  if (!ws) { Logger.log('ERROR: ws is undefined after delete+insert'); return; }
  ws.getRange(1,1).setValue('Recreated');
  Logger.log('Step 2 OK - ws = ' + ws.getName());
}

function testStep3_ConditionalFormat() {
  Logger.log('Step 3: Apply conditional formatting');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName('TEST_SIMPLE');
  if (!ws) { Logger.log('ERROR: TEST_SIMPLE not found - run Step 1 or 2 first'); return; }
  
  const range = ws.getRange(1, 1, 10, 5);
  Logger.log('Range OK: ' + range.getA1Notation());
  
  const rules = ws.getConditionalFormatRules();
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('PASSED')
      .setBackground('#C8E6C9')
      .setFontColor('#1B5E20')
      .setRanges([range])
      .build()
  );
  ws.setConditionalFormatRules(rules);
  Logger.log('Step 3 OK - conditional format applied');
}

function testStep4_CleanUp() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName('TEST_SIMPLE');
  if (ws) { ss.deleteSheet(ws); Logger.log('Cleaned up TEST_SIMPLE'); }
}
