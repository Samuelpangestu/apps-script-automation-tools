function createQASheet() {
  Logger.log('START');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Clean up
  ['TC_Master'].forEach(name => {
    const s = ss.getSheetByName(name);
    if (s) { try { ss.deleteSheet(s); } catch(e) {} }
  });
  SpreadsheetApp.flush();
  
  // Create sheet
  const ws = ss.insertSheet('TC_Master');
  Logger.log('ws = ' + (ws ? ws.getName() : 'UNDEFINED'));
  if (!ws) { Logger.log('ERROR: ws undefined'); return; }
  
  // Basic write
  ws.getRange(1,1).setValue('TC_ID');
  ws.getRange(1,2).setValue('Scenario');
  ws.getRange(1,3).setValue('Priority');
  
  // Apply conditional format
  try {
    const range = ws.getRange(2, 3, 100, 1);
    const rules = ws.getConditionalFormatRules();
    ['Critical','High','Medium','Low','Lowest'].forEach((v,i) => {
      const bgs = ['#FFCDD2','#FFE0B2','#FFF9C4','#F0F4F8','#FAFAFA'];
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo(v).setBackground(bgs[i])
          .setRanges([range]).build()
      );
    });
    ws.setConditionalFormatRules(rules);
    Logger.log('CF applied OK');
  } catch(e) {
    Logger.log('CF error: ' + e.message);
  }
  
  Logger.log('DONE');
}
