function readVAPTDropdowns() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('VAPT - Detail Finding');

    if (!sheet) {
      Logger.log('❌ VAPT - Detail Finding not found');
      return;
    }

    Logger.log('=== VAPT DETAIL FINDING - DROPDOWN VALUES ===\n');

    const headers = sheet.getRange(2, 1, 1, 31).getValues()[0];
    const cols = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE'];

    for (let i = 1; i <= 31; i++) {
      const validation = sheet.getRange(3, i).getDataValidation();
      if (validation) {
        Logger.log('Column ' + cols[i-1] + ': ' + headers[i-1]);
        try {
          const criteria = validation.getCriteriaType();
          if (criteria === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
            const values = validation.getCriteriaValues()[0];
            Logger.log('  Values: ' + JSON.stringify(values));
          } else if (criteria === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
            const range = validation.getCriteriaValues()[0];
            Logger.log('  Range: ' + range.getA1Notation());
          }
        } catch(e) {
          Logger.log('  Error: ' + e.message);
        }
        Logger.log('');
      }
    }

    Logger.log('=== DONE ===');
  }
