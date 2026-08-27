const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const source = fs.readFileSync(path.join(__dirname, '../src/TCManager.js'), 'utf8');

function run(masterName, executionName, scenarioColumn, scenario) {
  const formatted = [];
  const values = { '3:3': 'TC.001', [`3:${scenarioColumn}`]: scenario };
  const range = (sheetName, row, column, numColumns) => ({
    getValue: () => values[`${row}:${column}`],
    setValue: value => { values[`${row}:${column}`] = value; },
    setBackground: color => {
      formatted.push({ sheetName, row, column, numColumns, color });
      return { setFontColor: () => {} };
    },
  });
  const master = {
    getName: () => masterName,
    getActiveRange: () => ({ getRow: () => 3, getNumRows: () => 1 }),
    getRange: (row, column, _numRows, numColumns) => range(masterName, row, column, numColumns),
    getLastColumn: () => scenarioColumn,
  };
  const execution = {
    getName: () => executionName,
    getRange: (row, column, _numRows, numColumns) => range(executionName, row, column, numColumns),
    getLastColumn: () => 26,
  };
  const ui = {
    Button: { YES: 'YES' },
    ButtonSet: { YES_NO: 'YES_NO', OK: 'OK' },
    alert: (...args) => args[2] === 'YES_NO' ? 'YES' : undefined,
    createMenu: () => ({ addItem() { return this; }, addSeparator() { return this; }, addToUi() {} }),
  };
  const spreadsheet = {
    getActiveSheet: () => master,
    getSheetByName: name => name === executionName ? execution : null,
  };
  const context = {
    CacheService: { getScriptCache: () => ({ get: () => 'true', put: () => {} }) },
    SpreadsheetApp: {
      getUi: () => ui,
      getActiveSpreadsheet: () => spreadsheet,
    },
  };

  vm.createContext(context);
  vm.runInContext(source, context);
  context.markAsDeprecated();
  return formatted;
}

for (const [master, execution, scenarioColumn] of [
  ['TC_Master', 'TC_Execution', 11],
  ['API_Master', 'API_Execution', 13],
]) {
  const formatted = run(master, execution, scenarioColumn, '[DEPRECATED] Existing scenario');
  assert(formatted.some(item => item.sheetName === master && item.row === 3 && item.color === '#F5F5F5'));
  assert(formatted.some(item => item.sheetName === execution && item.row === 9 && item.color === '#F5F5F5'));
}
