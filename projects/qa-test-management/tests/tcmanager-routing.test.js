const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const source = fs.readFileSync(path.join(__dirname, '../src/TCManager.js'), 'utf8');
const context = {
  CacheService: { getScriptCache: () => ({ get: () => 'true', put: () => {} }) },
  SpreadsheetApp: { getUi: () => ({ createMenu: () => ({ addItem: () => {}, addSeparator: () => {}, addToUi: () => {} }) }) },
};
vm.createContext(context);
vm.runInContext(source, context);

function sheet(name) {
  return { getName: () => name };
}

function spreadsheet(activeName) {
  const sheets = {
    TC_Master: sheet('TC_Master'),
    TC_Execution: sheet('TC_Execution'),
    API_Master: sheet('API_Master'),
    API_Execution: sheet('API_Execution'),
  };
  return {
    getActiveSheet: () => sheets[activeName],
    getSheetByName: (name) => sheets[name],
  };
}

let pair = context.getActiveSheetPair(spreadsheet('API_Master'));
assert.strictEqual(pair.master.getName(), 'API_Master');
assert.strictEqual(pair.execution.getName(), 'API_Execution');

pair = context.getActiveSheetPair(spreadsheet('TC_Master'));
assert.strictEqual(pair.master.getName(), 'TC_Master');
assert.strictEqual(pair.execution.getName(), 'TC_Execution');

assert.strictEqual(context.getActiveSheetPair(spreadsheet('Summary')), null);
