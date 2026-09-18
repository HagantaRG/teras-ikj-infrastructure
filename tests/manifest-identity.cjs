const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert/strict');

function readSources(directory) {
  const sources = {};
  for (const entry of fs.readdirSync(directory, {withFileTypes: true})) {
    const filename = path.join(directory, entry.name);
    if (entry.isDirectory()) {
      Object.assign(sources, readSources(filename));
    } else if (entry.name.endsWith('.js')) {
      const source = fs.readFileSync(filename, 'utf8');
      new vm.Script(source, {filename});
      sources[filename] = source;
    }
  }
  return sources;
}

const candidate = readSources(path.join(__dirname, '..', 'src'));
console.log('PASS: all deployed JavaScript files parse');
const context = vm.createContext({
  Date, Map, Set,
  SpreadsheetApp: {DeveloperMetadataLocationType: {COLUMN: 'COLUMN'}, DeveloperMetadataVisibility: {DOCUMENT: 'DOCUMENT'}},
  FormApp: {ItemType: {TEXT: 'TEXT'}},
  Session: {getScriptTimeZone: () => 'Asia/Bangkok'},
  Utilities: {formatDate: date => date.toISOString().slice(0, 10)}
});
vm.runInContext(Object.values(candidate).join('\n'), context);
function manifest(rows) {
  return {
    getLastRow: () => rows.length,
    getLastColumn: () => rows[0].length,
    getRange: (row, column, count, width) => ({
      getDisplayValues: () => rows.slice(row - 1, row - 1 + count)
        .map(values => values.slice(column - 1, column - 1 + width).map(String))
    })
  };
}
const headers = ['ID Barang', 'Nama Barang', 'Unit', 'Level Stok Minim', 'Aktif'];
const manifestRows = [headers, ['TRS-0001', 'Kopi', 'kg', '2', 'TRUE']];
assert.equal(context.readManifestItems_(manifest(manifestRows))[0].displayName, 'Kopi/kg');
const reordered = manifestRows.map(row => [row[4], row[2], row[0], row[3], row[1]]);
assert.equal(context.readManifestItems_(manifest(reordered))[0].id, 'TRS-0001');
assert.throws(() => context.readManifestItems_(manifest([headers, manifestRows[1], manifestRows[1]])), /more than once/);
assert.throws(() => context.readManifestItems_(manifest([['ID Barang']])), /missing required headers/);
assert.equal(context.readManifestItems_(manifest([headers, ['', '', '', '', 'FALSE']])).length, 0);
console.log('PASS: manifest validation, reordered headers and empty checkbox rows');
const saved = {};
const properties = {
  getProperties: () => ({...saved}),
  setProperty: (key, value) => {saved[key] = value;},
  deleteProperty: key => {delete saved[key];}
};
const sheet = {getParent: () => ({getId: () => 'book'}), getSheetId: () => 1};
let count = 0;
const questions = [];
const form = {
  getId: () => 'form',
  getItems: () => questions.slice(),
  addTextItem: () => {
    const id = String(++count);
    const item = {
      getId: () => id, getType: () => 'TEXT',
      getIndex: () => questions.indexOf(item),
      asTextItem: () => item,
      setTitle: title => {item.title = title; return item;},
      setRequired: required => {item.required = required; return item;},
      setValidation: () => item
    };
    questions.push(item);
    return item;
  },
  deleteItem: item => questions.splice(questions.indexOf(item), 1),
  moveItem: (item, index) => {
    questions.splice(questions.indexOf(item), 1);
    questions.splice(index, 0, item);
  }
};
const columns = [{itemId: 'TRS-0001', columnNumber: 3, title: 'Kopi/kg', active: true}];
context.syncInventoryForm_(form, sheet, columns, properties, {});
const firstQuestion = questions[0];
assert.ok(firstQuestion);
columns[0].title = 'Kopi Bubuk/kg';
context.syncInventoryForm_(form, sheet, columns, properties, {});
assert.equal(questions.length, 1);
assert.equal(questions[0], firstQuestion);
assert.equal(firstQuestion.title, 'Kopi Bubuk/kg');
assert.equal(firstQuestion.required, false);
assert.ok(Object.keys(saved).some(key => key.endsWith(':TRS-0001')));
columns[0].active = false;
context.syncInventoryForm_(form, sheet, columns, properties, {});
assert.equal(questions.length, 0);
console.log('PASS: ID Barang question mapping, rename, optional fields and deactivation');
const inventory = {
  getLastRow: () => 2,
  getRange: (row, column) => ({
    getValues: () => column === 2 ? [[new Date('2026-09-18T00:00:00Z')]] : [[0]]
  })
};
const dashboardRows = context.buildCurrentDashboardRows_(inventory, columns);
assert.equal(dashboardRows[0][1], 'TRS-0001');
assert.equal(dashboardRows[0][3], 0);
console.log('PASS: dashboard IDs and zero stock, including inactive items');
const events = [];
context.createFormFromSheet = event => events.push(event);
for (const type of ['INSERT_ROW', 'REMOVE_ROW', 'INSERT_COLUMN', 'REMOVE_COLUMN']) {
  context.handleInventoryColumnChange({changeType: type});
}
assert.equal(events.length, 4);
context.handleInventoryHeaderEdit({range: {
  getSheet: () => ({getName: () => 'daftar-barang'}),
  getRow: () => 1, getColumn: () => 8
}});
assert.equal(events.length, 5);
console.log('PASS: row/column structural events and manifest edits outside A:E');

function isolatedContext(overrides) {
  const sandbox = vm.createContext({Date, Map, Set, ...overrides});
  vm.runInContext(Object.values(candidate).join('\n'), sandbox);
  return sandbox;
}

const writes = [];
const responseProperties = {
  INVENTORY_SPREADSHEET_ID: 'book',
  'INVENTORY_ITEM:book:1:form:TRS-0001': '101',
  'INVENTORY_ITEM:book:1:form:TRS-0002': '102'
};
const responseSheet = {
  getParent: () => ({getId: () => 'book'}),
  getSheetId: () => 1,
  getRange: (row, column) => ({setValue: value => writes.push({row, column, value})})
};
let released = 0;
let synced = 0;
const responses = isolatedContext({
  LockService: {getScriptLock: () => ({
    waitLock: () => {}, releaseLock: () => {released++;}
  })},
  PropertiesService: {getScriptProperties: () => ({
    getProperty: key => responseProperties[key],
    getProperties: () => ({...responseProperties})
  })},
  SpreadsheetApp: {openById: () => ({getSheetByName: () => responseSheet})}
});
responses.syncInventoryColumnsFromManifest_ = () => {};
responses.getInventoryColumns_ = () => [
  {itemId: 'TRS-0001', columnNumber: 5, active: true},
  {itemId: 'TRS-0002', columnNumber: 3, active: true}
];
responses.findOrCreateTodayRow_ = () => 7;
responses.syncDashboardData_ = () => {synced++;};
function submission(values) {
  return {
    source: {getId: () => 'form'},
    response: {getItemResponses: () => values.map(([id, answer]) => ({
      getItem: () => ({getId: () => id}),
      getResponse: () => answer
    }))}
  };
}
responses.recordInventoryFormResponse_(submission([['101', '0'], ['102', '']]));
assert.deepEqual(writes, [{row: 7, column: 5, value: 0}]);
assert.equal(synced, 1);
responses.recordInventoryFormResponse_(submission([['101', ''], ['102', '']]));
assert.equal(writes.length, 1);
assert.equal(synced, 1);
assert.throws(() => responses.recordInventoryFormResponse_(
  submission([['101', '3'], ['102', '-1']])
), /whole number/);
assert.equal(writes.length, 1);
assert.equal(released, 3);
console.log('PASS: submissions resolve by ID after column movement; blanks and invalid answers do not overwrite cells');

const resetEvents = [];
const resetProperties = {
  INVENTORY_FORM_ID: 'same-form',
  UNRELATED_SETTING: 'keep',
  'INVENTORY_ITEM:book:1:same-form:123': '101'
};
const resetSheet = {
  getParent: () => ({getId: () => 'book'}),
  getSheetId: () => 1,
  getMaxRows: () => 10,
  getMaxColumns: () => 6,
  getRange: (...args) => ({clearContent: () => resetEvents.push(['clear stock', ...args])}),
  deleteColumns: (...args) => resetEvents.push(['delete columns', ...args]),
  createDeveloperMetadataFinder: () => {
    const finder = {
      withKey: key => {
        finder.key = key;
        return finder;
      },
      find: () => [{remove: () => resetEvents.push(['remove metadata', finder.key])}]
    };
    return finder;
  }
};
const resetBook = {
  getSheetByName: name => name === 'stok-barang' ? resetSheet : manifest(manifestRows)
};
const reset = isolatedContext({
  LockService: {getScriptLock: () => ({
    waitLock: () => {}, releaseLock: () => resetEvents.push(['unlock'])
  })},
  SpreadsheetApp: {getActiveSpreadsheet: () => resetBook},
  PropertiesService: {getScriptProperties: () => ({
    getProperty: key => resetProperties[key],
    getProperties: () => ({...resetProperties}),
    deleteProperty: key => {delete resetProperties[key];}
  })},
  FormApp: {openById: id => {
    assert.equal(id, 'same-form');
    return {
      deleteAllResponses: () => resetEvents.push(['clear responses']),
      getItems: () => ['first', 'second'],
      deleteItem: item => resetEvents.push(['delete question', item])
    };
  }},
  Logger: {log: () => {}}
});
reset.getOrCreateDashboardSheet_ = () => ({
  getLastRow: () => 3,
  getRange: (...args) => ({clearContent: () => resetEvents.push(['clear dashboard', ...args])})
});
reset.refreshInventorySystem_ = (book, inventorySheet) => {
  assert.equal(book, resetBook);
  assert.equal(inventorySheet, resetSheet);
  resetEvents.push(['rebuild']);
};
reset.resetInventoryData();
assert.equal(resetProperties.INVENTORY_FORM_ID, 'same-form');
assert.equal(resetProperties.UNRELATED_SETTING, 'keep');
assert.equal(Object.keys(resetProperties).length, 2);
assert.ok(resetEvents.some(event => event.join(',') === 'clear stock,2,1,9,6'));
assert.ok(resetEvents.some(event => event.join(',') === 'delete columns,3,4'));
assert.ok(resetEvents.some(event => event.join(',') === 'clear dashboard,2,1,2,4'));
assert.deepEqual(resetEvents.slice(-2), [['rebuild'], ['unlock']]);
resetEvents.length = 0;
reset.readManifestItems_ = () => {throw new Error('Invalid manifest');};
assert.throws(() => reset.resetInventoryData(), /Invalid manifest/);
assert.deepEqual(resetEvents, [['unlock']]);
console.log('PASS: explicit reset scope, form ID preservation and validation before deletion');

const stockHeaders = new Map();
const tags = [];
let capacity = 2;
const manifestData = [headers,
  ['TRS-0001', 'Kopi', 'kg', '2', 'TRUE'],
  ['TRS-0002', 'Kopi', 'kg', '1', 'TRUE']
];
const stockBook = {getSheetByName: () => manifest(manifestData)};
const stockSheet = {
  getParent: () => stockBook,
  getLastColumn: () => Math.max(2, ...stockHeaders.keys()),
  getMaxColumns: () => capacity,
  insertColumnsAfter: (position, amount) => {
    assert.equal(position, capacity);
    capacity += amount;
  },
  getRange: (row, column, height, width) => {
    if (typeof row === 'string') {
      const columnNumber = row.charCodeAt(0) - 64;
      return {addDeveloperMetadata: (key, id) => {
        tags.push({
          getValue: () => id,
          getLocation: () => ({getColumn: () => ({getColumn: () => columnNumber})})
        });
      }};
    }
    return {
      getDisplayValues: () => [
        Array.from({length: width}, (_, index) => stockHeaders.get(column + index) || '')
      ],
      getDisplayValue: () => stockHeaders.get(column) || '',
      getA1Notation: () => String.fromCharCode(column + 64) + row,
      setValue: value => {stockHeaders.set(column, value);}
    };
  },
  createDeveloperMetadataFinder: () => {
    const finder = {
      withKey: () => finder,
      withLocationType: () => finder,
      find: () => tags
    };
    return finder;
  }
};
context.syncInventoryColumnsFromManifest_(stockBook, stockSheet);
context.syncInventoryColumnsFromManifest_(stockBook, stockSheet);
assert.equal(tags.length, 2);
assert.equal(capacity, 4);
assert.equal(stockHeaders.get(3), 'Kopi/kg');
assert.equal(stockHeaders.get(4), 'Kopi/kg');
manifestData[1][1] = 'Kopi Bubuk';
stockHeaders.set(3, 'Manually edited header');
context.syncInventoryColumnsFromManifest_(stockBook, stockSheet);
assert.equal(stockHeaders.get(3), 'Kopi Bubuk/kg');
assert.equal(tags.length, 2);
manifestData[1][4] = 'FALSE';
context.syncInventoryColumnsFromManifest_(stockBook, stockSheet);
const retired = context.getInventoryColumns_(stockSheet).find(column => column.itemId === 'TRS-0001');
assert.equal(retired.columnNumber, 3);
assert.equal(retired.active, false);
assert.equal(retired.title, 'Kopi Bubuk/kg');
stockHeaders.set(5, 'Unknown item');
assert.throws(
  () => context.syncInventoryColumnsFromManifest_(stockBook, stockSheet),
  /Unmanaged inventory column 5/
);
assert.equal(tags.length, 2);
console.log('PASS: repeated sync, duplicate names with distinct IDs, header restoration, retained inactive columns and rejection of unmanaged columns');
