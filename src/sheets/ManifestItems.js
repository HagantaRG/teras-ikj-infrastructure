const INVENTORY_ITEM_METADATA_KEY = 'INVENTORY_ITEM_ID';
const MANIFEST_ITEM_ID_PATTERN = /^TRS-[0-9]{4}$/;

// Call while holding the script lock. Never adopt columns by their header.
function syncInventoryColumnsFromManifest_(spreadsheet, inventorySheet) {
  const manifestSheet = spreadsheet.getSheetByName('daftar-barang');
  if (!manifestSheet) {
    throw new Error('The daftar-barang sheet was not found.');
  }
  const items = readManifestItems_(manifestSheet);
  const columnsById = readInventoryItemMetadata_(inventorySheet);
  const managedColumnNumbers = new Set(
    Array.from(columnsById.values()).map(column => column.columnNumber)
  );
  const lastColumn = inventorySheet.getLastColumn();

  // Refuse legacy or manually created headers instead of inventing item IDs.
  if (lastColumn >= 3) {
    const headers = inventorySheet.getRange(1, 3, 1, lastColumn - 2)
      .getDisplayValues()[0];
    for (let index = 0; index < headers.length; index++) {
      if (headers[index].trim() && !managedColumnNumbers.has(index + 3)) {
        throw new Error(
          'Unmanaged inventory column ' + (index + 3) +
          '. For initial setup, run resetInventoryData. Manage items in daftar-barang.'
        );
      }
    }
  }

  let nextColumn = Math.max(
    2, lastColumn, ...managedColumnNumbers
  ) + 1;
  for (const item of items) {
    let column = columnsById.get(item.id);
    if (!column && !item.active) {
      continue;
    }
    if (!column) {
      if (nextColumn > inventorySheet.getMaxColumns()) {
        inventorySheet.insertColumnsAfter(
          inventorySheet.getMaxColumns(),
          nextColumn - inventorySheet.getMaxColumns()
        );
      }
      addInventoryItemMetadata_(inventorySheet, nextColumn, item.id);
      column = { columnNumber: nextColumn, itemId: item.id };
      columnsById.set(item.id, column);
      nextColumn++;
    }
    inventorySheet.getRange(1, column.columnNumber).setValue(item.displayName);
  }

  // Inactive/removed columns stay intact for historical corrections.
  // Form inclusion is decided by the manifest's Aktif value, not the header.
}

function readManifestItems_(manifestSheet) {
  const lastRow = manifestSheet.getLastRow();
  const lastColumn = manifestSheet.getLastColumn();
  if (lastColumn === 0) {
    throw new Error('The daftar-barang sheet has no headers.');
  }

  const headers = manifestSheet
    .getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0]
    .map(header => header.trim().toLowerCase());
  const columnNumbers = {
    id: headers.indexOf('id barang'),
    name: headers.indexOf('nama barang'),
    unit: headers.indexOf('unit'),
    minimumStock: headers.indexOf('level stok minim'),
    active: headers.indexOf('aktif')
  };
  const missingHeaders = Object.entries(columnNumbers)
    .filter(([, columnNumber]) => columnNumber === -1)
    .map(([name]) => name);
  if (missingHeaders.length > 0) {
    throw new Error(
      'The daftar-barang sheet is missing required headers: ' +
      missingHeaders.join(', ') + '.'
    );
  }

  for (const requiredHeader of [
    'id barang', 'nama barang', 'unit', 'level stok minim', 'aktif'
  ]) {
    if (headers.filter(header => header === requiredHeader).length > 1) {
      throw new Error('Duplicate daftar-barang header: ' + requiredHeader);
    }
  }
  if (lastRow < 2) {
    return [];
  }

  const values = manifestSheet
    .getRange(2, 1, lastRow - 1, lastColumn)
    .getDisplayValues();
  const seenIds = new Set();
  const items = [];

  for (let index = 0; index < values.length; index++) {
    const row = values[index].map(value => value.trim());
    const id = row[columnNumbers.id];
    const name = row[columnNumbers.name];
    const unit = row[columnNumbers.unit];
    const minimumStock = row[columnNumbers.minimumStock];
    const activeValue = row[columnNumbers.active];
    if (!id && !name && !unit && !minimumStock &&
        (activeValue === '' || activeValue.toUpperCase() === 'FALSE')) {
      continue;
    }

    if (!MANIFEST_ITEM_ID_PATTERN.test(id)) {
      throw new Error(
        'The ID in daftar-barang row ' + (index + 2) +
        ' must use the format TRS-0000.'
      );
    }
    if (seenIds.has(id)) {
      throw new Error('The ID ' + id + ' appears more than once in daftar-barang.');
    }
    if (!name) {
      throw new Error('The item name in daftar-barang row ' + (index + 2) + ' is empty.');
    }
    if (!unit) {
      throw new Error('The unit in daftar-barang row ' + (index + 2) + ' is empty.');
    }

    if (!['', 'TRUE', 'FALSE'].includes(activeValue.toUpperCase())) {
      throw new Error('Aktif must be TRUE or FALSE in daftar-barang row ' + (index + 2) + '.');
    }

    seenIds.add(id);
    items.push({
      id,
      name,
      unit,
      displayName: name + '/' + unit,
      minimumStock,
      active: activeValue === '' || activeValue.toUpperCase() === 'TRUE'
    });
  }

  return items;
}

function readInventoryItemMetadata_(sheet) {
  const metadata = sheet.createDeveloperMetadataFinder()
    .withKey(INVENTORY_ITEM_METADATA_KEY)
    .withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.COLUMN)
    .find();
  const columnsByItemId = new Map();
  const seenColumns = new Set();

  for (const entry of metadata) {
    const itemId = entry.getValue();
    const columnNumber = entry.getLocation().getColumn().getColumn();
    if (!MANIFEST_ITEM_ID_PATTERN.test(itemId) ||
        columnsByItemId.has(itemId) || seenColumns.has(columnNumber) ||
        columnNumber < 3) {
      throw new Error('Each inventory item must have exactly one column metadata ID.');
    }
    seenColumns.add(columnNumber);
    columnsByItemId.set(itemId, {
      columnNumber,
      itemId,
      title: sheet.getRange(1, columnNumber).getDisplayValue().trim()
    });
  }

  return columnsByItemId;
}

function addInventoryItemMetadata_(sheet, columnNumber, itemId) {
  const columnName = sheet
    .getRange(1, columnNumber)
    .getA1Notation()
    .replace(/[0-9]+$/, '');
  sheet.getRange(columnName + ':' + columnName).addDeveloperMetadata(
    INVENTORY_ITEM_METADATA_KEY,
    itemId,
    SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT
  );
}
