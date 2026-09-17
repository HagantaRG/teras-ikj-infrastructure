const INVENTORY_ITEM_METADATA_KEY = 'INVENTORY_ITEM_ID';
const MANIFEST_ITEM_ID_PATTERN = /^TRS-[0-9]{4}$/;

// Call while holding the script lock.
function syncInventoryColumnsFromManifest_(spreadsheet, inventorySheet) {
  const manifestSheet = spreadsheet.getSheetByName('daftar-barang');
  if (!manifestSheet) {
    throw new Error('The daftar-barang sheet was not found.');
  }

  const manifestItems = readManifestItems_(manifestSheet);
  const activeItemIds = new Set(
    manifestItems
      .filter(item => item.active)
      .map(item => item.id)
  );
  const columns = getInventoryColumns_(inventorySheet);
  const columnsByItemId = readInventoryItemMetadata_(inventorySheet);
  const metadataColumnNumbers = new Set(
    Array.from(columnsByItemId.values()).map(column => column.columnNumber)
  );
  const manifestNames = new Map();

  for (const item of manifestItems) {
    if (manifestNames.has(item.name)) {
      manifestNames.set(item.name, null);
    } else {
      manifestNames.set(item.name, item.id);
    }
  }

  // Migrate existing columns whose header exactly identifies one manifest item.
  for (const column of columns) {
    if (metadataColumnNumbers.has(column.columnNumber)) {
      continue;
    }

    const itemId = manifestNames.get(column.title);
    if (!itemId) {
      continue;
    }

    addInventoryItemMetadata_(inventorySheet, column.columnNumber, itemId);
    metadataColumnNumbers.add(column.columnNumber);
    columnsByItemId.set(itemId, {
      columnNumber: column.columnNumber,
      itemId
    });
  }

  // Add new active items and update the names of existing items.
  for (const item of manifestItems.filter(item => item.active)) {
    let column = columnsByItemId.get(item.id);

    if (!column) {
      const newColumnNumber = Math.max(inventorySheet.getLastColumn() + 1, 3);
      inventorySheet.insertColumnAfter(Math.max(newColumnNumber - 1, 2));
      addInventoryItemMetadata_(inventorySheet, newColumnNumber, item.id);
      column = {
        columnNumber: newColumnNumber,
        itemId: item.id
      };
      columnsByItemId.set(item.id, column);
    }

    inventorySheet
      .getRange(1, column.columnNumber)
      .setValue(item.name);
  }

  // Keep removed columns for history, but remove their headers so they no
  // longer appear as questions in the form.
  for (const column of readInventoryItemMetadata_(inventorySheet).values()) {
    if (!activeItemIds.has(column.itemId)) {
      inventorySheet.getRange(1, column.columnNumber).clearContent();
    }
  }
}

function readManifestItems_(manifestSheet) {
  const lastRow = manifestSheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  const values = manifestSheet.getRange(2, 1, lastRow - 1, 4).getDisplayValues();
  const seenIds = new Set();
  const items = [];

  for (let index = 0; index < values.length; index++) {
    const [id, name, minimumStock, activeValue] = values[index].map(value => value.trim());
    if (!id && !name && !minimumStock && !activeValue) {
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

    seenIds.add(id);
    items.push({
      id,
      name,
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

  for (const entry of metadata) {
    const itemId = entry.getValue();
    const columnNumber = entry.getLocation().getColumn().getColumn();
    if (!itemId || columnsByItemId.has(itemId)) {
      throw new Error('Each inventory item must have exactly one column metadata ID.');
    }
    columnsByItemId.set(itemId, { columnNumber, itemId });
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
