function recordInventoryFormResponse_(event) {
  if (!event || !event.response || !event.source) {
    throw new Error('A Google Forms submission event is required.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const properties = PropertiesService.getScriptProperties();
    const spreadsheetId = properties.getProperty('INVENTORY_SPREADSHEET_ID');
    if (!spreadsheetId) {
      throw new Error('Inventory spreadsheet ID has not been configured.');
    }

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName('stok-barang');
    if (!sheet) {
      throw new Error('The stok-barang sheet was not found.');
    }

    syncInventoryColumnsFromManifest_(spreadsheet, sheet);
    const form = event.source;
    const prefix = getInventoryFormMappingPrefix_(sheet, form.getId());
    const savedProperties = properties.getProperties();
    const inventoryIdByQuestionId = new Map();

    for (const [key, questionId] of Object.entries(savedProperties)) {
      if (!key.startsWith(prefix) || key === prefix + 'INITIALIZED') {
        continue;
      }

      const inventoryId = key.slice(prefix.length);
      if (!MANIFEST_ITEM_ID_PATTERN.test(inventoryId)) {
        throw new Error('Legacy inventory mapping found. Run resetInventoryData once.');
      }
      inventoryIdByQuestionId.set(String(questionId), inventoryId);
    }

    const inventoryColumns = getInventoryColumns_(sheet);
    const columnsByInventoryId = new Map(
      inventoryColumns.map(column => [
        column.itemId,
        column
      ])
    );
    const valuesByColumn = new Map();

    for (const itemResponse of event.response.getItemResponses()) {
      const questionId = String(itemResponse.getItem().getId());
      const inventoryId = inventoryIdByQuestionId.get(questionId);

      // Ignore form questions that are not managed by the inventory mapping.
      if (!inventoryId) {
        continue;
      }

      const response = String(itemResponse.getResponse()).trim();
      if (response === '') {
        continue;
      }
      if (!/^[0-9]+$/.test(response)) {
        throw new Error(
          'Inventory response must be a whole number of zero or greater.'
        );
      }

      const column = columnsByInventoryId.get(inventoryId);
      if (!column || !column.active) {
        throw new Error(
          'No active inventory column exists for ID Barang ' + inventoryId + '.'
        );
      }

      valuesByColumn.set(column.columnNumber, Number(response));
    }

    // An entirely blank submission does not create an empty dated row.
    if (valuesByColumn.size === 0) {
      return;
    }

    const targetRow = findOrCreateTodayRow_(sheet);
    for (const [columnNumber, value] of valuesByColumn) {
      sheet.getRange(targetRow, columnNumber).setValue(value);
    }

    syncDashboardData_(spreadsheet, sheet, inventoryColumns);
  } finally {
    lock.releaseLock();
  }
}

function findOrCreateTodayRow_(sheet) {
  const timeZone = Session.getScriptTimeZone();
  const now = new Date();
  const today = Utilities.formatDate(now, timeZone, 'yyyy-MM-dd');
  const lastRow = sheet.getLastRow();
  const matchingRows = [];

  if (lastRow >= 2) {
    const dates = sheet.getRange(2, 2, lastRow - 1, 1).getValues();

    for (let index = 0; index < dates.length; index++) {
      const value = dates[index][0];
      if (!(value instanceof Date) || Number.isNaN(value.getTime())) {
        continue;
      }

      const date = Utilities.formatDate(value, timeZone, 'yyyy-MM-dd');
      if (date === today) {
        matchingRows.push(index + 2);
      }
    }
  }

  if (matchingRows.length > 1) {
    throw new Error("Multiple stok-barang rows contain today's date.");
  }
  if (matchingRows.length === 1) {
    return matchingRows[0];
  }

  const newRow = Math.max(lastRow + 1, 2);
  sheet.getRange(newRow, 2)
    .setValue(now)
    .setNumberFormat('yyyy-mm-dd');
  return newRow;
}
