function createDashboardData() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = spreadsheet.getSheetByName('stok-barang');
    if (!sourceSheet) {
      throw new Error('The stok-barang sheet was not found.');
    }

    syncInventoryColumnsFromManifest_(spreadsheet, sourceSheet);
    const columns = getInventoryColumns_(sourceSheet);
    syncDashboardData_(spreadsheet, sourceSheet, columns);
  } finally {
    lock.releaseLock();
  }
}

// Call while holding the script lock.
function syncDashboardData_(spreadsheet, sourceSheet, columns) {
  const dashboardSheet = getOrCreateDashboardSheet_(spreadsheet);
  const existingRows = readDashboardRows_(dashboardSheet);
  if (existingRows.some(row => !MANIFEST_ITEM_ID_PATTERN.test(String(row[1])))) {
    throw new Error('Legacy dashboard IDs found. Run resetInventoryData once.');
  }
  const retainedColumns = columns.filter(column => column.title !== '');
  const retainedInventoryIds = new Set(
    retainedColumns.map(column => column.itemId)
  );

  // Rebuild all retained columns, including inactive items; keep deleted-column history.
  const preservedRows = existingRows.filter(row =>
    !retainedInventoryIds.has(String(row[1]))
  );
  const currentRows = buildCurrentDashboardRows_(sourceSheet, retainedColumns, spreadsheet);
  const dashboardRows = preservedRows.concat(currentRows);

  dashboardRows.sort((first, second) => {
    const dateDifference = first[0].getTime() - second[0].getTime();
    if (dateDifference !== 0) {
      return dateDifference;
    }
    return String(first[2]).localeCompare(String(second[2]), 'id');
  });

  const lastRow = dashboardSheet.getLastRow();
  if (lastRow >= 2) {
    dashboardSheet.getRange(2, 1, lastRow - 1, 5).clearContent();
  }

  if (dashboardRows.length > 0) {
    dashboardSheet
      .getRange(2, 1, dashboardRows.length, 5)
      .setValues(dashboardRows);
    dashboardSheet
      .getRange(2, 1, dashboardRows.length, 1)
      .setNumberFormat('yyyy-mm-dd');
    dashboardSheet
      .getRange(2, 2, dashboardRows.length, 1)
      .setNumberFormat('@');
  }
}

function getOrCreateDashboardSheet_(spreadsheet) {
  const sheetName = 'dashboard-data';
  const headers = ['Tanggal', 'ID Produk', 'Produk/Unit', 'Jumlah', 'Penggunaan'];
  let sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    return sheet;
  }

  const actualHeaders = sheet
    .getRange(1, 1, 1, headers.length)
    .getDisplayValues()[0];

  if (actualHeaders.some((header, index) => header !== headers[index])) {
    throw new Error(
      'The dashboard-data sheet does not have the expected columns.'
    );
  }

  return sheet;
}

function buildCurrentDashboardRows_(sourceSheet, columns, spreadsheet) {
  const lastRow = sourceSheet.getLastRow();
  if (lastRow < 2 || columns.length === 0) {
    return [];
  }

  const lastInventoryColumn = Math.max(
    ...columns.map(column => column.columnNumber)
  );
  const dates = sourceSheet
    .getRange(2, 2, lastRow - 1, 1)
    .getValues();
  const inventoryValues = sourceSheet
    .getRange(2, 3, lastRow - 1, lastInventoryColumn - 2)
    .getValues();
  const rows = [];
  const keys = new Set();
  const timeZone = Session.getScriptTimeZone();

  for (let rowIndex = 0; rowIndex < dates.length; rowIndex++) {
    const date = dates[rowIndex][0];
    if (date === '') {
      continue;
    }
    if (!(date instanceof Date) || Number.isNaN(date.getTime())) {
      throw new Error(
        'stok-barang row ' + (rowIndex + 2) + ' does not contain a valid date.'
      );
    }

    const dateKey = Utilities.formatDate(date, timeZone, 'yyyy-MM-dd');

    for (const column of columns) {
      const value = inventoryValues[rowIndex][column.columnNumber - 3];
      if (value === '' || value === null) {
        continue;
      }

      const quantity = normalizeDashboardQuantity_(
        value,
        rowIndex + 2,
        column.columnNumber
      );
      const inventoryId = column.itemId;
      const key = dateKey + ':' + inventoryId;

      if (keys.has(key)) {
        throw new Error(
          'Duplicate dashboard value for date ' +
          dateKey +
          ' and product ID ' +
          inventoryId +
          '.'
        );
      }
      keys.add(key);

      rows.push([
        new Date(date.getTime()),
        inventoryId,
        column.title,
        quantity,
        ''
      ]);
    }
  }

  return populateUsageValues_(rows, spreadsheet);
}

function populateUsageValues_(rows, spreadsheet) {
  const deliveriesByItem = readDeliveryTotalsByItem_(spreadsheet);
  const timeZone = Session.getScriptTimeZone();
  const previousStockByItem = new Map();

  rows.sort((first, second) => {
    const dateDifference = first[0].getTime() - second[0].getTime();
    return dateDifference || String(first[1]).localeCompare(String(second[1]));
  });

  for (const row of rows) {
    const inventoryId = row[1];
    const currentDay = Utilities.formatDate(row[0], timeZone, 'yyyy-MM-dd');
    const previousStock = previousStockByItem.get(inventoryId);

    if (previousStock) {
      const deliveryRows = deliveriesByItem.get(inventoryId) || new Map();
      const deliveries = sumDeliveriesBetween_(
        deliveryRows,
        previousStock.day,
        currentDay
      );
      const usage = previousStock.quantity + deliveries - row[3];
      if (usage >= 0) {
        row[4] = usage;
      }
    }

    previousStockByItem.set(inventoryId, {
      day: currentDay,
      quantity: row[3]
    });
  }

  return rows;
}

function readDeliveryTotalsByItem_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName('barang-masuk');
  if (!sheet) {
    return new Map();
  }

  const expectedHeaders = ['Tanggal', 'ID Barang', 'Jumlah Masuk', 'Catatan'];
  const actualHeaders = sheet.getRange(1, 1, 1, expectedHeaders.length)
    .getDisplayValues()[0];
  if (actualHeaders.some((header, index) => header !== expectedHeaders[index])) {
    throw new Error('The barang-masuk sheet does not have the expected columns.');
  }

  const totalsByItem = new Map();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return totalsByItem;
  }

  const records = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  const timeZone = Session.getScriptTimeZone();

  for (let index = 0; index < records.length; index++) {
    const [date, rawId, rawQuantity, note] = records[index];
    if (date === '' && rawId === '' && rawQuantity === '' && note === '') {
      continue;
    }
    if (!(date instanceof Date) || Number.isNaN(date.getTime())) {
      throw new Error('barang-masuk row ' + (index + 2) + ' has an invalid date.');
    }

    const inventoryId = String(rawId).trim();
    if (!MANIFEST_ITEM_ID_PATTERN.test(inventoryId)) {
      throw new Error('barang-masuk row ' + (index + 2) + ' has an invalid ID Barang.');
    }
    const quantity = normalizeDeliveryQuantity_(rawQuantity, index + 2);
    const day = Utilities.formatDate(date, timeZone, 'yyyy-MM-dd');
    if (!totalsByItem.has(inventoryId)) {
      totalsByItem.set(inventoryId, new Map());
    }
    const totalsByDay = totalsByItem.get(inventoryId);
    const total = (totalsByDay.get(day) || 0) + quantity;
    if (!Number.isSafeInteger(total)) {
      throw new Error('The delivery total is too large for ' + inventoryId + ' on ' + day + '.');
    }
    totalsByDay.set(day, total);
  }

  return totalsByItem;
}

function normalizeDeliveryQuantity_(value, rowNumber) {
  const quantity = typeof value === 'number'
    ? value
    : Number(String(value).trim());
  if (!Number.isSafeInteger(quantity) || quantity <= 0) {
    throw new Error(
      'barang-masuk row ' + rowNumber +
      ' must contain a positive whole-number delivery quantity.'
    );
  }
  return quantity;
}

function sumDeliveriesBetween_(deliveriesByDay, previousStockDay, currentStockDay) {
  let total = 0;
  for (const [deliveryDay, quantity] of deliveriesByDay) {
    if (deliveryDay > previousStockDay && deliveryDay <= currentStockDay) {
      total += quantity;
    }
  }
  if (!Number.isSafeInteger(total)) {
    throw new Error('Delivery total between recorded stock dates is too large.');
  }
  return total;
}

function readDashboardRows_(dashboardSheet) {
  const lastRow = dashboardSheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  return dashboardSheet
    .getRange(2, 1, lastRow - 1, 5)
    .getValues()
    .filter(row => row.some(value => value !== ''));
}

function normalizeDashboardQuantity_(value, rowNumber, columnNumber) {
  const normalizedValue = typeof value === 'string'
    ? value.trim()
    : value;
  const quantity = Number(normalizedValue);

  if (!Number.isInteger(quantity) || quantity < 0) {
    throw new Error(
      'Invalid inventory quantity at row ' +
      rowNumber +
      ', column ' +
      columnNumber +
      '.'
    );
  }

  return quantity;
}

function syncDashboardAfterSheetEdit_(event) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const spreadsheet = event.source;
    const sourceSheet = spreadsheet.getSheetByName('stok-barang');
    if (!sourceSheet) {
      throw new Error('The stok-barang sheet was not found.');
    }

    syncInventoryColumnsFromManifest_(spreadsheet, sourceSheet);
    const columns = getInventoryColumns_(sourceSheet);
    syncDashboardData_(spreadsheet, sourceSheet, columns);
  } finally {
    lock.releaseLock();
  }
}
