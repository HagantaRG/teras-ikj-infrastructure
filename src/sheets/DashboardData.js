function createDashboardData() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = spreadsheet.getSheetByName('stok-barang');
    if (!sourceSheet) {
      throw new Error('The stok-barang sheet was not found.');
    }

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
  const activeColumns = columns.filter(column => column.title !== '');
  const activeMetadataIds = new Set(
    activeColumns.map(column => String(column.metadataId))
  );

  // Keep history for columns that have been deleted or temporarily unnamed.
  const preservedRows = existingRows.filter(row =>
    !activeMetadataIds.has(String(row[1]))
  );
  const currentRows = buildCurrentDashboardRows_(sourceSheet, activeColumns);
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
    dashboardSheet.getRange(2, 1, lastRow - 1, 4).clearContent();
  }

  if (dashboardRows.length > 0) {
    dashboardSheet
      .getRange(2, 1, dashboardRows.length, 4)
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
  const headers = ['Tanggal', 'ID Produk', 'Produk', 'Jumlah'];
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

function buildCurrentDashboardRows_(sourceSheet, columns) {
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
      const metadataId = String(column.metadataId);
      const key = dateKey + ':' + metadataId;

      if (keys.has(key)) {
        throw new Error(
          'Duplicate dashboard value for date ' +
          dateKey +
          ' and product ID ' +
          metadataId +
          '.'
        );
      }
      keys.add(key);

      rows.push([
        new Date(date.getTime()),
        metadataId,
        column.title,
        quantity
      ]);
    }
  }

  return rows;
}

function readDashboardRows_(dashboardSheet) {
  const lastRow = dashboardSheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  return dashboardSheet
    .getRange(2, 1, lastRow - 1, 4)
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

    const columns = getInventoryColumns_(sourceSheet);
    syncDashboardData_(spreadsheet, sourceSheet, columns);
  } finally {
    lock.releaseLock();
  }
}
