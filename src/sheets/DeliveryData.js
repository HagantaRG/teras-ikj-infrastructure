function getOrCreateDeliverySheet_(spreadsheet) {
  const name = 'barang-masuk';
  const headers = ['Tanggal', 'ID Barang', 'Jumlah Masuk', 'Catatan'];
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(name);
  }
  if (sheet.getLastRow() === 0) {
    if (sheet.getMaxColumns() < headers.length) {
      sheet.insertColumnsAfter(
        sheet.getMaxColumns(), headers.length - sheet.getMaxColumns()
      );
    }
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  } else {
    const actual = sheet.getRange(1, 1, 1, headers.length).getDisplayValues()[0];
    if (actual.some((value, index) => value !== headers[index])) {
      throw new Error(
        'barang-masuk must have these headers in order: ' + headers.join(', ') + '.'
      );
    }
  }
  return sheet;
}

// Call while holding the script lock.
function appendDeliveryRows_(sheet, rows) {
  if (rows.length === 0) {
    return;
  }
  const firstRow = sheet.getLastRow() + 1;
  const lastRow = firstRow + rows.length - 1;
  if (lastRow > sheet.getMaxRows()) {
    sheet.insertRowsAfter(sheet.getMaxRows(), lastRow - sheet.getMaxRows());
  }
  sheet.getRange(firstRow, 1, rows.length, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(firstRow, 2, rows.length, 1).setNumberFormat('@');
  sheet.getRange(firstRow, 3, rows.length, 1).setNumberFormat('0');
  sheet.getRange(firstRow, 4, rows.length, 1).setNumberFormat('@');
  sheet.getRange(firstRow, 1, rows.length, 4).setValues(rows);
}
