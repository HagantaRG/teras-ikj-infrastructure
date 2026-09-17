// Call while holding the script lock to prevent duplicate metadata creation.
function getInventoryColumns_(sheet) {
  const metadataKey = 'INVENTORY_COLUMN';
  const lastColumn = sheet.getLastColumn();
  const headers = lastColumn >= 3
    ? sheet.getRange(1, 3, 1, lastColumn - 2).getDisplayValues()[0]
    : [];

  function findColumnMetadata() {
    return sheet.createDeveloperMetadataFinder()
      .withKey(metadataKey)
      .withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.COLUMN)
      .find();
  }

  const metadataByColumn = new Map();
  for (const metadata of findColumnMetadata()) {
    const columnNumber = metadata.getLocation().getColumn().getColumn();
    if (metadataByColumn.has(columnNumber)) {
      throw new Error(
        'Multiple inventory metadata tags found in column ' + columnNumber + '.'
      );
    }
    metadataByColumn.set(columnNumber, metadata);
  }

  for (let index = 0; index < headers.length; index++) {
    const title = headers[index].trim();
    if (title === '') {
      continue;
    }

    const columnNumber = index + 3;
    if (!metadataByColumn.has(columnNumber)) {
      const columnName = sheet
        .getRange(1, columnNumber)
        .getA1Notation()
        .replace(/[0-9]+$/, '');
      sheet.getRange(columnName + ':' + columnName).addDeveloperMetadata(
        metadataKey,
        SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT
      );
    }
  }

  // Read again to obtain the IDs assigned to newly tagged columns.
  for (const metadata of findColumnMetadata()) {
    const columnNumber = metadata.getLocation().getColumn().getColumn();
    metadataByColumn.set(columnNumber, metadata);
  }

  const identifiedColumns = Array.from(metadataByColumn.entries())
    .sort(([firstColumn], [secondColumn]) => firstColumn - secondColumn);

  return identifiedColumns.map(([columnNumber, metadata]) => ({
    spreadsheetId: sheet.getParent().getId(),
    sheetId: sheet.getSheetId(),
    metadataId: metadata.getId(),
    columnNumber,
    title: columnNumber >= 3 ? (headers[columnNumber - 3] || '').trim() : ''
  }));
}
