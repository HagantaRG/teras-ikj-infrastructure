// Call while holding the script lock to prevent duplicate metadata creation.
function getInventoryColumns_(sheet) {
  const metadataKey = 'INVENTORY_COLUMN';
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 3) {
    return [];
  }

  const headers = sheet
    .getRange(1, 3, 1, lastColumn - 2)
    .getDisplayValues()[0];

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

  const namedColumns = [];
  for (let index = 0; index < headers.length; index++) {
    const title = headers[index].trim();
    if (title === '') {
      continue;
    }

    const columnNumber = index + 3;
    if (!metadataByColumn.has(columnNumber)) {
      sheet.getRange(1, columnNumber, sheet.getMaxRows(), 1)
        .addDeveloperMetadata(
          metadataKey,
          SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT
        );
    }

    namedColumns.push({ columnNumber, title });
  }

  // Read again to obtain the IDs assigned to newly tagged columns.
  for (const metadata of findColumnMetadata()) {
    const columnNumber = metadata.getLocation().getColumn().getColumn();
    metadataByColumn.set(columnNumber, metadata);
  }

  return namedColumns.map(column => ({
    spreadsheetId: sheet.getParent().getId(),
    sheetId: sheet.getSheetId(),
    metadataId: metadataByColumn.get(column.columnNumber).getId(),
    columnNumber: column.columnNumber,
    title: column.title
  }));
}
