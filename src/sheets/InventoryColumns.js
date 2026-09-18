// Inventory identity comes from ID Barang, never a header or Google's metadata ID.
function getInventoryColumns_(sheet) {
  const manifestSheet = sheet.getParent().getSheetByName('daftar-barang');
  if (!manifestSheet) {
    throw new Error('The daftar-barang sheet was not found.');
  }
  const itemsById = new Map(
    readManifestItems_(manifestSheet).map(item => [item.id, item])
  );

  return Array.from(readInventoryItemMetadata_(sheet).values())
    .sort((first, second) => first.columnNumber - second.columnNumber)
    .map(column => {
      const item = itemsById.get(column.itemId);
      return {
        itemId: column.itemId,
        columnNumber: column.columnNumber,
        title: item ? item.displayName : column.title,
        active: Boolean(item && item.active)
      };
    });
}
