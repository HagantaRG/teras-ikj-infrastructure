// Destructive setup operation: run manually once, never from a trigger.
function resetInventoryData() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const inventorySheet = spreadsheet.getSheetByName('stok-barang');
    const manifestSheet = spreadsheet.getSheetByName('daftar-barang');
    if (!inventorySheet || !manifestSheet) {
      throw new Error('Both stok-barang and daftar-barang are required.');
    }

    // Validate the manifest and open existing resources before deleting data.
    readManifestItems_(manifestSheet);
    const properties = PropertiesService.getScriptProperties();
    const formId = properties.getProperty('INVENTORY_FORM_ID');
    const form = formId ? FormApp.openById(formId) : null;
    const dashboardSheet = getOrCreateDashboardSheet_(spreadsheet);

    if (form) {
      form.deleteAllResponses();
      for (const item of form.getItems().reverse()) {
        form.deleteItem(item);
      }
      const prefix = getInventoryFormMappingPrefix_(inventorySheet, formId);
      for (const key of Object.keys(properties.getProperties())) {
        if (key.startsWith(prefix)) {
          properties.deleteProperty(key);
        }
      }
    }

    for (const key of ['INVENTORY_COLUMN', INVENTORY_ITEM_METADATA_KEY]) {
      const entries = inventorySheet.createDeveloperMetadataFinder()
        .withKey(key)
        .find();
      for (const entry of entries) {
        entry.remove();
      }
    }

    // Preserve columns A/B and their headings; clear all old dated records.
    if (inventorySheet.getMaxRows() > 1) {
      inventorySheet.getRange(
        2, 1, inventorySheet.getMaxRows() - 1, inventorySheet.getMaxColumns()
      ).clearContent();
    }
    if (inventorySheet.getMaxColumns() > 2) {
      inventorySheet.deleteColumns(3, inventorySheet.getMaxColumns() - 2);
    }
    if (dashboardSheet.getLastRow() > 1) {
      dashboardSheet.getRange(2, 1, dashboardSheet.getLastRow() - 1, 4)
        .clearContent();
    }

    refreshInventorySystem_(spreadsheet, inventorySheet);
    Logger.log('Inventory reset completed. Run installInventoryTriggers if needed.');
  } finally {
    lock.releaseLock();
  }
}
