function createFormFromSheet(event) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const spreadsheet = event && event.source
      ? event.source
      : SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName('stok-barang');

    if (!sheet) {
      throw new Error('The stok-barang sheet was not found.');
    }

    syncInventoryColumnsFromManifest_(spreadsheet, sheet);
    const inventoryColumns = getInventoryColumns_(sheet);
    syncDashboardData_(spreadsheet, sheet, inventoryColumns);

    // Save the form ID once, then reuse that form on subsequent runs.
    const properties = PropertiesService.getScriptProperties();
    const savedFormId = properties.getProperty('INVENTORY_FORM_ID');
    if (!savedFormId && !inventoryColumns.some(column => column.title !== '')) {
      return;
    }
    let form;

    if (savedFormId) {
      form = FormApp.openById(savedFormId);
    } else {
      form = FormApp.create('Formulir Stok Barang TERAS IKJ');
      properties.setProperty('INVENTORY_FORM_ID', form.getId());
    }

    form.setTitle('Formulir Stok Barang');

    const integerValidation = FormApp.createTextValidation()
      .requireTextMatchesPattern('^[0-9]+$')
      .setHelpText('Isi dengan angka seperti 0, 1, 2, atau 3, tanpa koma, titik, atau tanda minus. Boleh dikosongkan.')
      .build();

    syncInventoryForm_(form, sheet, inventoryColumns, properties, integerValidation);

    Logger.log('Form URL: ' + form.getPublishedUrl());
    Logger.log('Edit URL: ' + form.getEditUrl());
  } finally {
    lock.releaseLock();
  }
}
