function createFormFromSheet(event) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const spreadsheet = event && event.source
      ? event.source
      : SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName('stok-barang');

    if (!sheet || sheet.getLastColumn() < 3) {
      throw new Error('Expected inventory headers in stok-barang, starting at C1.');
    }

    const inventoryColumns = getInventoryColumns_(sheet);
    const headers = inventoryColumns.map(column => column.title);

    if (headers.length === 0) {
      throw new Error('No inventory headers found.');
    }

    // Save the form ID once, then reuse that form on subsequent runs.
    const properties = PropertiesService.getScriptProperties();
    const savedFormId = properties.getProperty('INVENTORY_FORM_ID');
    let form;

    if (savedFormId) {
      form = FormApp.openById(savedFormId);
    } else {
      form = FormApp.create('Formulir Stok Barang');
      properties.setProperty('INVENTORY_FORM_ID', form.getId());
    }

    form.setTitle('Formulir Stok Barang');

    // Keep existing questions and add only missing inventory headers.
    const existingTitles = new Set(
      form.getItems().map(item => item.getTitle())
    );

    for (const header of headers) {
      if (!existingTitles.has(header)) {
        form.addTextItem()
          .setTitle(header);
        existingTitles.add(header);
      }
    }

    const integerValidation = FormApp.createTextValidation()
      .requireTextMatchesPattern('^[0-9]+$')
      .setHelpText('Isi dengan angka seperti 0, 1, 2, atau 3, tanpa koma, titik, atau tanda minus. Boleh dikosongkan.')
      .build();

    // Apply the rules to both existing and newly added text fields.
    for (const item of form.getItems(FormApp.ItemType.TEXT)) {
      item.asTextItem()
        .setRequired(false)
        .setValidation(integerValidation);
    }

    Logger.log('Form URL: ' + form.getPublishedUrl());
    Logger.log('Edit URL: ' + form.getEditUrl());
  } finally {
    lock.releaseLock();
  }
}
