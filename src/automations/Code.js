function createFormFromSheet() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName('stok-barang');

    if (!sheet || sheet.getLastColumn() < 3) {
      throw new Error('Expected inventory headers in stok-barang, starting at C1.');
    }

    const headers = sheet
      .getRange(1, 3, 1, sheet.getLastColumn() - 2)
      .getDisplayValues()[0]
      .map(header => header.trim())
      .filter(header => header !== '');

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
      form = FormApp.create('Dynamic Form');
      properties.setProperty('INVENTORY_FORM_ID', form.getId());
    }

    // Keep existing questions and add only missing inventory headers.
    const existingTitles = new Set(
      form.getItems().map(item => item.getTitle())
    );

    for (const header of headers) {
      if (!existingTitles.has(header)) {
        form.addTextItem()
          .setTitle(header)
          .setRequired(true);
        existingTitles.add(header);
      }
    }

    Logger.log('Form URL: ' + form.getPublishedUrl());
    Logger.log('Edit URL: ' + form.getEditUrl());
  } finally {
    lock.releaseLock();
  }
}
