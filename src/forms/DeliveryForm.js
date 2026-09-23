function createDeliveryForm(event) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const spreadsheet = event && event.source
      ? event.source
      : SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      throw new Error('Run delivery setup from the inventory spreadsheet.');
    }
    const properties = PropertiesService.getScriptProperties();
    const savedSpreadsheetId = properties.getProperty('DELIVERY_SPREADSHEET_ID');
    if (savedSpreadsheetId && savedSpreadsheetId !== spreadsheet.getId()) {
      throw new Error('The delivery form is already configured for another spreadsheet.');
    }
    const items = getDeliveryManifestItems_(spreadsheet);
    const sheet = getOrCreateDeliverySheet_(spreadsheet);
    const formId = properties.getProperty('DELIVERY_FORM_ID');
    const form = formId
      ? FormApp.openById(formId)
      : FormApp.create('Formulir Barang Masuk');

    // Persist IDs immediately so subsequent runs reuse the same form URL.
    properties.setProperty('DELIVERY_FORM_ID', form.getId());
    properties.setProperty('DELIVERY_SPREADSHEET_ID', spreadsheet.getId());
    form.setTitle('Formulir Barang Masuk');
    form.setDescription(
      'Isi jumlah barang yang diterima. Kosongkan barang yang tidak diterima. ' +
      'Tanggal dicatat otomatis sesuai tanggal pengiriman formulir. ' +
      'Kirim satu kali untuk setiap penerimaan barang.'
    );
    form.setConfirmationMessage('Terima kasih. Laporan barang masuk telah dikirim.');

    const validation = FormApp.createTextValidation()
      .requireTextMatchesPattern('^0*[1-9][0-9]*$')
      .setHelpText(
        'Isi dengan bilangan bulat lebih dari 0, misalnya 1, 2, atau 3. ' +
        'Jangan gunakan koma, titik, atau tanda minus. ' +
        'Kosongkan jika barang tidak diterima.'
      )
      .build();
    const fields = items.map(item => ({
      itemId: item.id,
      title: item.displayName,
      active: item.active
    }));
    syncInventoryForm_(
      form, sheet, fields, properties, validation,
      getDeliveryFormMappingPrefix_(sheet, form.getId())
    );
    syncDeliveryNotes_(form, properties);
    Logger.log('Delivery form URL: ' + form.getPublishedUrl());
    Logger.log('Delivery form edit URL: ' + form.getEditUrl());
  } finally {
    lock.releaseLock();
  }
}

function getDeliveryManifestItems_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName('daftar-barang');
  if (!sheet) {
    throw new Error('The daftar-barang sheet was not found.');
  }
  return readManifestItems_(sheet);
}

function getDeliveryFormMappingPrefix_(sheet, formId) {
  return [
    'DELIVERY_ITEM', sheet.getParent().getId(), sheet.getSheetId(), formId
  ].join(':') + ':';
}

function syncDeliveryNotes_(form, properties) {
  const savedId = properties.getProperty('DELIVERY_NOTES_ITEM_ID');
  const existing = form.getItems().find(item => String(item.getId()) === savedId);
  if (existing && existing.getType() !== FormApp.ItemType.PARAGRAPH_TEXT) {
    throw new Error('The delivery notes question is no longer a paragraph field.');
  }
  const notes = existing ? existing.asParagraphTextItem() : form.addParagraphTextItem();
  properties.setProperty('DELIVERY_NOTES_ITEM_ID', String(notes.getId()));
  notes.setTitle('Catatan')
    .setHelpText('Opsional. Catatan ini berlaku untuk semua barang dalam laporan ini.')
    .setRequired(false);
  const lastIndex = form.getItems().length - 1;
  if (notes.getIndex() !== lastIndex) {
    form.moveItem(notes, lastIndex);
  }
}
