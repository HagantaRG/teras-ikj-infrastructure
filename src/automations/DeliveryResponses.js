function recordDeliveryFormResponse_(event) {
  if (!event || !event.source || !event.response) {
    throw new Error('A Google Forms submission event is required.');
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const properties = PropertiesService.getScriptProperties();
    const spreadsheetId = properties.getProperty('DELIVERY_SPREADSHEET_ID');
    const formId = properties.getProperty('DELIVERY_FORM_ID');
    if (!spreadsheetId || formId !== event.source.getId()) {
      throw new Error('This submission does not match the configured delivery form.');
    }
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const manifestItems = getDeliveryManifestItems_(spreadsheet);
    const sheet = getOrCreateDeliverySheet_(spreadsheet);
    const prefix = getDeliveryFormMappingPrefix_(sheet, formId);
    const saved = properties.getProperties();
    const itemIdsByQuestion = new Map();
    for (const [key, questionId] of Object.entries(saved)) {
      if (!key.startsWith(prefix) || key === prefix + 'INITIALIZED') {
        continue;
      }
      const itemId = key.slice(prefix.length);
      if (!MANIFEST_ITEM_ID_PATTERN.test(itemId) || itemIdsByQuestion.has(questionId)) {
        throw new Error('Invalid or duplicate delivery question mapping: ' + key);
      }
      itemIdsByQuestion.set(String(questionId), itemId);
    }

    const timestamp = event.response.getTimestamp();
    if (!(timestamp instanceof Date) || Number.isNaN(timestamp.getTime())) {
      throw new Error('The delivery submission has no valid timestamp.');
    }
    const day = Utilities.formatDate(
      timestamp, Session.getScriptTimeZone(), 'yyyy-MM-dd'
    );
    // Store the chosen calendar day in the destination sheet's timezone.
    const date = Utilities.parseDate(
      day, spreadsheet.getSpreadsheetTimeZone(), 'yyyy-MM-dd'
    );
    const rows = buildDeliveryRows_(
      event.response.getItemResponses(), itemIdsByQuestion,
      properties.getProperty('DELIVERY_NOTES_ITEM_ID'), manifestItems, date
    );
    if (rows.length > 0) {
      appendDeliveryRows_(sheet, rows);

      // A new delivery can change usage for a stock count already recorded today.
      const inventorySheet = spreadsheet.getSheetByName('stok-barang');
      if (!inventorySheet) {
        throw new Error('The stok-barang sheet was not found.');
      }
      syncInventoryColumnsFromManifest_(spreadsheet, inventorySheet);
      const inventoryColumns = getInventoryColumns_(inventorySheet);
      syncDashboardData_(spreadsheet, inventorySheet, inventoryColumns);
    }
  } finally {
    lock.releaseLock();
  }
}

function buildDeliveryRows_(responses, itemIdsByQuestion, notesQuestionId, items, date) {
  const activeIds = new Set(items.filter(item => item.active).map(item => item.id));
  const quantities = new Map();
  let notes = '';

  // Validate the entire submission before appending any rows.
  for (const response of responses) {
    const questionId = String(response.getItem().getId());
    const answer = String(response.getResponse()).trim();
    if (questionId === notesQuestionId) {
      notes = answer;
      continue;
    }
    if (answer === '') {
      continue;
    }
    const itemId = itemIdsByQuestion.get(questionId);
    if (!itemId || !activeIds.has(itemId)) {
      throw new Error('A delivered item is no longer active or mapped: ' + questionId);
    }
    if (!/^0*[1-9][0-9]*$/.test(answer) || !Number.isSafeInteger(Number(answer))) {
      throw new Error('Delivery quantity must be a positive whole number for ' + itemId);
    }
    if (quantities.has(itemId)) {
      throw new Error('Multiple answers were supplied for delivery item ' + itemId);
    }
    quantities.set(itemId, Number(answer));
  }

  // User notes must remain text, including notes beginning with an equals sign.
  const cellNotes = notes.startsWith('=') ? "'" + notes : notes;
  return Array.from(quantities, ([itemId, quantity]) => [
    date, itemId, quantity, cellNotes
  ]);
}
