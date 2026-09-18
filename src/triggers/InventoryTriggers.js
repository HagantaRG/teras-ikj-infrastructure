function installInventoryTriggers() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet || !spreadsheet.getSheetByName('stok-barang')) {
    throw new Error('Run setup from the spreadsheet containing stok-barang.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const spreadsheetId = spreadsheet.getId();
    const triggers = ScriptApp.getProjectTriggers();
    const properties = PropertiesService.getScriptProperties();
    const formId = properties.getProperty('INVENTORY_FORM_ID');

    if (!formId) {
      throw new Error(
        'Run createFormFromSheet before installing inventory triggers.'
      );
    }

    properties.setProperty('INVENTORY_SPREADSHEET_ID', spreadsheetId);
    const form = FormApp.openById(formId);
    const definitions = [
      {
        handler: 'handleInventoryColumnChange',
        eventType: ScriptApp.EventType.ON_CHANGE
      },
      {
        handler: 'handleInventoryHeaderEdit',
        eventType: ScriptApp.EventType.ON_EDIT
      }
    ];

    for (const definition of definitions) {
      const exists = triggers.some(trigger =>
        trigger.getHandlerFunction() === definition.handler &&
        trigger.getTriggerSourceId() === spreadsheetId &&
        trigger.getEventType() === definition.eventType
      );

      if (exists) {
        continue;
      }

      const builder = ScriptApp.newTrigger(definition.handler)
        .forSpreadsheet(spreadsheet);

      if (definition.eventType === ScriptApp.EventType.ON_CHANGE) {
        builder.onChange().create();
      } else {
        builder.onEdit().create();
      }
    }

    const formTriggerExists = triggers.some(trigger =>
      trigger.getHandlerFunction() === 'handleInventoryFormSubmit' &&
      trigger.getTriggerSourceId() === formId &&
      trigger.getEventType() === ScriptApp.EventType.ON_FORM_SUBMIT
    );

    if (!formTriggerExists) {
      ScriptApp.newTrigger('handleInventoryFormSubmit')
        .forForm(form)
        .onFormSubmit()
        .create();
    }
  } finally {
    lock.releaseLock();
  }
}

function removeInventoryTriggers() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('Run removal from the inventory spreadsheet.');
  }

  const handlers = [
    'handleInventoryColumnChange',
    'handleInventoryHeaderEdit',
    'handleInventoryFormSubmit'
  ];
  const formId = PropertiesService.getScriptProperties()
    .getProperty('INVENTORY_FORM_ID');
  const sourceIds = new Set([spreadsheet.getId(), formId]);

  for (const trigger of ScriptApp.getProjectTriggers()) {
    if (
      handlers.includes(trigger.getHandlerFunction()) &&
      sourceIds.has(trigger.getTriggerSourceId())
    ) {
      ScriptApp.deleteTrigger(trigger);
    }
  }
}

function handleInventoryFormSubmit(event) {
  recordInventoryFormResponse_(event);
}

function handleInventoryColumnChange(event) {
  if (!event || !['INSERT_COLUMN', 'REMOVE_COLUMN', 'INSERT_ROW', 'REMOVE_ROW', 'OTHER'].includes(event.changeType)) {
    return;
  }

  createFormFromSheet(event);
}

function handleInventoryHeaderEdit(event) {
  if (!event || !event.range) {
    return;
  }

  const range = event.range;
  const editedSheetName = range.getSheet().getName();
  if (editedSheetName === 'daftar-barang') {
    // Headers may be reordered, and a paste may include the header row.
    createFormFromSheet(event);
    return;
  }

  if (editedSheetName !== 'stok-barang') {
    return;
  }

  const touchesInventoryHeader =
    range.getRow() === 1 && range.getLastColumn() >= 3;
  if (touchesInventoryHeader) {
    createFormFromSheet(event);
    return;
  }

  const touchesInventoryData =
    range.getLastRow() >= 2 && range.getLastColumn() >= 2;
  if (touchesInventoryData) {
    syncDashboardAfterSheetEdit_(event);
  }
}
