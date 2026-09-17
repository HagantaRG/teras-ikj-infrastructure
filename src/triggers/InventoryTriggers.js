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
    'handleInventoryHeaderEdit'
  ];

  for (const trigger of ScriptApp.getProjectTriggers()) {
    if (
      handlers.includes(trigger.getHandlerFunction()) &&
      trigger.getTriggerSourceId() === spreadsheet.getId()
    ) {
      ScriptApp.deleteTrigger(trigger);
    }
  }
}

function handleInventoryColumnChange(event) {
  if (!event || !['INSERT_COLUMN', 'REMOVE_COLUMN', 'OTHER'].includes(event.changeType)) {
    return;
  }

  createFormFromSheet(event);
}

function handleInventoryHeaderEdit(event) {
  if (!event || !event.range) {
    return;
  }

  const range = event.range;
  if (range.getSheet().getName() !== 'stok-barang' ||
      range.getRow() !== 1 || range.getLastColumn() < 3) {
    return;
  }
  createFormFromSheet(event);
}
