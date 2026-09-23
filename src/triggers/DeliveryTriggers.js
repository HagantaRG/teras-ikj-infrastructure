function installDeliveryTriggers() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const properties = PropertiesService.getScriptProperties();
    const formId = properties.getProperty('DELIVERY_FORM_ID');
    if (!spreadsheet || !formId ||
        properties.getProperty('DELIVERY_SPREADSHEET_ID') !== spreadsheet.getId()) {
      throw new Error('Run createDeliveryForm from this spreadsheet first.');
    }
    const form = FormApp.openById(formId);
    const definitions = [
      {
        handler: 'handleDeliveryFormSubmit',
        sourceId: formId,
        eventType: ScriptApp.EventType.ON_FORM_SUBMIT
      },
      {
        handler: 'handleDeliveryManifestEdit',
        sourceId: spreadsheet.getId(),
        eventType: ScriptApp.EventType.ON_EDIT
      },
      {
        handler: 'handleDeliveryManifestChange',
        sourceId: spreadsheet.getId(),
        eventType: ScriptApp.EventType.ON_CHANGE
      }
    ];
    const existing = ScriptApp.getProjectTriggers();
    for (const definition of definitions) {
      const found = existing.some(trigger =>
        trigger.getHandlerFunction() === definition.handler &&
        trigger.getTriggerSourceId() === definition.sourceId &&
        trigger.getEventType() === definition.eventType
      );
      if (found) {
        continue;
      }
      const builder = ScriptApp.newTrigger(definition.handler);
      if (definition.eventType === ScriptApp.EventType.ON_FORM_SUBMIT) {
        builder.forForm(form).onFormSubmit().create();
      } else if (definition.eventType === ScriptApp.EventType.ON_EDIT) {
        builder.forSpreadsheet(spreadsheet).onEdit().create();
      } else {
        builder.forSpreadsheet(spreadsheet).onChange().create();
      }
    }
  } finally {
    lock.releaseLock();
  }
}

function removeDeliveryTriggers() {
  const properties = PropertiesService.getScriptProperties();
  const sourceIds = new Set([
    properties.getProperty('DELIVERY_SPREADSHEET_ID'),
    properties.getProperty('DELIVERY_FORM_ID')
  ]);
  const handlers = new Set([
    'handleDeliveryFormSubmit', 'handleDeliveryManifestEdit', 'handleDeliveryManifestChange'
  ]);
  for (const trigger of ScriptApp.getProjectTriggers()) {
    if (handlers.has(trigger.getHandlerFunction()) &&
        sourceIds.has(trigger.getTriggerSourceId())) {
      ScriptApp.deleteTrigger(trigger);
    }
  }
}

function handleDeliveryFormSubmit(event) {
  recordDeliveryFormResponse_(event);
}

function handleDeliveryManifestEdit(event) {
  if (event && event.range && event.range.getSheet().getName() === 'daftar-barang') {
    refreshConfiguredDeliveryForm_(event);
  }
}

function handleDeliveryManifestChange(event) {
  const changeTypes = ['INSERT_ROW', 'REMOVE_ROW', 'INSERT_COLUMN', 'REMOVE_COLUMN', 'OTHER'];
  if (event && changeTypes.includes(event.changeType)) {
    refreshConfiguredDeliveryForm_(event);
  }
}

function refreshConfiguredDeliveryForm_(event) {
  const properties = PropertiesService.getScriptProperties();
  if (event.source &&
      event.source.getId() === properties.getProperty('DELIVERY_SPREADSHEET_ID') &&
      properties.getProperty('DELIVERY_FORM_ID')) {
    createDeliveryForm(event);
  }
}
