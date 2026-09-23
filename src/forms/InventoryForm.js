function getInventoryFormMappingPrefix_(sheet, formId) {
  return [
    'INVENTORY_ITEM',
    sheet.getParent().getId(),
    sheet.getSheetId(),
    formId
  ].join(':') + ':';
}

// Call while holding the script lock.
function syncInventoryForm_(form, sheet, columns, properties, validation, mappingPrefix) {
  const prefix = mappingPrefix || getInventoryFormMappingPrefix_(sheet, form.getId());
  const initializedKey = prefix + 'INITIALIZED';
  const savedProperties = properties.getProperties();
  const mappings = new Map();
  const claimedItemIds = new Set();
  const items = form.getItems();
  const itemsById = new Map(
    items.map(item => [String(item.getId()), item])
  );

  // Validate saved associations before changing any questions.
  for (const [key, itemId] of Object.entries(savedProperties)) {
    if (!key.startsWith(prefix) || key === initializedKey) {
      continue;
    }

    const inventoryId = key.slice(prefix.length);
    if (!MANIFEST_ITEM_ID_PATTERN.test(inventoryId) || !/^-?[0-9]+$/.test(itemId)) {
      throw new Error('Invalid inventory question mapping: ' + key);
    }
    if (claimedItemIds.has(itemId)) {
      throw new Error('Multiple inventory columns reference question ' + itemId);
    }

    const item = itemsById.get(itemId);
    if (item && item.getType() !== FormApp.ItemType.TEXT) {
      throw new Error('Mapped inventory question is no longer a text field: ' + itemId);
    }

    mappings.set(inventoryId, itemId);
    claimedItemIds.add(itemId);
  }

  const activeColumns = columns.filter(column => column.active);
  const plans = activeColumns.map(column => ({
    column,
    inventoryId: column.itemId,
    item: itemsById.get(mappings.get(column.itemId))
  }));

  // Persist each association as it is established so retries can reuse it.
  for (const plan of plans) {
    const item = plan.item
      ? plan.item.asTextItem()
      : form.addTextItem().setTitle(plan.column.title);
    const itemId = String(item.getId());

    properties.setProperty(prefix + plan.inventoryId, itemId);
    mappings.set(plan.inventoryId, itemId);
    itemsById.set(itemId, item);

    item.setTitle(plan.column.title)
      .setRequired(false)
      .setValidation(validation);
  }

  // Only active manifest items belong in the form.
  const liveInventoryIds = new Set(
    activeColumns.map(column => column.itemId)
  );
  for (const [inventoryId, itemId] of mappings) {
    if (liveInventoryIds.has(inventoryId)) {
      continue;
    }

    const item = itemsById.get(itemId);
    if (item) {
      form.deleteItem(item);
    }
    properties.deleteProperty(prefix + inventoryId);
    mappings.delete(inventoryId);
  }

  // Reorder managed questions within their slots, preserving unrelated items.
  const orderedItems = activeColumns
    .map(column => itemsById.get(mappings.get(column.itemId)))
    .filter(item => item !== undefined);
  const managedIds = new Set(orderedItems.map(item => String(item.getId())));
  let managedIndex = 0;
  const desiredOrder = form.getItems().map(item => {
    if (managedIds.has(String(item.getId()))) {
      return orderedItems[managedIndex++];
    }
    return item;
  });

  for (let index = 0; index < desiredOrder.length; index++) {
    const item = desiredOrder[index];
    if (item.getIndex() !== index) {
      form.moveItem(item, index);
    }
  }

  properties.setProperty(initializedKey, 'true');
}
