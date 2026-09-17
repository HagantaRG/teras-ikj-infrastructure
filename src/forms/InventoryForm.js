// Call while holding the script lock.
function syncInventoryForm_(form, sheet, columns, properties, validation) {
  const prefix = [
    'INVENTORY_ITEM',
    sheet.getParent().getId(),
    sheet.getSheetId(),
    form.getId()
  ].join(':') + ':';
  const initializedKey = prefix + 'INITIALIZED';
  const savedProperties = properties.getProperties();
  const initialized = savedProperties[initializedKey] === 'true';
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

    const metadataId = key.slice(prefix.length);
    if (!/^[0-9]+$/.test(metadataId) || !/^-?[0-9]+$/.test(itemId)) {
      throw new Error('Invalid inventory question mapping: ' + key);
    }
    if (claimedItemIds.has(itemId)) {
      throw new Error('Multiple inventory columns reference question ' + itemId);
    }

    const item = itemsById.get(itemId);
    if (item && item.getType() !== FormApp.ItemType.TEXT) {
      throw new Error('Mapped inventory question is no longer a text field: ' + itemId);
    }

    mappings.set(metadataId, itemId);
    claimedItemIds.add(itemId);
  }

  const titleCounts = new Map();
  for (const column of columns) {
    if (column.title !== '') {
      titleCounts.set(column.title, (titleCounts.get(column.title) || 0) + 1);
    }
  }

  // Plan initial adoption before making changes; ambiguous matches need review.
  const plans = [];
  for (const column of columns) {
    const metadataId = String(column.metadataId);
    const mappedItemId = mappings.get(metadataId);
    let item = itemsById.get(mappedItemId);

    if (column.title === '') {
      continue;
    }

    if (!initialized && !mappedItemId) {
      const candidates = items.filter(candidate =>
        candidate.getType() === FormApp.ItemType.TEXT &&
        candidate.getTitle() === column.title &&
        !claimedItemIds.has(String(candidate.getId()))
      );

      if (candidates.length > 1 ||
          (candidates.length === 1 && titleCounts.get(column.title) > 1)) {
        throw new Error('Ambiguous existing question for header: ' + column.title);
      }

      item = candidates[0];
      if (item) {
        claimedItemIds.add(String(item.getId()));
      }
    }

    plans.push({ column, metadataId, item });
  }

  // Persist each association as it is established so retries can reuse it.
  for (const plan of plans) {
    const item = plan.item
      ? plan.item.asTextItem()
      : form.addTextItem().setTitle(plan.column.title);
    const itemId = String(item.getId());

    properties.setProperty(prefix + plan.metadataId, itemId);
    mappings.set(plan.metadataId, itemId);
    itemsById.set(itemId, item);

    item.setTitle(plan.column.title)
      .setRequired(false)
      .setValidation(validation);
  }

  // Blank headers still have metadata, so they are not treated as deletions.
  const liveMetadataIds = new Set(
    columns.map(column => String(column.metadataId))
  );
  for (const [metadataId, itemId] of mappings) {
    if (liveMetadataIds.has(metadataId)) {
      continue;
    }

    const item = itemsById.get(itemId);
    if (item) {
      form.deleteItem(item);
    }
    properties.deleteProperty(prefix + metadataId);
    mappings.delete(metadataId);
  }

  // Reorder managed questions within their slots, preserving unrelated items.
  const orderedItems = columns
    .map(column => itemsById.get(mappings.get(String(column.metadataId))))
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
