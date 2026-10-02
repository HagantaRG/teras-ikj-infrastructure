// Requires the given spreadsheet ID property in Project Settings > Script Properties.
function getConfiguredSpreadsheet_(propertyKey) {
  const spreadsheetId = PropertiesService.getScriptProperties()
    .getProperty(propertyKey);
  if (!spreadsheetId) {
    throw new Error(
      'Set ' + propertyKey + ' in Project Settings > Script Properties.'
    );
  }
  return SpreadsheetApp.openById(spreadsheetId);
}
