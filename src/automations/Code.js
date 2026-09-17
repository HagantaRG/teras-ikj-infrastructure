function createFormFromSheet() {
  // Get the active sheet
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('stok-barang');
  
  // Get the header row (first row) and all column headers
   const headers = sheet.getRange('C1:' + sheet.getLastColumn() + '1').getValues()[0];
  
  // Create a new form
  const form = FormApp.create('Dynamic Form');
  
  // Loop through each header and create a form question
  headers.forEach(header => {
    // Skip empty headers
    if (header.trim() !== '') {
      Logger.log('Header is:' + header)
      form.addTextItem()
        .setTitle(header)
        .setRequired(true);
    }
  });

  // Log the form URL so you can access it
  Logger.log('Form created: ' + form.getEditUrl());
}
