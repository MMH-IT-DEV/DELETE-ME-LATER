function migrateLogAD() {
  var message =
    'This migration targeted the retired Log-based tracker and is disabled.\n\n' +
    'Use setupSheet() to initialize the 3-tab system workbook instead.';

  try {
    SpreadsheetApp.getUi().alert('Legacy Migration Disabled', message, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (err) {
    Logger.log(message);
  }
}
