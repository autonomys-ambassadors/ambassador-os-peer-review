function syncRegistryColumnsToOverallScore() {
  try {
    Logger.log('Starting synchronization of columns between "Registry" and "Overall score".');

    // Opening Registry and Overall score sheets
    const registrySheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(REGISTRY_SHEET_NAME);
    const overallScoreSheet = SpreadsheetApp.openById(AMBASSADORS_SCORES_SPREADSHEET_ID).getSheetByName(OVERALL_SCORE_SHEET_NAME);

    if (!registrySheet || !overallScoreSheet) {
      Logger.log('Error: One or both sheets not found. Exiting function.');
      return;
    }

    // Getting data from Registry
    const registryData = registrySheet.getRange(2, 2, registrySheet.getLastRow() - 1, 2).getValues(); // Колонки: Discord Handle, Status

    // Sync "Ambassadors' Discord Handles"
    const discordHandleColumnIndex = getColumnIndexByName(overallScoreSheet, "Ambassadors' Discord Handles");
    if (discordHandleColumnIndex === 0) {
      Logger.log('Error: "Ambassadors\' Discord Handles" column not found.');
      return;
    }
    overallScoreSheet.getRange(2, discordHandleColumnIndex, registryData.length, 1).setValues(
      registryData.map(row => [row[0]]) // Copying Discord Handles
    );
    overallScoreSheet.getRange(2, discordHandleColumnIndex, overallScoreSheet.getLastRow() - 1).setHorizontalAlignment('left');
    Logger.log('"Ambassadors\' Discord Handles" column synchronized and aligned to the left.');

    // Sync "Ambassador Status"
    let statusColumnIndex = getColumnIndexByName(overallScoreSheet, "Ambassador Status");
    if (statusColumnIndex === 0) {
      statusColumnIndex = overallScoreSheet.getLastColumn() + 1;
      overallScoreSheet.getRange(1, statusColumnIndex).setValue("Ambassador Status");
      Logger.log(`Created "Ambassador Status" column at index ${statusColumnIndex}.`);
    }
    overallScoreSheet.getRange(2, statusColumnIndex, registryData.length, 1).setValues(
      registryData.map(row => [row[1]]) // Copying Status
    );
    overallScoreSheet.getRange(2, statusColumnIndex, overallScoreSheet.getLastRow() - 1).setHorizontalAlignment('left');
    Logger.log('"Ambassador Status" column synchronized and aligned to the left.');

    Logger.log('Synchronization completed successfully.');
  } catch (error) {
    Logger.log(`Error in syncRegistryColumnsToOverallScore: ${error}`);
  }
}
