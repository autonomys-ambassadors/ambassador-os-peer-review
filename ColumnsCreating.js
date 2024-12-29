function syncRegistryColumnsToOverallScore() {
  const expectedRegistryHeaders = [
    'Ambassador Id',
    'Ambassador Email Address',
    'Ambassador Discord Handle',
    'Ambassador Status',
  ];

  try {
    Logger.log('Starting synchronization of columns between "Registry" and "Overall score".');

    // Opening Registry and Overall score sheets
    const registrySheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(
      REGISTRY_SHEET_NAME
    );
    const overallScoreSheet = SpreadsheetApp.openById(AMBASSADORS_SCORES_SPREADSHEET_ID).getSheetByName(
      OVERALL_SCORE_SHEET_NAME
    );

    if (!registrySheet || !overallScoreSheet) {
      alertAndLog('Error: One or both sheets not found. Exiting function.');
      return;
    }

    // Getting all data from Registry
    const registryData = registrySheet.getDataRange().getValues(); // Fetch all data

    // Confirm the column headings
    const registryHeaders = registryData[0]; // First row contains the headers
    if (!validateHeaders('Registry', registryHeaders, expectedRegistryHeaders)) {
      alertAndLog('Error: Ambassador Registry sheet headers do not match expected headers ', expectedRegistryHeaders);
      return;
    }

    // Verify columns "Ambassadors' Discord Handles" in overall score
    const discordHandleColumnIndex = getColumnIndexByName(overallScoreSheet, AMBASSADOR_DISCORD_HANDLE_COLUMN);
    if (discordHandleColumnIndex === 0) {
      alertAndLog('Error: "Ambassadors\' Discord Handles" column not found.');
      return;
    }

    // Verify "Ambassador Id" in overall score
    const ambassadorIdColumnIndex = getColumnIndexByName(overallScoreSheet, 'Ambassador Id');
    if (ambassadorIdColumnIndex === 0) {
      alertAndLog('Error: "Ambassador Id" column not found in Overall Score Sheet.');
      return;
    }

    // Ensure "Ambassador Status" column is in overall score, add to the end if it is not found
    let statusColumnIndex = getColumnIndexByName(overallScoreSheet, 'Ambassador Status');
    if (statusColumnIndex === 0) {
      statusColumnIndex = overallScoreSheet.getLastColumn() + 1;
      overallScoreSheet.getRange(1, statusColumnIndex).setValue('Ambassador Status');
      Logger.log(`Created "Ambassador Status" column at index ${statusColumnIndex}.`);
    }

    // Sync "Ambassadors' Discord Handles" and "Ambassador Id"
    for (let i = 1; i < registryData.length; i++) {
      let ambassadorId = registryData[i][0]; // Ambassador Id from registry
      const email = registryData[i][1]?.trim().toLowerCase(); // Ensure email is lowercased and trimmed
      const discordHandle = registryData[i][2]; // Discord Handle
      const registryAmbassadorStatus = registryData[i][3]; // Ambassador Status

      // Ensure email is not empty before generating hash
      if (!email) {
        Logger.log(`Row ${i + 1}: Empty email. Skipping this record.`);
        continue; // Skip processing if email is missing
      }

      const newHash = generateMD5Hash(email); // Generate hash from email
      if (ambassadorId !== newHash) {
        try {
          let existingRow = null;
          if (ambassadorId) {
            existingRow = overallScoreSheet.createTextFinder(ambassadorId).findNext()?.getRow();
          }
          if (!existingRow && discordHandle) {
            existingRow = overallScoreSheet.createTextFinder(discordHandle).findNext()?.getRow();
          }
          if (existingRow) {
            overallScoreSheet.getRange(existingRow, ambassadorIdColumnIndex).setValue(newHash);
          } else {
            const newRowIndex = overallScoreSheet.getLastRow() + 1;
            overallScoreSheet.getRange(newRowIndex, ambassadorIdColumnIndex).setValue(newHash);
            overallScoreSheet.getRange(newRowIndex, discordHandleColumnIndex).setValue(discordHandle);
          }
          ambassadorId = newHash;
          registrySheet.getRange(i + 1, ambassadorIdColumnIndex).setValue(ambassadorId); // Update registry with new hash
        } catch (error) {
          alertAndLog('Error in updating ambassador id in overall score and registry', error);
          alertAndLog('Ambassador Id may be in unknown state:', ambassadorId);
          return;
        }
      } else {
        let existingRow = overallScoreSheet.createTextFinder(ambassadorId).findNext()?.getRow();
        if (!existingRow) {
          existingRow = overallScoreSheet.createTextFinder(discordHandle).findNext()?.getRow();
        }
        if (!existingRow) {
          const newRowIndex = overallScoreSheet.getLastRow() + 1;
          overallScoreSheet.getRange(newRowIndex, discordHandleColumnIndex).setValue(discordHandle);
          overallScoreSheet.getRange(newRowIndex, ambassadorIdColumnIndex).setValue(ambassadorId);
        } else {
          overallScoreSheet.getRange(existingRow, ambassadorIdColumnIndex).setValue(ambassadorId);
        }
      }
      const ambassadorOverallScoreRow = overallScoreSheet.createTextFinder(ambassadorId).findNext()?.getRow();
      if (ambassadorOverallScoreRow) {
        overallScoreSheet.getRange(ambassadorOverallScoreRow, statusColumnIndex).setValue(registryAmbassadorStatus);
      }
    }

    overallScoreSheet
      .getRange(2, discordHandleColumnIndex, overallScoreSheet.getLastRow() - 1)
      .setHorizontalAlignment('left');
    Logger.log('"Ambassadors\' Discord Handles" column synchronized and aligned to the left.');

    overallScoreSheet.getRange(2, statusColumnIndex, overallScoreSheet.getLastRow() - 1).setHorizontalAlignment('left');
    Logger.log('"Ambassador Status" column synchronized and aligned to the left.');

    Logger.log('Synchronization completed successfully.');
  } catch (error) {
    Logger.log(`Error in syncRegistryColumnsToOverallScore: ${error}`);
  }
}
