function syncRegistryColumnsToOverallScore() {
  const expectedRegistryHeaders = [
    'Number (Unique ID)',
    'Ambassador Id',
    'Ambassador Email Address',
    'Ambassador Discord Handle',
    'Primary Team',
    'Secondary Team',
    'Ambassador Status',
    'Start Date',
  ];

  try {
    Logger.log('Starting synchronization of columns between "Registry" and "Overall score".');

    // Sync with Notion FIRST before any other processing
    try {
      syncRegistryWithNotion();
      Logger.log('Notion sync completed successfully.');
    } catch (error) {
      alertAndLog(`CRITICAL: Notion sync failed - stopping process: ${error.message}`);
      throw error;
    }

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

    const registryAmbassadorIdColumnIndex = getRequiredColumnIndexByName(registrySheet, AMBASSADOR_ID_COLUMN);
    const registryEmailColumnIndex = getRequiredColumnIndexByName(registrySheet, AMBASSADOR_EMAIL_COLUMN);
    const registryDiscordColumnIndex = getRequiredColumnIndexByName(registrySheet, AMBASSADOR_DISCORD_HANDLE_COLUMN);
    const registryStatusColumnIndex = getRequiredColumnIndexByName(registrySheet, AMBASSADOR_STATUS_COLUMN);

    // Verify columns "Ambassadors' Discord Handles" in overall score
    const scoreDiscordHandleColumnIndex = getRequiredColumnIndexByName(
      overallScoreSheet,
      AMBASSADOR_DISCORD_HANDLE_COLUMN
    );
    const scoreAmbassadorIdColumnIndex = getRequiredColumnIndexByName(overallScoreSheet, AMBASSADOR_ID_COLUMN);

    // Ensure "Ambassador Status" column is in overall score, add to the end if it is not found
    let scoreStatusColumnIndex = getColumnIndexByName(overallScoreSheet, AMBASSADOR_STATUS_COLUMN);
    if (scoreStatusColumnIndex === -1) {
      scoreStatusColumnIndex = overallScoreSheet.getLastColumn() + 1;
      overallScoreSheet.getRange(1, scoreStatusColumnIndex).setValue(AMBASSADOR_STATUS_COLUMN);
      Logger.log(`Created "Ambassador Status" column at index ${scoreStatusColumnIndex}.`);
    }

    // Sync "Ambassadors' Discord Handles" and "Ambassador Id"
    for (let i = 1; i < registryData.length; i++) {
      let ambassadorId = registryData[i][registryAmbassadorIdColumnIndex - 1]; // Ambassador Id from registry
      const email = normalizeEmail(registryData[i][registryEmailColumnIndex - 1]); // Ensure email is lowercased and trimmed
      const discordHandle = registryData[i][registryDiscordColumnIndex - 1]; // Discord Handle
      const registryAmbassadorStatus = registryData[i][registryStatusColumnIndex - 1]; // Ambassador Status

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
            overallScoreSheet.getRange(existingRow, scoreAmbassadorIdColumnIndex).setValue(newHash);
          } else {
            const newRowIndex = overallScoreSheet.getLastRow() + 1;
            overallScoreSheet.getRange(newRowIndex, scoreAmbassadorIdColumnIndex).setValue(newHash);
            overallScoreSheet.getRange(newRowIndex, scoreDiscordHandleColumnIndex).setValue(discordHandle);
          }
          ambassadorId = newHash;
          registrySheet.getRange(i + 1, registryAmbassadorIdColumnIndex).setValue(ambassadorId); // Update registry with new hash
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
          overallScoreSheet.getRange(newRowIndex, scoreDiscordHandleColumnIndex).setValue(discordHandle);
          overallScoreSheet.getRange(newRowIndex, scoreAmbassadorIdColumnIndex).setValue(ambassadorId);
        } else {
          overallScoreSheet.getRange(existingRow, scoreAmbassadorIdColumnIndex).setValue(ambassadorId);
        }
      }
      const ambassadorOverallScoreRow = overallScoreSheet.createTextFinder(ambassadorId).findNext()?.getRow();
      if (ambassadorOverallScoreRow) {
        overallScoreSheet
          .getRange(ambassadorOverallScoreRow, scoreStatusColumnIndex)
          .setValue(registryAmbassadorStatus);
      }
    }

    overallScoreSheet
      .getRange(2, scoreDiscordHandleColumnIndex, overallScoreSheet.getLastRow() - 1)
      .setHorizontalAlignment('left');
    Logger.log('"Ambassadors\' Discord Handles" column synchronized and aligned to the left.');

    overallScoreSheet
      .getRange(2, scoreStatusColumnIndex, overallScoreSheet.getLastRow() - 1)
      .setHorizontalAlignment('left');
    Logger.log('"Ambassador Status" column synchronized and aligned to the left.');

    Logger.log('Synchronization completed successfully.');
  } catch (error) {
    Logger.log(`Error in syncRegistryColumnsToOverallScore: ${error}`);
  }
}
