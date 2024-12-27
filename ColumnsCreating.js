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
      alertAndLog('Error: "Ambassador Id" column not foundin Overall Score Sheet.');
      return;
    }

    // Ensure "Ambassador Status" column is in overall score, add to the end if it is not found
    let statusColumnIndex = getColumnIndexByName(overallScoreSheet, 'Ambassador Status');
    if (statusColumnIndex === 0) {
      statusColumnIndex = overallScoreSheet.getLastColumn() + 1;
      overallScoreSheet.getRange(1, statusColumnIndex).setValue('Ambassador Status');
      Logger.log(`Created "Ambassador Status" column at index ${statusColumnIndex}.`);
    }

    // Sync "Ambassadors' Discord Handles" and "Ambassador Id",
    // loop through the ambassador registry list, skipping first row (header)
    // check that id matches current email hash or update if not
    // update the overall score sheet with the new hash if needed
    // ensure every ambassador from the registry can be found in the overall score sheet, or add them
    // TODO Suggestion: change to use named columns
    for (let i = 1; i < registryData.length; i++) {
      let ambassadorId = registryData[i][0]; // expects ambassador is is first registry column
      // TODO Decide: what to do wtih empty registry data - "" all generate same hash.
      const newHash = generateMD5Hash(registryData[i][1]); // expects email is second registry column
      const discordHandle = registryData[i][2]; // expects discord handle is third registry column
      const registryAmbassadorStatus = registryData[i][3]; // expects status is fourth registry column

      // if the ambassador id is not the hash of the email, we need to update the ambassador IDs
      if (ambassadorId !== newHash) {
        try {
          // Update the old overall score row with the new hash, matching first by ambassadorid, then discord handle
          // assumes only one row per id or discord handle in the overall score sheet
          // if we don't find the ambassador by discord handle or id, we should add a new row
          let existingRow = null;
          if (ambassadorId) {
            // Find the old overall score row, matching first by ambassador id
            existingRow = overallScoreSheet.createTextFinder(ambassadorId).findNext()?.getRow();
          }
          if (!existingRow && discordHandle) {
            // If old hash not found, check for a score row by ambassador discord handle
            existingRow = overallScoreSheet.createTextFinder(discordHandle).findNext()?.getRow();
          }
          if (existingRow) {
            overallScoreSheet.getRange(existingRow, ambassadorIdColumnIndex).setValue(newHash);
          } else {
            // if not found by id or by discord handle, we should add a new row in the overall sheet
            const newRowIndex = overallScoreSheet.getLastRow() + 1;
            overallScoreSheet.getRange(newRowIndex, ambassadorIdColumnIndex).setValue(newHash);
            overallScoreSheet.getRange(newRowIndex, discordHandleColumnIndex).setValue(discordHandle);
          }
          // Overwrite old ambassador id with new hash
          ambassadorId = newHash;
          registrySheet.getRange(i + 1, ambassadorIdColumnIndex).setValue(ambassadorId);
        } catch (error) {
          alertAndLog('Error in updating ambassador id in overall score and registry', error);
          alertAndLog('Ambassador Id may be in unknown state:', ambassadorId);
          return;
        }
      } else {
        // If the ambassador id is already the hash of the email, we should still check if the id and discord handle is in the overall score sheet
        // if the discord handle is not in the overall score sheet, we should add it
        // first check by ambassador id, then by discord handle, as above.
        existingRow = overallScoreSheet.createTextFinder(ambassadorId).findNext()?.getRow();
        if (!existingRow) {
          existingRow = overallScoreSheet.createTextFinder(discordHandle).findNext()?.getRow();
        }
        if (!existingRow) {
          // if the discord handle is not in the overall score sheet, we should add it
          const newRowIndex = overallScoreSheet.getLastRow() + 1;
          overallScoreSheet.getRange(newRowIndex, discordHandleColumnIndex).setValue(discordHandle);
          overallScoreSheet.getRange(newRowIndex, ambassadorIdColumnIndex).setValue(ambassadorId);
        } else {
          // if the discord handle is in the overall sheet but the id is not, update the id for that discord handle
          overallScoreSheet.getRange(existingRow, ambassadorIdColumnIndex).setValue(ambassadorId);
        }
      }
      // Now that we know they are all there, let's update their status!
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