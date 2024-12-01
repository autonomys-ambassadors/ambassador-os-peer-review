// MODULE 3

function runComplianceAudit() {
  // Run evaluation window check and exit if the user presses "Cancel"
  if (!checkEvaluationWindowStart()) {
    Logger.log('runComplianceAudit process stopped by user.');
    return;
  }
  // Check and create Penalty Points and Max 6-Month PP columns, if they do not exist
  checkAndCreateColumns();
  SpreadsheetApp.flush();
  // Copying all Final Score values to month column in Overall score.
  // Note: Even if Evaluations came late, they anyway are helpful, though evaluators are penalized.
  copyFinalScoresToOverallScore();
  SpreadsheetApp.flush();
  // ⚠️DESIGNED TO RUN ONLY ONCE. Calculates penalty points for past months, colors cells, adds PP to PP column.
  detectNonRespondersPastMonths();
  SpreadsheetApp.flush();
  // Calculate penalty points for missing Submissions and Evaluations for the current reporting month
  calculatePenaltyPoints();
  SpreadsheetApp.flush();
  // Calculate the maximum number of penalty points for any contiguous 6-month period
  calculateMaxPenaltyPointsForSixMonths();
  SpreadsheetApp.flush();
  // Check for ambassadors eligible for expulsion
  expelAmbassadors();
  SpreadsheetApp.flush();
  // Send expulsion notifications
  sendExpulsionNotifications();
  SpreadsheetApp.flush();
  // Calling the function to sync Ambassador Status columns in Overall score with it in Registry, to reflect changes
  syncRegistryColumnsToOverallScore();
  SpreadsheetApp.flush();
  Logger.log('Compliance Audit process completed.');
}

/**
 * Checks if the evaluation window has finished.
 * If not, displays a warning to the user with an "OK, I understand" button.
 * Returns true if the user chooses to proceed, and false if the user presses "Cancel."
 */
function checkEvaluationWindowStart() {
  const { evaluationWindowStart, evaluationWindowEnd } = getEvaluationWindowTimes();
  const currentDate = new Date();

  if (currentDate < evaluationWindowEnd) {
    const response = promptAndLog(
      'Warning',
      'This module should be run after the evaluation window has ended in the current cycle. \n\nClick CANCEL to wait, or OK to proceed anyway.',
      ButtonSet.OK_CANCEL
    );

    // Return true to proceed if the user clicked "OK"; return false to exit
    if (response == ButtonResponse.OK) {
      Logger.log('User acknowledged the warning and chose to proceed.');
      return true;
    } else {
      Logger.log('User chose to exit after warning.');
      return false;
    }
  } else {
    Logger.log('Current date is after the the evaluation window ends; no warning needed.');
    return true; // No warning needed, proceed with the audit
  }
}

/**
 * Copies Final Score from the current month sheet to the current month column in the Overall score sheet.
 * Note: Even if Evaluations came late, they anyway helpful for accountability, while those evaluators are fined.
 */
function copyFinalScoresToOverallScore() {
  try {
    Logger.log('Starting copy of Final Scores to Overall Score sheet.');

    // open table "Ambassadors' Scores" and needed sheets
    const scoresSpreadsheet = SpreadsheetApp.openById(AMBASSADORS_SCORES_SPREADSHEET_ID);
    const overallScoreSheet = scoresSpreadsheet.getSheetByName(OVERALL_SCORE_SHEET_NAME);

    if (!overallScoreSheet) {
      Logger.log(`Sheet "${OVERALL_SCORE_SHEET_NAME}" isn't found.`);
      return;
    }

    const spreadsheetTimeZone = getProjectTimeZone(); // Get project time zone
    const currentMonthDate = getPreviousMonthDate(spreadsheetTimeZone); // getting reporting month
    Logger.log(`Current month date for copying scores: ${currentMonthDate.toISOString()}`);

    const monthSheetName = Utilities.formatDate(currentMonthDate, spreadsheetTimeZone, 'MMMM yyyy');
    const monthSheet = scoresSpreadsheet.getSheetByName(monthSheetName);

    if (!monthSheet) {
      Logger.log(`Month sheet "${monthSheetName}" not found.`);
      return;
    }

    // Searching column index in "Overall score" by date
    const existingColumns = overallScoreSheet.getRange(1, 1, 1, overallScoreSheet.getLastColumn()).getValues()[0];
    const monthColumnIndex =
      existingColumns.findIndex((header) => header instanceof Date && header.getTime() === currentMonthDate.getTime()) +
      1;

    if (monthColumnIndex === 0) {
      Logger.log(`Column for "${monthSheetName}" not found in Overall score sheet.`);
      return;
    }

    // Fetch data from Final Score on month sheet
    const finalScores = monthSheet
      .getRange(2, 1, monthSheet.getLastRow() - 1, 11)
      .getValues()
      .map((row) => ({ handle: row[0], score: row[10] }));

    Logger.log(`Retrieved ${finalScores.length} scores from "${monthSheetName}" sheet.`);

    //Copy Final Score values to proper rows Overall score" by Discord Handles
    const overallHandles = overallScoreSheet
      .getRange(2, 1, overallScoreSheet.getLastRow() - 1, 1)
      .getValues()
      .flat();
    finalScores.forEach(({ handle, score }) => {
      const rowIndex = overallHandles.findIndex((overallHandle) => overallHandle === handle) + 2;
      if (rowIndex > 1 && score !== '') {
        overallScoreSheet.getRange(rowIndex, monthColumnIndex).setValue(score);
        Logger.log(`Copied score for handle ${handle} to row ${rowIndex} in Overall score sheet.`);
      }
    });

    Logger.log('Copy of Final Scores to Overall Score sheet completed.');
  } catch (error) {
    Logger.log(`Error in copyFinalScoresToOverallScore: ${error}`);
  }
}

// Function to check and create "Penalty Points" and "Max 6-Month PP" columns in the correct order
function checkAndCreateColumns() {
  const overallScoresSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OVERALL_SCORE_SHEET_NAME);
  const headersRange = overallScoresSheet.getRange(1, 1, 1, overallScoresSheet.getLastColumn()).getValues()[0];

  // Find the index of the "Average Score" column
  const averageScoreColIndex = headersRange.indexOf('Average Score') + 1;
  if (averageScoreColIndex === 0) {
    Logger.log('Error: "Average Score" column not found.');
    return;
  }

  let nextColIndex = averageScoreColIndex; // Start position for the next column

  // Check if "Penalty Points" column exists
  let penaltyPointsColIndex = headersRange.indexOf('Penalty Points') + 1;
  if (penaltyPointsColIndex === 0) {
    nextColIndex += 1; // Next column after "Average Score"
    overallScoresSheet.insertColumnAfter(averageScoreColIndex);
    overallScoresSheet.getRange(1, nextColIndex).setValue('Penalty Points');
    Logger.log('Created "Penalty Points" column.');
    penaltyPointsColIndex = nextColIndex; // Update index for the newly created column
  }

  // Refresh headers to account for the newly added column
  const updatedHeadersRange = overallScoresSheet.getRange(1, 1, 1, overallScoresSheet.getLastColumn()).getValues()[0];

  // Check if "Max 6-Month PP" column exists
  let maxPenaltyPointsColIndex = updatedHeadersRange.indexOf('Max 6-Month PP') + 1;
  if (maxPenaltyPointsColIndex === 0) {
    nextColIndex = penaltyPointsColIndex + 1; // Next column after "Penalty Points"
    overallScoresSheet.insertColumnAfter(penaltyPointsColIndex);
    overallScoresSheet.getRange(1, nextColIndex).setValue('Max 6-Month PP');
    Logger.log('Created "Max 6-Month PP" column.');
  }
}

// ⚠️ One time run only! Detect non-responders for past months (highlighting with COLOR_OLD_MISSED_SUBMISSION)
function detectNonRespondersPastMonths() {
  const scriptProperties = PropertiesService.getScriptProperties();
  let hasRun = scriptProperties.getProperty('detectNonRespondersPastMonthsRan');

  // Check if the function has already been executed
  if (hasRun === 'true') {
    Logger.log('Warning: This function has already been executed and is locked from repeated runs.');

    // Show warning to the user in the UI if trying to run again
    alertAndLog(
      "Warning! Processing Past Months function is designed to run only once. To allow a re-run, set 'detectNonRespondersPastMonthsRan' to 'false' in the script properties."
    );
    return; // Terminate execution if function already ran once
  }

  Logger.log('Executing detectNonRespondersPastMonths for the first time.');

  const overallScoresSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OVERALL_SCORE_SHEET_NAME);
  const headersRange = overallScoresSheet.getRange(1, 1, 1, overallScoresSheet.getLastColumn()).getValues()[0];
  const penaltyPointsColIndex = headersRange.indexOf('Penalty Points') + 1;

  if (penaltyPointsColIndex === 0 || penaltyPointsColIndex === -1) {
    Logger.log('Error: Penalty Points column not found.');
    return;
  }

  const spreadsheetTimeZone = getProjectTimeZone(); // Get project time zone
  const currentMonthDate = getPreviousMonthDate(spreadsheetTimeZone);
  const currentMonthName = Utilities.formatDate(currentMonthDate, spreadsheetTimeZone, 'MMMM yyyy');
  Logger.log(`Current reporting month: ${currentMonthName}`);

  const lastRow = overallScoresSheet.getLastRow();
  const lastColumn = overallScoresSheet.getLastColumn();
  const sheetData = overallScoresSheet.getRange(1, 1, lastRow, lastColumn).getValues();

  for (let row = 2; row <= lastRow; row++) {
    let currentPenaltyPoints = sheetData[row - 1][penaltyPointsColIndex - 1] || 0;

    for (let col = 1; col <= lastColumn; col++) {
      const cellValue = sheetData[0][col - 1];

      if (cellValue instanceof Date) {
        const cellMonthName = Utilities.formatDate(cellValue, spreadsheetTimeZone, 'MMMM yyyy');

        if (cellMonthName === currentMonthName) continue;

        const pastMonthValue = sheetData[row - 1][col - 1];

        if (typeof pastMonthValue === 'string') {
          const markers = pastMonthValue.split(';').map((s) => s.trim());
          markers.forEach((marker) => {
            if (marker === "didn't submit" || marker === 'late submission') {
              currentPenaltyPoints += 1;
              const cell = overallScoresSheet.getRange(row, col);
              cell.setBackground(COLOR_MISSED_SUBMISSION); // initially (COLOR_OLD_MISSED_SUBMISSION) wes here, but it makes no sense to separate them. Lerr error prone.
            }
          });
        }
      }
    }

    sheetData[row - 1][penaltyPointsColIndex - 1] = currentPenaltyPoints;
  }

  overallScoresSheet
    .getRange(2, penaltyPointsColIndex, lastRow - 1, 1)
    .setValues(sheetData.slice(1).map((row) => [row[penaltyPointsColIndex - 1]]));

  // Set the property to 'true' to lock the function from running again
  scriptProperties.setProperty('detectNonRespondersPastMonthsRan', 'true');
  Logger.log('detectNonRespondersPastMonths completed. Function locked from repeated runs.');
}

/**
 * Calculates and assigns penalty points for ambassadors based on their participation in submissions and evaluations for the current reporting month.
 * Highlights the corresponding cells in the Overall Scores sheet to reflect missed activities:
 * - Missed submission: COLOR_MISSED_SUBMISSION
 * - Missed evaluation: COLOR_MISSED_EVALUATION
 * - Both missed submission and evaluation: COLOR_MISSED_SUBM_AND_EVAL
 * Adds 1 penalty point for each missed activity or 2 points for both missed activities.
 */
function calculatePenaltyPoints() {
  Logger.log('Starting penalty points calculation for submissions and evaluations.');

  // Open necessary sheets
  const registrySheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(REGISTRY_SHEET_NAME);
  const overallScoresSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OVERALL_SCORE_SHEET_NAME);
  const reviewLogSheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(
    REVIEW_LOG_SHEET_NAME
  );
  const evaluationResponsesSheet = SpreadsheetApp.openById(EVALUATION_RESPONSES_SPREADSHEET_ID).getSheetByName(
    EVAL_FORM_RESPONSES_SHEET_NAME
  );

  if (!registrySheet || !overallScoresSheet || !reviewLogSheet || !evaluationResponsesSheet) {
    Logger.log('Error: One or more required sheets not found.');
    return;
  }

  // Get headers and indices
  const headersRange = overallScoresSheet.getRange(1, 1, 1, overallScoresSheet.getLastColumn()).getValues()[0];
  const penaltyPointsColIndex = headersRange.indexOf('Penalty Points') + 1;
  if (penaltyPointsColIndex === 0) {
    Logger.log('Error: Penalty Points column not found.');
    return;
  }

  const spreadsheetTimeZone = getProjectTimeZone();
  const currentReportingMonth = getPreviousMonthDate(spreadsheetTimeZone); // Assume previous month is the reporting month
  const currentMonthColIndex =
    headersRange.findIndex((header) => header instanceof Date && header.getTime() === currentReportingMonth.getTime()) +
    1;
  if (currentMonthColIndex === 0) {
    Logger.log('Error: Current reporting month column not found.');
    return;
  }
  Logger.log(`Current reporting month column index: ${currentMonthColIndex}`);

  // Get valid submitters and evaluators
  const validSubmitters = getValidSubmissionEmails();
  const validEvaluators = getValidEvaluationEmails(evaluationResponsesSheet);
  Logger.log(`Valid submitters: ${validSubmitters.join(', ')}`);
  Logger.log(`Valid evaluators: ${validEvaluators.join(', ')}`);

  // Get assignments from Review Log (who evaluates whom)
  const assignments = getReviewLogAssignments();
  //Logger.log(`Assignments from Review Log: ${JSON.stringify(assignments)}`); //too extensive log

  // Fetch data from Registry and filter non-expelled ambassadors
  const registryData = registrySheet.getRange(2, 1, registrySheet.getLastRow() - 1, 3).getValues();
  const ambassadorData = registryData
    .filter((row) => row[0]?.trim() && !row[2]?.includes('Expelled'))
    .map((row) => ({
      email: row[0].trim().toLowerCase(),
      discordHandle: row[1]?.trim(),
    }));

  Logger.log(`Filtered ambassadors: ${ambassadorData.length} valid rows`);

  // Retrieve penalty points
  const penaltyPoints = overallScoresSheet
    .getRange(2, penaltyPointsColIndex, overallScoresSheet.getLastRow() - 1, 1)
    .getValues()
    .flat();

  // Process each ambassador
  ambassadorData.forEach(({ email, discordHandle }) => {
    const rowInScores = overallScoresSheet.createTextFinder(discordHandle).findNext()?.getRow();
    if (!rowInScores) {
      Logger.log(`Discord handle not found in Overall Scores: ${discordHandle}`);
      return;
    }

    const rowIndex = rowInScores - 2; // Adjust for header offset
    const cell = overallScoresSheet.getRange(rowInScores, currentMonthColIndex);
    const currentPenaltyPoints = penaltyPoints[rowIndex] || 0;

    // Determine non-submitter status
    const isNonSubmitter = !validSubmitters.includes(email);

    // Determine non-evaluator status
    const isNonEvaluator = Object.values(assignments).some(
      (evaluators) => evaluators.includes(email) && !validEvaluators.includes(email)
    );

    // Update colors and penalty points based on detected violations
    if (isNonSubmitter && isNonEvaluator) {
      cell.setBackground(COLOR_MISSED_SUBM_AND_EVAL).setValue('');
      penaltyPoints[rowIndex] = currentPenaltyPoints + 2;
      Logger.log(`Added 2 penalty points for ${discordHandle} (missed submission and evaluation).`);
    } else if (isNonSubmitter) {
      cell.setBackground(COLOR_MISSED_SUBMISSION).setValue('');
      penaltyPoints[rowIndex] = currentPenaltyPoints + 1;
      Logger.log(`Added 1 penalty point for ${discordHandle} (missed submission).`);
    } else if (isNonEvaluator) {
      cell.setBackground(COLOR_MISSED_EVALUATION).setValue('');
      penaltyPoints[rowIndex] = currentPenaltyPoints + 1;
      Logger.log(`Added 1 penalty point for ${discordHandle} (missed evaluation).`);
    }
  });

  // Update penalty points column
  overallScoresSheet
    .getRange(2, penaltyPointsColIndex, penaltyPoints.length, 1)
    .setValues(penaltyPoints.map((val) => [val]));

  Logger.log('Penalty points calculation for submissions and evaluations completed.');
}

/**
 * This function calculates the maximum number of penalty points for any full 6-month contiguous period for each ambassador,
 * and records this value in the "Max 6-Month PP" column.
 *
 * - Light tone for old months' "Didn't submit" events - adds 1 penalty point
 * - Light tone for "Didn't submit" events - adds 1 penalty point
 * - Middle tone for "Didn't evaluate" events - adds 1 point
 * - Dark tone for "Didn't submit"+"didn't evaluate" events - adds 2 penalty points
 */
function calculateMaxPenaltyPointsForSixMonths() {
  Logger.log('Starting calculation of Max 6-Month Penalty Points.');

  const overallScoresSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OVERALL_SCORE_SHEET_NAME);
  const headers = overallScoresSheet.getRange(1, 1, 1, overallScoresSheet.getLastColumn()).getValues()[0];

  const penaltyPointsCol = headers.indexOf('Penalty Points') + 1;
  const maxPPCol = headers.indexOf('Max 6-Month PP') + 1;

  if (penaltyPointsCol === 0 || maxPPCol === 0) {
    Logger.log('Error: Either "Penalty Points" or "Max 6-Month PP" column not found.');
    return;
  }

  const lastRow = overallScoresSheet.getLastRow();
  const lastColumn = overallScoresSheet.getLastColumn();
  const spreadsheetTimeZone = getProjectTimeZone();

  // Collect indices of all month columns
  const monthColumns = [];
  for (let col = 1; col <= lastColumn; col++) {
    const cellValue = headers[col - 1];
    if (cellValue instanceof Date) {
      monthColumns.push(col);
      Logger.log(
        `Found month column at index ${col} with date: ${Utilities.formatDate(cellValue, spreadsheetTimeZone, 'MMMM yyyy')}`
      );
    }
  }

  // Process each row
  for (let row = 2; row <= lastRow; row++) {
    let maxPP = 0;
    Logger.log(`\nCalculating Max 6-Month Penalty Points for row ${row}.`);

    // Collect background colors for month columns
    const backgroundColors = overallScoresSheet
      .getRange(row, monthColumns[0], 1, monthColumns.length)
      .getBackgrounds()[0];

    Logger.log(`Row ${row}: Collected background colors for all month columns.`);

    // Iterate over all possible full 6-month periods
    for (let i = 0; i <= monthColumns.length - 6; i++) {
      // Only consider full 6-month periods
      let sixMonthTotal = 0;

      for (let j = i; j < i + 6; j++) {
        // Exactly 6 months
        const cellBackgroundColor = backgroundColors[j].toLowerCase();

        switch (cellBackgroundColor) {
          case COLOR_OLD_MISSED_SUBMISSION:
          case COLOR_MISSED_SUBMISSION:
            sixMonthTotal += 1;
            Logger.log(`Row ${row}: Adding 1 point for missed submission at column ${monthColumns[j]}.`);
            break;
          case COLOR_MISSED_EVALUATION:
            sixMonthTotal += 1;
            Logger.log(`Row ${row}: Adding 1 point for missed evaluation at column ${monthColumns[j]}.`);
            break;
          case COLOR_MISSED_SUBM_AND_EVAL:
            sixMonthTotal += 2;
            Logger.log(
              `Row ${row}: Adding 2 points for missed submission and evaluation at column ${monthColumns[j]}.`
            );
            break;
          default:
            // No penalty for other colors
            break;
        }
      }

      // Logger.log(`Row ${row}: Total penalty points for this 6-month period: ${sixMonthTotal}.`);
      maxPP = Math.max(maxPP, sixMonthTotal); //  Update maxPP if current total is greater
    }

    // Write the maximum penalty points to the "Max 6-Month PP" column
    const maxPPCell = overallScoresSheet.getRange(row, maxPPCol);
    maxPPCell.setValue(maxPP);

    // Set cell background to red if maxPP >= 3
    if (maxPP >= 3) {
      maxPPCell.setBackground(COLOR_EXPELLED);
      Logger.log(`Row ${row}: Max PP >= 3. Setting red background in Max 6-Month PP column.`);
    }

    Logger.log(`Row ${row}: Max 6-Month Penalty Points finalized as ${maxPP}.`);
  }

  Logger.log('Completed calculating Max Penalty Points for all rows.');
}

// Expel ambassadors based on Max 6-Month PP. Tracks and notify only newly expelled ambassadors.
function expelAmbassadors() {
  Logger.log('Starting expelAmbassadors process.');

  const registrySheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(REGISTRY_SHEET_NAME);
  const overallScoresSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OVERALL_SCORE_SHEET_NAME);

  // Get headers for column indices
  const headers = overallScoresSheet.getRange(1, 1, 1, overallScoresSheet.getLastColumn()).getValues()[0];
  const maxPenaltyPointsIndex = headers.indexOf('Max 6-Month PP') + 1;

  const registryHeaders = registrySheet.getRange(1, 1, 1, registrySheet.getLastColumn()).getValues()[0];
  Logger.log(`Registry Headers: ${registryHeaders.join(', ')}`); // Log headers to verify

  const emailColIndex = registryHeaders.indexOf(AMBASSADOR_EMAIL_COLUMN) + 1;
  const discordHandleColIndex = registryHeaders.indexOf(AMBASSADOR_DISCORD_HANDLE_COLUMN) + 1;
  const statusColIndex = registryHeaders.indexOf(AMBASSADOR_STATUS_COLUMN) + 1;

  if (emailColIndex === 0 || discordHandleColIndex === 0 || statusColIndex === 0) {
    Logger.log(
      `Error: Column '${AMBASSADOR_EMAIL_COLUMN}', '${AMBASSADOR_DISCORD_HANDLE_COLUMN}', or '${AMBASSADOR_STATUS_COLUMN}' not found in registry headers.`
    );
    return;
  }

  const newlyExpelled = []; // List to track newly expelled ambassadors

  // Retrieve ambassador data from the Overall Scores sheet
  const ambassadorData = overallScoresSheet
    .getRange(2, 1, overallScoresSheet.getLastRow() - 1, overallScoresSheet.getLastColumn())
    .getValues();

  ambassadorData.forEach((row, i) => {
    const discordHandle = row[0];
    const maxPenaltyPoints = row[maxPenaltyPointsIndex - 1];

    if (maxPenaltyPoints >= 3) {
      const registryRowIndex =
        registrySheet
          .getRange(2, discordHandleColIndex, registrySheet.getLastRow() - 1, 1)
          .getValues()
          .findIndex((regRow) => regRow[0] === discordHandle) + 2; // Adding 2 to account for header row

      if (registryRowIndex > 1) {
        const currentStatus = registrySheet.getRange(registryRowIndex, statusColIndex).getValue();
        // Check if 'Expelled' is already in the status
        if (currentStatus.includes('Expelled')) {
          Logger.log(`Notice: Ambassador ${discordHandle} is already marked as expelled.`);
          return;
        }

        const currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yy');

        // Concatenate 'Expelled [DD MMM YY]' to the current status
        const updatedStatus = `${currentStatus} Expelled [${currentDate}].`;
        registrySheet.getRange(registryRowIndex, statusColIndex).setValue(updatedStatus);

        Logger.log(`Ambassador ${discordHandle} status updated to: "${updatedStatus}"`);
        newlyExpelled.push(discordHandle); // Add to newly expelled list
      }
    }
  });

  // Notify newly expelled ambassadors
  newlyExpelled.forEach((discordHandle) => {
    sendExpulsionNotifications(discordHandle);
  });
}

/**
 * Sends expulsion notifications to the expelled ambassador and sponsor.
 * @param {string} discordHandle - The ambassador's discord handle to look up the email.
 */
function sendExpulsionNotifications(discordHandle) {
  Logger.log(`Sending expulsion notifications for ambassador with discord handle: ${discordHandle}`);

  const registrySheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(REGISTRY_SHEET_NAME);
  const registryHeaders = registrySheet.getRange(1, 1, 1, registrySheet.getLastColumn()).getValues()[0];
  const emailColIndex = registryHeaders.indexOf(AMBASSADOR_EMAIL_COLUMN) + 1;
  const statusColIndex = registryHeaders.indexOf(AMBASSADOR_STATUS_COLUMN) + 1;

  if (emailColIndex === 0 || statusColIndex === 0) {
    Logger.log(
      `Error: Column '${AMBASSADOR_EMAIL_COLUMN}' or '${AMBASSADOR_STATUS_COLUMN}' not found in registry headers.`
    );
    return;
  }

  // Find ambassador's row by discord handle
  const registryRowIndex =
    registrySheet
      .getRange(2, 2, registrySheet.getLastRow() - 1, 1)
      .getValues()
      .findIndex((row) => row[0] === discordHandle) + 2;

  if (registryRowIndex > 1) {
    const email = registrySheet.getRange(registryRowIndex, emailColIndex).getValue();

    // Check and skip if email is missing
    if (!email || !email.trim()) {
      Logger.log(`Skipping notification for ambassador with discord handle: ${discordHandle} due to missing email.`);
      return;
    }

    const subject = 'Expulsion from the Program';
    const body = EXPULSION_EMAIL_TEMPLATE.replace('{AmbassadorEmail}', email);
    const sponsorBody = `Ambassador ${email} (${discordHandle}) has been expelled from the program.`;

    // Send notification to the expelled ambassador using generic email function
    sendEmailNotification(email, subject, body);
    Logger.log(`Expulsion email sent to ${email}.`);

    // Send notification to the sponsor using generic email function
    sendEmailNotification(SPONSOR_EMAIL, subject, sponsorBody);
    Logger.log(`Notification sent to sponsor for expelled ambassador: ${email} (${discordHandle}).`);
  } else {
    Logger.log(`Error: Ambassador with discord handle: ${discordHandle} not found in the registry.`);
  }
}
