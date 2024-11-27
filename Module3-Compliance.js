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
  // Note: Even if Evaluations came late, they anyway helpful for accountability, while those evaluators are fined.
  copyFinalScoresToOverallScore();
  SpreadsheetApp.flush();
  // [⚠️DESIGNED TO RUN ONLY ONCE] - Calculates penalty points for past months violations, colors events, adds PP to PP column.
  detectNonRespondersPastMonths();
  SpreadsheetApp.flush();
  // Calculate penalty points for missing Submissions for the current month
  calculatePenaltyPointsForSubmissions();
  SpreadsheetApp.flush();
  // Calculate penalty points for missing Evaluations for the current month
  calculatePenaltyPointsForEvaluations();
  SpreadsheetApp.flush();
  // Calculate the maximum number of penalty points for any continuous 6-month period
  calculateMaxPenaltyPointsForSixMonths();
  SpreadsheetApp.flush();
  // Check for ambassadors eligible for expulsion
  expelAmbassadors();
  SpreadsheetApp.flush();
  // Send expulsion notifications
  sendExpulsionNotifications();
  SpreadsheetApp.flush();
  // Calling function to sync Ambassador Status columns in Overall score and Registry to display updated statuses
  syncRegistryColumnsToOverallScore();
  SpreadsheetApp.flush();
  Logger.log('Compliance Audit process completed.');
}

/**
 * Checks if 7 days have passed since the start of the evaluation window.
 * If not, displays a warning to the user with an "OK, I understand" button.
 * Returns true if the user chooses to proceed, and false if the user presses "Cancel."
 */
function checkEvaluationWindowStart() {
  const startDateProperty = PropertiesService.getScriptProperties().getProperty('evaluationWindowStart');

  if (!startDateProperty) {
    Logger.log('Error: Evaluation window start date is not set.');
    SpreadsheetApp.getUi().alert('Error: Evaluation window start date is not set.');
    return false;
  }

  const startDate = new Date(startDateProperty);
  const currentDate = new Date();

  // Calculate the difference in days between the current date and the start date
  const daysSinceStart = (currentDate - startDate) / (1000 * 60 * 60 * 24);

  if (daysSinceStart < 7) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Warning',
      'This module should be run 7 days after evaluations are requested in the current cycle. \n\nClick CANCEL to wait, or OK to proceed anyway.',
      ui.ButtonSet.OK_CANCEL
    );

    // Return true to proceed if the user clicked "OK"; return false to exit
    if (response == ui.Button.OK) {
      Logger.log('User acknowledged the warning and chose to proceed.');
      return true;
    } else {
      Logger.log('User chose to exit after warning.');
      return false;
    }
  } else {
    Logger.log('7 days have passed since the evaluation window started; no warning needed.');
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

// Function to check and create Penalty Points and Max 6-Month PP columns
function checkAndCreateColumns() {
  const overallScoresSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OVERALL_SCORE_SHEET_NAME);
  const headersRange = overallScoresSheet.getRange(1, 1, 1, overallScoresSheet.getLastColumn()).getValues()[0];

  let penaltyPointsColIndex = headersRange.indexOf('Penalty Points') + 1;
  if (penaltyPointsColIndex === 0) {
    penaltyPointsColIndex = overallScoresSheet.getLastColumn() + 1;
    overallScoresSheet.getRange(1, penaltyPointsColIndex).setValue('Penalty Points');
    Logger.log('Created "Penalty Points" column.');
  }

  let maxPenaltyPointsColIndex = headersRange.indexOf('Max 6-Month PP') + 1;
  if (maxPenaltyPointsColIndex === 0) {
    maxPenaltyPointsColIndex = overallScoresSheet.getLastColumn() + 1;
    overallScoresSheet.getRange(1, maxPenaltyPointsColIndex).setValue('Max 6-Month PP');
    Logger.log('Created "Max 6-Month PP" column.');
  }
}

// Detect non-responders for past months (highlighting #FFD580) ⚠️ One time run only, otherwase will have to manually deduct duplicated scores
function detectNonRespondersPastMonths() {
  const scriptProperties = PropertiesService.getScriptProperties();
  let hasRun = scriptProperties.getProperty('detectNonRespondersPastMonthsRan');

  // Check if the function has already been executed
  if (hasRun === 'true') {
    Logger.log('Warning: This function has already been executed and is locked from repeated runs.');

    // Show warning to the user in the UI if trying to run again
    SpreadsheetApp.getUi().alert(
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
 * Calculates and adds penalty points for failing to participate in submissions across all months.
 * Highlights cells with light tone for the current reporting month.
 */
function calculatePenaltyPointsForSubmissions() {
  Logger.log('Starting penalty points calculation for missing submissions.');

  // Getting Overall Score sheet
  const overallScoresSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OVERALL_SCORE_SHEET_NAME);
  if (!overallScoresSheet) {
    Logger.log('Error: Overall score sheet not found.');
    return;
  }

  // Getting sheet headers and index of Penalty Point column
  const headersRange = overallScoresSheet.getRange(1, 1, 1, overallScoresSheet.getLastColumn()).getValues()[0];
  const penaltyPointsColIndex = headersRange.indexOf('Penalty Points') + 1;
  if (penaltyPointsColIndex === 0) {
    Logger.log('Error: Penalty Points column not found.');
    return;
  }

  // Reading all rows all columns cells colors
  const backgrounds = overallScoresSheet
    .getRange(2, 1, overallScoresSheet.getLastRow() - 1, overallScoresSheet.getLastColumn())
    .getBackgrounds();

  // Reading penalty points values, save to 1D array
  const penaltyPoints = overallScoresSheet
    .getRange(2, penaltyPointsColIndex, overallScoresSheet.getLastRow() - 1, 1)
    .getValues()
    .flat();

  // Iterate over each month column
  headersRange.forEach((header, colIndex) => {
    if (header instanceof Date) {
      const monthColumnIndex = colIndex + 1;

      // Check for cells with COLOR_MISSED_SUBMISSION
      for (let rowIndex = 0; rowIndex < backgrounds.length; rowIndex++) {
        if (backgrounds[rowIndex][colIndex] === COLOR_MISSED_SUBMISSION) {
          penaltyPoints[rowIndex] = (penaltyPoints[rowIndex] || 0) + 1;
        }
      }
    }
  });

  // Update penalty points and colors
  overallScoresSheet
    .getRange(2, penaltyPointsColIndex, penaltyPoints.length, 1)
    .setValues(penaltyPoints.map((val) => [val]));
  overallScoresSheet
    .getRange(2, 1, backgrounds.length, backgrounds[0].length)
    .setBackgrounds(backgrounds);

  Logger.log('Penalty points calculation for missing submissions completed.');
}


/**
 * Calculates and adds penalty points for failing to participate in evaluations across all months.
 * Highlights cells based on tone: middle(for 1 violation) or dark(for 2 violations) if there were.
 */
function calculatePenaltyPointsForEvaluations() {
  try {
    Logger.log('Starting penalty points calculation for missed evaluations.');

    const reviewLogSheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(
      REVIEW_LOG_SHEET_NAME
    );
    const evaluationResponsesSheet = SpreadsheetApp.openById(AMBASSADORS_SCORES_SPREADSHEET_ID).getSheetByName(
      EVAL_FORM_RESPONSES_SHEET_NAME
    );
    const overallScoresSpreadsheet = SpreadsheetApp.openById(AMBASSADORS_SCORES_SPREADSHEET_ID);
    const overallScoresSheet = overallScoresSpreadsheet.getSheetByName(OVERALL_SCORE_SHEET_NAME);

    const validEvaluators = getValidEvaluationEmails(evaluationResponsesSheet); // getValidEvaluationEmails: returns array of email-addrresses, submitted valid Eval.responses.
    const assignments = getReviewLogAssignments(); // getReviewLogAssignments: returns object with assignments (who evaluates whom)).

    const headersRange = overallScoresSheet.getRange(1, 1, 1, overallScoresSheet.getLastColumn()).getValues()[0]; // getting columns' headers
    const penaltyPointsColIndex = headersRange.indexOf('Penalty Points') + 1; // locates Penalty Point column

    if (penaltyPointsColIndex === 0) {
      Logger.log('Error: Penalty Points column not found.'); //if PP column is not found
      return;
    }

    const spreadsheetTimeZone = getProjectTimeZone(); // Get project time zone
    const deliverableMonthDate = getPreviousMonthDate(spreadsheetTimeZone);
    Logger.log(`Previous month date: ${deliverableMonthDate} (ISO: ${deliverableMonthDate.toISOString()})`);
    
    //Finding current reporting month month-sheet column index
    const monthColumnIndex =
      headersRange.findIndex(
        (header) => header instanceof Date && header.getTime() === deliverableMonthDate.getTime()
      ) + 1;
    // IF couldn't find..
    if (!monthColumnIndex) {
      Logger.log(`Month column "${deliverableMonthDate}" not found.`);
      return;
    }
    //Gets data on every row, skipping headers.
    const sheetData = overallScoresSheet
      .getRange(2, 1, overallScoresSheet.getLastRow() - 1, overallScoresSheet.getLastColumn())
      .getValues();
    //Gets backgrounds on every row, skipping headers
    const backgroundData = overallScoresSheet
      .getRange(2, 1, overallScoresSheet.getLastRow() - 1, overallScoresSheet.getLastColumn())
      .getBackgrounds();


    const missedEvaluators = new Set(); // Creates a set to hold Discord handles of those who missed to evaluate all assigned submitters

    /** Loop through assignments and identify evaluators who missed to evaluate all assigned submitters. 
    * For each submitter from assignments(Review Matrix) gets a list of assigned evaluator.
    * For each evaluator in review log checks, if his email is in valid evaluator (within Eval.window).
    * If email isn't among valid evals: gets Discord-handle using getDiscordHandleFromEmail, 
    * finds this evaluator's row (using findRowByDiscordHandle), and if row is found, adds his Discord-handle into missedEvaluators array.
    * Marks this evaluator as already penalized, Increments PP for this evaluator, set proper color, depending if 1 or 2 violations in total.
    */
    Object.keys(assignments).forEach((submitter) => {
      const evaluators = assignments[submitter];

      evaluators.forEach((evaluator) => {
        // If evaluator didn't submit any evaluations, penalize them
        if (!validEvaluators.includes(evaluator)) {
          const discordHandle = getDiscordHandleFromEmail(evaluator);

          // New check: Skip processing if no email found in Registry
          if (!evaluator) {
            Logger.log(`Skipping evaluator without email: ${discordHandle}`);
            return; // Skip this evaluator
          }

          const evaluatorRow = findRowByDiscordHandle(discordHandle);

          if (evaluatorRow && !missedEvaluators.has(discordHandle)) {
            missedEvaluators.add(discordHandle); // Marks this evaluator as already penalized for missing all evaluations

            // Increments penalty points for evaluators
            let currentPenaltyPoints = sheetData[evaluatorRow - 2][penaltyPointsColIndex - 1] || 0;
            sheetData[evaluatorRow - 2][penaltyPointsColIndex - 1] = currentPenaltyPoints + 1;

            // Update already missed-submission background to both-violations color; and not colored cell to missed_eval color
            const cellColor = backgroundData[evaluatorRow - 2][monthColumnIndex - 1];
            if (cellColor === COLOR_MISSED_SUBMISSION) {
              backgroundData[evaluatorRow - 2][monthColumnIndex - 1] = COLOR_MISSED_SUBM_AND_EVAL; // For both violations
            } else {
              backgroundData[evaluatorRow - 2][monthColumnIndex - 1] = COLOR_MISSED_EVALUATION; // For missed evaluation
            }
          }
        }
      });
    });

    // Writes the updated cell values and colors to Overall score sheet to reflect the penalty
    overallScoresSheet
      .getRange(2, 1, overallScoresSheet.getLastRow() - 1, overallScoresSheet.getLastColumn())
      .setValues(sheetData);
    overallScoresSheet
      .getRange(2, 1, overallScoresSheet.getLastRow() - 1, overallScoresSheet.getLastColumn())
      .setBackgrounds(backgroundData);

    Logger.log('Penalty points calculation for missed evaluations completed.');
  } catch (error) {
    Logger.log(`Error in calculatePenaltyPointsForEvaluations: ${error.message}`);
  }
}


/**
 * This function calculates the maximum number of penalty points for any 6-month contiguous period for each ambassador
 * and records this value in the "Max 6-Month PP" column.
 *
 * - Light tone for old months' "Didn't submit" events - adds 1 penalty point
 * - Light tone for "Didn't submit" events - adds 1 penalty point
 * - Middle tone for "Didn't evaluate" events - adds 1 point
 * - Dark tone for "Didn't submit"+"didn't evaluate" events  - adds 2 penalty points
 */
function calculateMaxPenaltyPointsForSixMonths() {
  const overallScoresSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OVERALL_SCORE_SHEET_NAME);
  const headers = overallScoresSheet.getRange(1, 1, 1, overallScoresSheet.getLastColumn()).getValues()[0];

  const penaltyPointsCol = headers.indexOf('Penalty Points') + 1;
  const maxPPCol = headers.indexOf('Max 6-Month PP') + 1;
  if (penaltyPointsCol === 0 || maxPPCol === 0) {
    Logger.log('Error: Either Penalty Points or Max 6-Month PP column not found.');
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

  for (let row = 2; row <= lastRow; row++) {
    let maxPP = 0;
    Logger.log(`\nCalculating Max 6-Month Penalty Points for row ${row}`);

    const backgroundColors = overallScoresSheet
      .getRange(row, monthColumns[0], 1, monthColumns.length)
      .getBackgrounds()[0];
    Logger.log(`Row ${row}: Collected background colors for all month columns.`);

    // Iterate over possible 6-month periods
    for (let i = 0; i <= monthColumns.length - 6; i++) {
      let sixMonthTotal = 0;

      // Logger.log(
      // `Checking 6-month period starting from column ${monthColumns[i]} (${Utilities.formatDate(headers[monthColumns[i] - 1], spreadsheetTimeZone, 'MMMM yyyy')})`
      // );

      for (let j = i; j < i + 6; j++) {
        const cellBackgroundColor = backgroundColors[j].toLowerCase();

        // Log the actual background color detected for each cell
        ////////Logger.log(`Row ${row}, Column ${monthColumns[j]}: Detected background color = ${cellBackgroundColor}`);

        switch (cellBackgroundColor) {
          case COLOR_OLD_MISSED_SUBMISSION: // for old months' didn't submit
            sixMonthTotal += 1;
            Logger.log(`Adding 1 point for COLOR_OLD_MISSED_SUBMISSION at column ${monthColumns[j]}`);
            break;
          case COLOR_MISSED_SUBMISSION: // didn't submit
            sixMonthTotal += 1;
            Logger.log(`Adding 1 point for COLOR_MISSED_SUBMISSION at column ${monthColumns[j]}`);
            break;
          case COLOR_MISSED_EVALUATION: // didn't evaluate
            sixMonthTotal += 1;
            Logger.log(`Adding 1 point for COLOR_MISSED_EVALUATION at column ${monthColumns[j]}`);
            break;
          case COLOR_MISSED_SUBM_AND_EVAL: // didn't submit and didn't evaluate
            sixMonthTotal += 2;
            Logger.log(`Adding 2 points for COLOR_MISSED_SUBM_AND_EVAL at column ${monthColumns[j]}`);
            break;
          default:
          ////////  Logger.log(`No penalty for background color at column ${monthColumns[j]}`);
        }
      }

      Logger.log(`Total penalty points for this 6-month period: ${sixMonthTotal}`);
      maxPP = Math.max(maxPP, sixMonthTotal);
    }

    // Write the maximum penalty points for any 6-month period to the Max 6-Month PP column
    const maxPPCell = overallScoresSheet.getRange(row, maxPPCol);
    maxPPCell.setValue(maxPP);

    // Color the cell red if maxPP is 3 or greater
    if (maxPP >= 3) {
      maxPPCell.setBackground(COLOR_EXPELLED); // Red color
      Logger.log(`Max PP >= 3. Setting red background for row ${row} in Max 6-Month PP column.`);
    }

    Logger.log(`Row ${row}: Max 6-Month Penalty Points finalized as ${maxPP}`);
  }

  Logger.log('Completed calculating Max Penalty Points for all rows.');
}

// Expel ambassadors based on Max 6-Month PP
// Expel ambassadors based on Max 6-Month PP
// Modified to track and notify only newly expelled ambassadors.
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
      Logger.log(`Skipping notification for ambassador with discord handle ${discordHandle} due to missing email.`);
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
    Logger.log(`Error: Ambassador with discord handle ${discordHandle} not found in the registry.`);
  }
}

