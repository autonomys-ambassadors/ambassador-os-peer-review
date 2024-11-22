// MODULE 3

function runComplianceAudit() {
  // Run evaluation window check and exit if the user presses "Cancel"
  if (!checkEvaluationWindowStart()) {
    Logger.log('runComplianceAudit process stopped by user.');
    return;
  }
  // Step 1: Check and create Penalty Points and Max 6-Month PP columns, if they do not exist
  checkAndCreateColumns();
  SpreadsheetApp.flush();
  // Copying all Final Score values to month column in Overall score.
  // Note: Even if Evaluations came late, they anyway helpful for accountability, while those evaluators are fined.
  copyFinalScoresToOverallScore();
  SpreadsheetApp.flush();
  // Step 4: [⚠️DESIGNED TO RUN ONLY ONCE] - Calculates penalty points for past months violations, colors events, adds PP to PP column.
  detectNonRespondersPastMonths();
  SpreadsheetApp.flush();
  // Step 2: Calculate penalty points for missing Submissions for the current month
  calculatePenaltyPointsForSubmissions();
  SpreadsheetApp.flush();
  // Step 3: Calculate penalty points for missing Evaluations for the current month
  calculatePenaltyPointsForEvaluations();
  SpreadsheetApp.flush();
  // Step 5: Calculate the maximum number of penalty points for any continuous 6-month period
  calculateMaxPenaltyPointsForSixMonths();
  SpreadsheetApp.flush();
  // Step 6: Check for ambassadors eligible for expulsion
  expelAmbassadors();
  SpreadsheetApp.flush();
  // Step 7: Send expulsion notifications
  sendExpulsionNotifications();
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

    const spreadsheetTimeZone = scoresSpreadsheet.getSpreadsheetTimeZone();
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

  const spreadsheetTimeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
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
              cell.setBackground(COLOR_OLD_MISSED_SUBMISSION);
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
 * Calculates and adds penalty points for failing to participate in Submission in current reporting month,
 * Highlights cells with light tone.
 */
function calculatePenaltyPointsForSubmissions() {
  Logger.log('Starting penalty points calculation for missing submissions.');

  const registrySheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(REGISTRY_SHEET_NAME);
  const formResponsesSheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(
    FORM_RESPONSES_SHEET_NAME
  );
  const overallScoresSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OVERALL_SCORE_SHEET_NAME);

  const headersRange = overallScoresSheet.getRange(1, 1, 1, overallScoresSheet.getLastColumn()).getValues()[0];
  const penaltyPointsColIndex = headersRange.indexOf('Penalty Points') + 1;
  const reportingMonthDate = getPreviousMonthDate(SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone()); // Получаем предыдущий месяц
  const reportingMonthName = Utilities.formatDate(
    reportingMonthDate,
    SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(),
    'MMMM yyyy'
  );
  const monthColumnIndex =
    headersRange.findIndex(
      (header) =>
        header instanceof Date &&
        Utilities.formatDate(header, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'MMMM yyyy') ===
          reportingMonthName
    ) + 1;

  if (!monthColumnIndex) {
    Logger.log(`Month column "${reportingMonthName}" not found.`);
    return;
  }

  const submissionWindowStart = new Date(PropertiesService.getScriptProperties().getProperty('submissionWindowStart'));
  const submissionWindowEnd = new Date(submissionWindowStart);
  submissionWindowEnd.setMinutes(submissionWindowStart.getMinutes() + SUBMISSION_WINDOW_MINUTES);
  Logger.log(`Submission window is from ${submissionWindowStart} to ${submissionWindowEnd}`);
  const submittedEmails = new Set(
    formResponsesSheet
      .getRange(2, 1, formResponsesSheet.getLastRow() - 1, 2)
      .getValues()
      .filter((row) => {
        const timestamp = new Date(row[0]);
        return timestamp >= submissionWindowStart && timestamp <= submissionWindowEnd;
      })
      .map((row) => row[1].trim().toLowerCase())
  );

  registrySheet
    .getRange(2, 1, registrySheet.getLastRow() - 1, 1)
    .getValues()
    .forEach((row, index) => {
      const ambassadorEmail = row[0].trim().toLowerCase();

      if (!submittedEmails.has(ambassadorEmail)) {
        const discordHandle = getDiscordHandleFromEmail(ambassadorEmail); // From SharedUtilities
        const rowInScores = overallScoresSheet.createTextFinder(discordHandle).findNext()?.getRow();

        if (rowInScores) {
          const currentPenaltyPoints = overallScoresSheet.getRange(rowInScores, penaltyPointsColIndex).getValue() || 0;
          overallScoresSheet.getRange(rowInScores, penaltyPointsColIndex).setValue(currentPenaltyPoints + 1);
          overallScoresSheet
            .getRange(rowInScores, monthColumnIndex)
            .setBackground(COLOR_MISSED_SUBMISSION)
            .setValue('');
          Logger.log(`Added 1 penalty point for missing submission for ${discordHandle}.`);
        }
      }
    });

  Logger.log('Penalty points calculation for missing submissions completed.');
}

/**
 * Calculates and adds penalty points for failing to participate in Evaluation in current reporting month,
 * only those who were assigned as evaluators, but did not submit evaluations within the valid timeframe are penalized.
 * coloring cells with middle tone or dark tone if already colored (if both violations occur).
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

    const validEvaluators = getValidEvaluationEmails(evaluationResponsesSheet); // From SharedUtilities
    const assignments = getReviewLogAssignments(); // From SharedUtilities

    const headersRange = overallScoresSheet.getRange(1, 1, 1, overallScoresSheet.getLastColumn()).getValues()[0];
    const penaltyPointsColIndex = headersRange.indexOf('Penalty Points') + 1;

    if (penaltyPointsColIndex === 0) {
      Logger.log('Error: Penalty Points column not found.');
      return;
    }

    const spreadsheetTimeZone = overallScoresSpreadsheet.getSpreadsheetTimeZone();
    const deliverableMonthDate = getPreviousMonthDate(spreadsheetTimeZone);
    Logger.log(`Previous month date: ${deliverableMonthDate} (ISO: ${deliverableMonthDate.toISOString()})`);

    const monthColumnIndex =
      headersRange.findIndex(
        (header) => header instanceof Date && header.getTime() === deliverableMonthDate.getTime()
      ) + 1;

    if (!monthColumnIndex) {
      Logger.log(`Month column "${deliverableMonthDate}" not found.`);
      return;
    }

    const sheetData = overallScoresSheet
      .getRange(2, 1, overallScoresSheet.getLastRow() - 1, overallScoresSheet.getLastColumn())
      .getValues();
    const backgroundData = overallScoresSheet
      .getRange(2, 1, overallScoresSheet.getLastRow() - 1, overallScoresSheet.getLastColumn())
      .getBackgrounds();

    // Store evaluators who missed all evaluations
    const missedEvaluators = new Set();

    // Loop through assignments and identify evaluators who missed ALL assigned evaluations
    Object.keys(assignments).forEach((submitter) => {
      const evaluators = assignments[submitter];

      evaluators.forEach((evaluator) => {
        // If evaluator didn't submit any evaluations, penalize them
        if (!validEvaluators.includes(evaluator)) {
          const discordHandle = getDiscordHandleFromEmail(evaluator);
          const evaluatorRow = findRowByDiscordHandle(discordHandle);

          if (evaluatorRow && !missedEvaluators.has(discordHandle)) {
            missedEvaluators.add(discordHandle); // Mark this evaluator as already penalized for missing all evaluations

            // Update penalty points
            let currentPenaltyPoints = sheetData[evaluatorRow - 2][penaltyPointsColIndex - 1] || 0;
            sheetData[evaluatorRow - 2][penaltyPointsColIndex - 1] = currentPenaltyPoints + 1;

            // Update background color to reflect the penalty
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

    // Apply penalty points and color updates
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
  const spreadsheetTimeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();

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

      Logger.log(
        `Checking 6-month period starting from column ${monthColumns[i]} (${Utilities.formatDate(headers[monthColumns[i] - 1], spreadsheetTimeZone, 'MMMM yyyy')})`
      );

      for (let j = i; j < i + 6; j++) {
        const cellBackgroundColor = backgroundColors[j].toLowerCase();

        // Log the actual background color detected for each cell
        Logger.log(`Row ${row}, Column ${monthColumns[j]}: Detected background color = ${cellBackgroundColor}`);

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
            Logger.log(`No penalty for background color at column ${monthColumns[j]}`);
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
// Modified to concatenate 'Expelled [DD MMM YY]' to the current status.
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
        // Avoid processing headers by ensuring registryRowIndex is > 1
        // Send expulsion notifications first
        sendExpulsionNotifications(discordHandle);

        const currentStatus = registrySheet.getRange(registryRowIndex, statusColIndex).getValue();
        // Check if 'Expelled' is already in the status
        if (currentStatus.includes('Expelled')) {
          Logger.log(`Error: This ambassador's status is already 'Expelled': "${currentStatus}"`);
          return;
        }

        const currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yy');

        // Concatenate 'Expelled [DD MMM YY]' to the current status
        const updatedStatus = `${currentStatus} Expelled [${currentDate}].`;
        registrySheet.getRange(registryRowIndex, statusColIndex).setValue(updatedStatus);

        Logger.log(`Ambassador ${discordHandle} status updated to: "${updatedStatus}"`);
      }
    }
  });
}

/**
 * Sends expulsion notifications to the expelled ambassador and sponsor.
 * @param {string} discordHandle - The ambassador's discord handle to look up the original or expelled email.
 */
function sendExpulsionNotifications(discordHandle) {
  Logger.log(`Sending expulsion notifications for ambassador with discord handle: ${discordHandle}`);

  const registrySheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(REGISTRY_SHEET_NAME);
  const registryHeaders = registrySheet.getRange(1, 1, 1, registrySheet.getLastColumn()).getValues()[0];
  const emailColIndex = registryHeaders.indexOf(AMBASSADOR_EMAIL_COLUMN) + 1;

  // Find ambassador's row by discord handle
  const registryRowIndex =
    registrySheet
      .getRange(2, 2, registrySheet.getLastRow() - 1, 1)
      .getValues()
      .findIndex((row) => row[0] === discordHandle) + 2;

  if (registryRowIndex) {
    let email = registrySheet.getRange(registryRowIndex, emailColIndex).getValue();

    // Check if the email is in expelled format and modify accordingly
    const expelledEmail = email.startsWith('(EXPELLED) ') ? email : `(EXPELLED) ${email}`;

    const subject = 'Expulsion from the Program';
    const body = EXPULSION_EMAIL_TEMPLATE.replace('{AmbassadorEmail}', expelledEmail);
    const sponsorBody = `Ambassador ${expelledEmail} has been expelled from the program.`;

    // Send notification to the expelled ambassador
    sendEmailNotification(expelledEmail, subject, body);
    Logger.log(`Expulsion email sent to ${expelledEmail}.`);

    // Send notification to the sponsor
    sendEmailNotification(SPONSOR_EMAIL, subject, sponsorBody);
    Logger.log(`Notification sent to sponsor for expelled ambassador: ${expelledEmail}.`);
  } else {
    Logger.log(`Error: Ambassador with discord handle ${discordHandle} not found in the registry.`);
  }
}
