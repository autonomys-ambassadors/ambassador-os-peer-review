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

  // Let's sync the data to make sure overall score has all ambassadors and knows who has been expelled before now
  syncRegistryColumnsToOverallScore();
  SpreadsheetApp.flush();

  // Copying all Final Score values to month column in Overall score.
  // Note: Even if Evaluations came late, they anyway are helpful, though evaluators are penalized.
  copyFinalScoresToOverallScore();
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

  // Calling the function to sync Ambassador Status columns from Registry back to Overall score, to reflect changes
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
 * Note: Even if Evaluations came late, they anyway helpful for accountability, while those evaluators are subject to fine.
 */
function copyFinalScoresToOverallScore() {
  try {
    Logger.log('Starting copy of Final Scores to Overall Score sheet.');

    // open table "Ambassadors' Scores" and needed sheets
    const scoresSpreadsheet = SpreadsheetApp.openById(AMBASSADORS_SCORES_SPREADSHEET_ID);
    const overallScoreSheet = scoresSpreadsheet.getSheetByName(OVERALL_SCORE_SHEET_NAME);

    if (!overallScoreSheet) {
      alertAndLog(`Sheet "${OVERALL_SCORE_SHEET_NAME}" isn't found.`);
      throw new Error('Overall Score sheet not found.');
    }

    // Use the latest Evaluation request to determine the reporting month
    const reportingMonth = getReportingMonthFromRequestLog('Evaluation');
    if (!reportingMonth) {
      alertAndLog('Error: Could not determine reporting month from Request Log.');
      throw new Error('Reporting month not found.');
    }
    const currentMonthDate = reportingMonth.firstDayDate;
    const spreadsheetTimeZone = getProjectTimeZone(); // Get project time zone
    Logger.log(`Current month date for copying scores: ${currentMonthDate.toISOString()}`);

    const monthSheetName = Utilities.formatDate(currentMonthDate, spreadsheetTimeZone, 'MMMM yyyy');
    const monthSheet = scoresSpreadsheet.getSheetByName(monthSheetName);

    if (!monthSheet) {
      alertAndLog(`Month sheet "${monthSheetName}" not found.`);
      throw new Error('Month sheet not found.');
    }

    // Searching column index in "Overall score" by date
    const existingColumns = overallScoreSheet.getRange(1, 1, 1, overallScoreSheet.getLastColumn()).getValues()[0];
    const monthColumnIndex =
      existingColumns.findIndex((header) => header instanceof Date && header.getTime() === currentMonthDate.getTime()) +
      1;
    const monthDiscordColIndex = getRequiredColumnIndexByName(monthSheet, GRADE_SUBMITTER_COLUMN);
    const monthFinalScoreColIndex = getRequiredColumnIndexByName(monthSheet, GRADE_FINAL_SCORE_COLUMN);

    if (monthColumnIndex === 0) {
      alertAndLog(`Column for "${monthSheetName}" not found in Overall score sheet.`);
      throw new Error('Column for monthly score not found in Overall score sheet.');
    }

    // Fetch data from Final Score on month sheet
    const finalScores = monthSheet
      .getRange(2, 1, monthSheet.getLastRow() - 1, monthSheet.getLastColumn())
      .getValues()
      .map((row) => ({
        handle: row[monthDiscordColIndex - 1],
        score: row[monthFinalScoreColIndex - 1],
      }));

    Logger.log(`Retrieved ${finalScores.length} scores from "${monthSheetName}" sheet.`);

    //Copy Final Score values to proper rows Overall score" by Discord Handles
    const overallSheetDiscordColumn = getRequiredColumnIndexByName(overallScoreSheet, AMBASSADOR_DISCORD_HANDLE_COLUMN);
    const overallHandles = overallScoreSheet
      .getRange(2, overallSheetDiscordColumn, overallScoreSheet.getLastRow() - 1, 1)
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
    alertAndLog(`Error in copyFinalScoresToOverallScore: ${error}`);
    throw error;
  }
}

// Function to check and create "Penalty Points" and "Max 6-Month PP" columns in the correct order
function checkAndCreateColumns() {
  const scoresSpreadsheet = SpreadsheetApp.openById(AMBASSADORS_SCORES_SPREADSHEET_ID);
  const overallScoresSheet = scoresSpreadsheet.getSheetByName(OVERALL_SCORE_SHEET_NAME);

  // Find the index of the "Average Score" column
  const averageScoreColIndex = getColumnIndexByName(overallScoresSheet, SCORE_AVERAGE_SCORE_COLUMN);
  let nextColIndex = averageScoreColIndex; // Start position for the next column

  // Check if "Penalty Points" column exists

  let penaltyPointsColIndex = getColumnIndexByName(overallScoresSheet, SCORE_PENALTY_POINTS_COLUMN);

  if (penaltyPointsColIndex === -1) {
    nextColIndex += 1; // Next column after "Average Score"
    overallScoresSheet.insertColumnAfter(averageScoreColIndex);
    overallScoresSheet.getRange(1, nextColIndex).setValue(SCORE_PENALTY_POINTS_COLUMN);
    Logger.log('Created "Penalty Points" column.');
    penaltyPointsColIndex = nextColIndex; // Update index for the newly created column
  }

  // Check if "Max 6-Month PP" column exists
  let maxPenaltyPointsColIndex = getColumnIndexByName(overallScoresSheet, SCORE_MAX_6M_PP_COLUMN);
  if (maxPenaltyPointsColIndex === -1) {
    nextColIndex = penaltyPointsColIndex + 1; // Next column after "Penalty Points"
    overallScoresSheet.insertColumnAfter(penaltyPointsColIndex);
    overallScoresSheet.getRange(1, nextColIndex).setValue(SCORE_MAX_6M_PP_COLUMN);
    Logger.log('Created "Max 6-Month PP" column.');
    maxPenaltyPointsColIndex = nextColIndex;
  }

  // Check if "Inadequate Contribution Count" column exists
  let inadequateContributionColIndex = getColumnIndexByName(overallScoresSheet, SCORE_INADEQUATE_CONTRIBUTION_COLUMN);
  if (inadequateContributionColIndex === -1) {
    nextColIndex = maxPenaltyPointsColIndex + 1; // Next column after "Max 6-Month PP"
    overallScoresSheet.insertColumnAfter(maxPenaltyPointsColIndex);
    overallScoresSheet.getRange(1, nextColIndex).setValue(SCORE_INADEQUATE_CONTRIBUTION_COLUMN);
    Logger.log('Created "Inadequate Contribution Count" column.');
  }
}

/**
 * Calculates and assigns penalty points for ambassadors based on their participation in submissions and evaluations for the current reporting month.
 * Only considers the last 6 months for penalty points calculation.
 * Highlights the corresponding cells in the Overall Scores sheet to reflect missed activities:
 * - Missed submission: COLOR_MISSED_SUBMISSION
 * - Missed evaluation: COLOR_MISSED_EVALUATION
 * - Both missed submission and evaluation: COLOR_MISSED_SUBM_AND_EVAL
 */
function calculatePenaltyPoints() {
  Logger.log('Starting penalty points calculation for submissions and evaluations (last 6 months only).');

  // Open necessary sheets
  const registrySheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(REGISTRY_SHEET_NAME);
  const scoresSpreadsheet = SpreadsheetApp.openById(AMBASSADORS_SCORES_SPREADSHEET_ID);
  const overallScoresSheet = scoresSpreadsheet.getSheetByName(OVERALL_SCORE_SHEET_NAME);
  const reviewLogSheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(
    REVIEW_LOG_SHEET_NAME
  );
  const evaluationResponsesSheet = SpreadsheetApp.openById(EVALUATION_RESPONSES_SPREADSHEET_ID).getSheetByName(
    EVAL_FORM_RESPONSES_SHEET_NAME
  );
  const submissionsSheet = SpreadsheetApp.openById(AMBASSADORS_SUBMISSIONS_SPREADSHEET_ID).getSheetByName(
    FORM_RESPONSES_SHEET_NAME
  );

  if (!registrySheet || !overallScoresSheet || !reviewLogSheet || !evaluationResponsesSheet) {
    Logger.log('Error: One or more required sheets not found.');
    throw new Error('Overall Score, Review Log, or Evaluation Response sheets are missing');
  }

  // Use the latest Evaluation request to determine the reporting month
  const reportingMonth = getReportingMonthFromRequestLog('Evaluation');
  if (!reportingMonth) {
    alertAndLog('Error: Could not determine reporting month from Request Log.');
    throw new Error('Reporting month not found.');
  }
  const currentReportingMonth = reportingMonth.firstDayDate;

  // Get headers and indices
  const headersRange = overallScoresSheet.getRange(1, 1, 1, overallScoresSheet.getLastColumn()).getValues()[0];
  const penaltyPointsColIndex = getRequiredColumnIndexByName(overallScoresSheet, SCORE_PENALTY_POINTS_COLUMN);
  const currentMonthColIndex =
    headersRange.findIndex((header) => header instanceof Date && header.getTime() === currentReportingMonth.getTime()) +
    1;

  if (currentMonthColIndex === 0) {
    Logger.log('Error: Current reporting month column not found.');
    throw new Error('Current reporting month column not found.');
  }
  Logger.log(`Current reporting month column index: ${currentMonthColIndex}`);

  // Get valid submitters and evaluators
  const validSubmitters = getValidSubmissionEmails(submissionsSheet);
  const validEvaluators = getValidEvaluationEmails(evaluationResponsesSheet);
  Logger.log(`Valid submitters: ${validSubmitters.join(', ')}`);
  Logger.log(`Valid evaluators: ${validEvaluators.join(', ')}`);

  // Get assignments from Review Log (who evaluates whom)
  const assignments = getReviewLogAssignments();
  //Logger.log(`Assignments from Review Log: ${JSON.stringify(assignments)}`); //too extensive log

  // Fetch data from Registry and filter non-expelled ambassadors
  const registryData = registrySheet
    .getRange(2, 1, registrySheet.getLastRow() - 1, registrySheet.getLastColumn())
    .getValues();
  const registryEmailColumn = getRequiredColumnIndexByName(registrySheet, AMBASSADOR_EMAIL_COLUMN) - 1;
  const registryDiscordColumn = getRequiredColumnIndexByName(registrySheet, AMBASSADOR_DISCORD_HANDLE_COLUMN) - 1;
  const registryStatusColumn = getRequiredColumnIndexByName(registrySheet, AMBASSADOR_STATUS_COLUMN) - 1;

  const ambassadorData = registryData
    .filter((row) => row[registryEmailColumn]?.trim() && !row[registryStatusColumn]?.toLowerCase().includes('expelled'))
    .map((row) => ({
      email: row[registryEmailColumn].trim().toLowerCase(),
      discordHandle: row[registryDiscordColumn]?.trim().toLowerCase(),
    }));

  Logger.log(`Filtered ambassadors: ${ambassadorData.length} valid rows`);

  // Collect month column indices in chronological order
  const monthColumns = [];
  for (let i = 0; i < headersRange.length; i++) {
    if (headersRange[i] instanceof Date) {
      monthColumns.push(i + 1); // Add 1 to convert from 0-based to 1-based indexing
      Logger.log(`Found valid month column at index ${i + 1} with date: ${headersRange[i].toISOString()}`);
    } else {
      Logger.log(`Invalid date value: ${headersRange[i]} found in column ${i + 1}. Skipping this column.`);
    }
  }

  // Sort month columns chronologically
  monthColumns.sort((a, b) => {
    const dateA = headersRange[a - 1];
    const dateB = headersRange[b - 1];
    return dateA - dateB;
  });

  Logger.log(`Found ${monthColumns.length} month columns`);

  // Get recent 6 months (or fewer if not enough months available)
  const recentMonths = monthColumns.slice(-Math.min(6, monthColumns.length));
  Logger.log(`Using ${recentMonths.length} most recent months for penalty calculation`);

  const inadequateContributionColIndex = getRequiredColumnIndexByName(
    overallScoresSheet,
    SCORE_INADEQUATE_CONTRIBUTION_COLUMN
  );

  // Process each ambassador
  ambassadorData.forEach(({ email, discordHandle }) => {
    const rowInScores = overallScoresSheet.createTextFinder(discordHandle).findNext()?.getRow();
    if (!rowInScores) {
      alertAndLog(`Discord handle not found in Overall Scores: ${discordHandle}`);
      return;
    }

    // Reset penalty points and recalculate based only on the last 6 months
    let totalPenaltyPoints = 0;
    let inadequateContributionCount = 0;

    // Process recent months (up to 6)
    for (const colIndex of recentMonths) {
      // if processing current month, first color-code the nonEvaluators and non-submitters
      if (colIndex === currentMonthColIndex) {
        Logger.log(`Processing current month column (${colIndex}) for ${discordHandle}`);
        const isNonSubmitter = !validSubmitters.includes(email);
        const isNonEvaluator = Object.values(assignments).some(
          (evaluators) => evaluators.includes(email) && !validEvaluators.includes(email)
        );

        const currentCell = overallScoresSheet.getRange(rowInScores, currentMonthColIndex);

        if (isNonSubmitter && isNonEvaluator) {
          currentCell.setBackground(COLOR_MISSED_SUBM_AND_EVAL);
          Logger.log(`Added 2 penalty points for ${discordHandle} (missed submission and evaluation).`);
        } else if (isNonSubmitter) {
          currentCell.setBackground(COLOR_MISSED_SUBMISSION);
          Logger.log(`Added 1 penalty point for ${discordHandle} (missed submission).`);
        } else if (isNonEvaluator) {
          currentCell.setBackground(COLOR_MISSED_EVALUATION);
          Logger.log(`Added 1 penalty point for ${discordHandle} (missed evaluation).`);
        }
      }

      const cell = overallScoresSheet.getRange(rowInScores, colIndex);
      const backgroundColor = cell.getBackground().toLowerCase();

      if (backgroundColor === COLOR_MISSED_SUBMISSION.toLowerCase()) {
        totalPenaltyPoints += 1;
      } else if (backgroundColor === COLOR_MISSED_EVALUATION.toLowerCase()) {
        totalPenaltyPoints += 1;
      } else if (backgroundColor === COLOR_MISSED_SUBM_AND_EVAL.toLowerCase()) {
        totalPenaltyPoints += 2;
      }

      // Inadequate Contribution: check if Final Score < INADEQUATE_CONTRIBUTION_SCORE_THRESHOLD
      // Get the value from the cell (should be the score for that month)
      const scoreValue = cell.getValue();
      if (typeof scoreValue === 'number' && scoreValue < INADEQUATE_CONTRIBUTION_SCORE_THRESHOLD) {
        inadequateContributionCount++;
      }
    }

    // Update penalty points
    overallScoresSheet.getRange(rowInScores, penaltyPointsColIndex).setValue(totalPenaltyPoints);
    Logger.log(`Updated penalty points for ${discordHandle} to ${totalPenaltyPoints}`);

    // Update Inadequate Contribution Count
    overallScoresSheet.getRange(rowInScores, inadequateContributionColIndex).setValue(inadequateContributionCount);
    Logger.log(`Updated Inadequate Contribution Count for ${discordHandle} to ${inadequateContributionCount}`);

    // Refer to CRT if threshold met
    if (inadequateContributionCount >= MAX_INADEQUATE_CONTRIBUTION_COUNT_TO_REFER) {
      referInadequateContributionToCRT(discordHandle, inadequateContributionCount);
    }
  });

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

  const scoresSpreadsheet = SpreadsheetApp.openById(AMBASSADORS_SCORES_SPREADSHEET_ID);
  const overallScoresSheet = scoresSpreadsheet.getSheetByName(OVERALL_SCORE_SHEET_NAME);
  const penaltyPointsCol = getRequiredColumnIndexByName(overallScoresSheet, SCORE_PENALTY_POINTS_COLUMN);
  const maxPPCol = getRequiredColumnIndexByName(overallScoresSheet, SCORE_MAX_6M_PP_COLUMN);
  const lastRow = overallScoresSheet.getLastRow();
  const lastColumn = overallScoresSheet.getLastColumn();
  const spreadsheetTimeZone = getProjectTimeZone();

  // Collect indices of all month columns
  const headers = overallScoresSheet.getRange(1, 1, 1, overallScoresSheet.getLastColumn()).getValues()[0];
  const monthColumns = [];
  for (let col = 1; col <= lastColumn; col++) {
    const cellValue = headers[col - 1];
    if (cellValue instanceof Date) {
      monthColumns.push(col);
      Logger.log(
        `Found month column at index ${col} with date: ${Utilities.formatDate(
          cellValue,
          spreadsheetTimeZone,
          'MMMM yyyy'
        )}`
      );
    }
  }

  if (monthColumns.length === 0) {
    Logger.log('No month columns found. Exiting calculation.');
    throw new Error('No month columns found in the Overall Scores sheet.');
  }

  // Set the period length to the minimum of 6 or available months
  const periodLength = Math.min(6, monthColumns.length);
  Logger.log(`Period length for calculation: ${periodLength} months.`);

  // Process each row
  for (let row = 2; row <= lastRow; row++) {
    let maxPP = 0;
    Logger.log(`\nCalculating Max 6-Month Penalty Points for row ${row}.`);

    // Collect background colors for month columns
    const backgroundColors = overallScoresSheet
      .getRange(row, monthColumns[0], 1, monthColumns.length)
      .getBackgrounds()[0];

    Logger.log(`Row ${row}: Collected background colors for all month columns.`);

    // Iterate over all possible periods
    for (let i = 0; i <= monthColumns.length - periodLength; i++) {
      let periodTotal = 0;

      for (let j = i; j < i + periodLength; j++) {
        const cellBackgroundColor = backgroundColors[j].toLowerCase();

        switch (cellBackgroundColor) {
          case COLOR_MISSED_SUBMISSION:
            periodTotal += 1;
            Logger.log(`Row ${row}: Adding 1 point for missed submission at column ${monthColumns[j]}.`);
            break;
          case COLOR_MISSED_EVALUATION:
            periodTotal += 1;
            Logger.log(`Row ${row}: Adding 1 point for missed evaluation at column ${monthColumns[j]}.`);
            break;
          case COLOR_MISSED_SUBM_AND_EVAL:
            periodTotal += 2;
            Logger.log(
              `Row ${row}: Adding 2 points for missed submission and evaluation at column ${monthColumns[j]}.`
            );
            break;
          default:
            // No penalty for other colors
            break;
        }
      }

      maxPP = Math.max(maxPP, periodTotal); // Update maxPP if current total is greater
    }

    // Write the maximum penalty points to the "Max 6-Month PP" column
    const maxPPCell = overallScoresSheet.getRange(row, maxPPCol);
    maxPPCell.setValue(maxPP);

    // Center-align the data in the cell
    maxPPCell.setHorizontalAlignment('center');

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
// If an ambassador already has status containing "Expelled", they are assumed to have already been notified and will be skipped.
function expelAmbassadors() {
  Logger.log('Starting expelAmbassadors process.');

  const registrySheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(REGISTRY_SHEET_NAME);
  const scoresSpreadsheet = SpreadsheetApp.openById(AMBASSADORS_SCORES_SPREADSHEET_ID);
  const overallScoresSheet = scoresSpreadsheet.getSheetByName(OVERALL_SCORE_SHEET_NAME);
  const scorePenaltiesColIndex = getRequiredColumnIndexByName(overallScoresSheet, SCORE_PENALTY_POINTS_COLUMN);
  const scoreDiscordHandleColIndex = getRequiredColumnIndexByName(overallScoresSheet, AMBASSADOR_DISCORD_HANDLE_COLUMN);
  const registryDiscordHandleColIndex = getRequiredColumnIndexByName(registrySheet, AMBASSADOR_DISCORD_HANDLE_COLUMN);
  const registryStatusColIndex = getRequiredColumnIndexByName(registrySheet, AMBASSADOR_STATUS_COLUMN);

  const newlyExpelled = [];

  const scoreData = overallScoresSheet
    .getRange(2, 1, overallScoresSheet.getLastRow() - 1, overallScoresSheet.getLastColumn()) // Correct range
    .getValues()
    .filter((row) => row[scorePenaltiesColIndex - 1] >= MAX_PENALTY_POINTS_TO_EXPEL); // -1 for array index

  scoreData.forEach((row) => {
    const discordHandle = row[scoreDiscordHandleColIndex - 1]; // -1 for array index
    const registryRowIndex =
      registrySheet
        .getRange(2, registryDiscordHandleColIndex, registrySheet.getLastRow() - 1, 1) // Correct range
        .getValues()
        .findIndex((regRow) => regRow[0] === discordHandle) + 2; // +2 to adjust for headers and 0-based index

    if (registryRowIndex > 1) {
      const currentStatus = registrySheet.getRange(registryRowIndex, registryStatusColIndex).getValue();
      if (currentStatus.includes('Expelled')) {
        Logger.log(`Notice: Ambassador ${discordHandle} is already marked as expelled.`);
        return;
      }

      const currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yy');
      const updatedStatus = `${currentStatus} | Expelled [${currentDate}].`;
      registrySheet.getRange(registryRowIndex, registryStatusColIndex).setValue(updatedStatus);

      Logger.log(`Ambassador ${discordHandle} status updated to: "${updatedStatus}"`);
      newlyExpelled.push(discordHandle);
    }
  });

  newlyExpelled.forEach((discordHandle) => {
    sendExpulsionNotifications(discordHandle);
  });
  Logger.log('expelAmbassadors process completed.');
}

/**
 * Sends expulsion notifications to the expelled ambassador and sponsor.
 * @param {string} discordHandle - The ambassador's discord handle to look up the email.
 */
function sendExpulsionNotifications(discordHandle) {
  Logger.log(`Sending expulsion notifications for ambassador with discord handle: ${discordHandle}`);

  const registrySheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(REGISTRY_SHEET_NAME);
  const registryEmailColIndex = getRequiredColumnIndexByName(registrySheet, AMBASSADOR_EMAIL_COLUMN);
  const registryDiscordColIndex = getRequiredColumnIndexByName(registrySheet, AMBASSADOR_DISCORD_HANDLE_COLUMN);

  // Find ambassador's row by discord handle
  const registryRowIndex =
    registrySheet
      .getRange(2, registryDiscordColIndex, registrySheet.getLastRow() - 1, 1)
      .getValues()
      .findIndex((row) => row[0] === discordHandle) + 2;

  if (registryRowIndex > 1) {
    const email = registrySheet.getRange(registryRowIndex, registryEmailColIndex).getValue();

    // Check and skip if email is missing
    if (!email || !email.trim()) {
      Logger.log(`Skipping notification for ambassador with discord handle: ${discordHandle} due to missing email.`);
      return;
    }

    const subject = 'Expulsion from the Program';
    const sponsorBody = `Ambassador ${email} (${discordHandle}) has been expelled from the program for Failure to Participate according to Article 2, Section 10 of the Bylaws.`;

    // Send notification to the expelled ambassador using generic email function
    sendEmailNotification(email, subject, EXPULSION_EMAIL_TEMPLATE);
    Logger.log(`Expulsion email sent to ${email}.`);

    // Send notification to the sponsor using generic email function
    sendEmailNotification(SPONSOR_EMAIL, subject, sponsorBody);
    Logger.log(`Notification sent to sponsor for expelled ambassador: ${email} (${discordHandle}).`);
  } else {
    Logger.log(`Error: Ambassador with discord handle: ${discordHandle} not found in the registry.`);
  }
}

/**
 * Refers an ambassador to the CRT for inadequate contribution.
 * Sends a notification email directly to the subject ambassador, copying the sponsor.
 * @param {string} discordHandle - The Discord handle of the ambassador being referred.
 * @param {number} inadequateContributionCount - The number of times the ambassador scored below the inadequate contribution threshold in the last 6 months.
 */
function referInadequateContributionToCRT(discordHandle, inadequateContributionCount) {
  try {
    // Get ambassador's email and discord handle using utility
    const accused = lookupEmailAndDiscord(discordHandle);
    const ambassadorEmail = accused ? accused.email : '';
    const ambassadorDiscord = accused ? accused.discordHandle : discordHandle;

    if (!ambassadorEmail) {
      Logger.log(
        `Error: No email found for ambassador ${ambassadorDiscord}. Cannot send inadequate contribution notification.`
      );
      return;
    }

    // Get the current reporting month name and calculate deadline
    const currentMonthName = getCurrentReportingMonthName();
    const deadlineDate = getBusinessDaysFromToday(3);

    // Compose the email using the new template
    const emailBody = INADEQUATE_CONTRIBUTION_NOTIFICATION_EMAIL_TEMPLATE.replaceAll(
      '{monthName}',
      currentMonthName
    ).replaceAll('{deadlineDate}', deadlineDate);

    const subject = `Autonomys AmbasasadorOS CRT complaint - Inadequate Contribution`;

    // Send to ambassador, CC sponsor
    sendEmailNotification(ambassadorEmail, subject, emailBody, '', SPONSOR_EMAIL);
    Logger.log(
      `Sent inadequate contribution notification to ${ambassadorEmail} (${ambassadorDiscord}), CC: ${SPONSOR_EMAIL}`
    );
  } catch (e) {
    Logger.log('Error in referInadequateContributionToCRT: ' + e);
  }
}
