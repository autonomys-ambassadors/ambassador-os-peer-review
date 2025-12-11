// MODULE 3

// Compliance calculation constants
const COMPLIANCE_PERIOD_MONTHS = 6; // Number of months to consider for penalty calculations
const COMPLIANCE_PENALTY_POINT_MISSED_SUBMISSION = 1; // Penalty points for missed submission
const COMPLIANCE_PENALTY_POINT_MISSED_EVALUATION = 1; // Penalty points for missed evaluation
const COMPLIANCE_PENALTY_POINT_MISSED_BOTH = 2; // Penalty points for missing both submission and evaluation
const COMPLIANCE_BUSINESS_DAYS_DEADLINE = 3; // Business days for CRT complaint deadline
const COMPLIANCE_HEADER_ROW = 1; // Row index for headers
const COMPLIANCE_FIRST_DATA_ROW = 2; // Row index for first data row

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

  // Publish anonymous scores to Google Sheet if configured
  try {
    publishAnonymousScoresToGoogleSheet();
  } catch (error) {
    Logger.log(`Error publishing anonymous scores: ${error}`);
  }
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
 * Gets the reporting month data and creates the corresponding month sheet.
 * @returns {Object} Object containing currentMonthDate, monthSheetName, and monthSheet
 */
function getReportingMonthForScoreCopy() {
  // Use the latest Evaluation request to determine the reporting month
  const reportingMonth = getReportingMonthFromRequestLog('Evaluation');
  if (!reportingMonth) {
    const errorMsg = 'Error: Could not determine reporting month from Request Log.';
    alertAndLog(errorMsg);
    throw new Error(errorMsg);
  }
  const currentMonthDate = reportingMonth.firstDayDate;
  const spreadsheetTimeZone = getProjectTimeZone();
  Logger.log(`Current month date for copying scores: ${currentMonthDate.toISOString()}`);

  const monthSheetName = Utilities.formatDate(currentMonthDate, spreadsheetTimeZone, 'MMMM yyyy');
  const scoresSpreadsheet = getScoresSpreadsheet();
  const monthSheet = scoresSpreadsheet.getSheetByName(monthSheetName);

  if (!monthSheet) {
    const errorMsg = `Month sheet "${monthSheetName}" not found.`;
    alertAndLog(errorMsg);
    throw new Error(errorMsg);
  }

  return { currentMonthDate, monthSheetName, monthSheet };
}

/**
 * Gets column indices needed for score copying.
 * @param {Sheet} overallScoreSheet - The Overall Score sheet
 * @param {Sheet} monthSheet - The month sheet
 * @param {Date} currentMonthDate - The current month date
 * @param {string} monthSheetName - The month sheet name
 * @returns {Object} Object containing column indices
 */
function getScoreCopyColumnIndices(overallScoreSheet, monthSheet, currentMonthDate, monthSheetName) {
  // Find column index in "Overall score" by date
  const existingColumns = overallScoreSheet
    .getRange(COMPLIANCE_HEADER_ROW, 1, 1, overallScoreSheet.getLastColumn())
    .getValues()[0];
  const monthColumnIndex =
    existingColumns.findIndex((header) => header instanceof Date && header.getTime() === currentMonthDate.getTime()) +
    1;

  if (monthColumnIndex === 0) {
    const errorMsg = `Column for "${monthSheetName}" not found in Overall score sheet.`;
    alertAndLog(errorMsg);
    throw new Error(errorMsg);
  }

  const monthDiscordColIndex = getRequiredColumnIndexByName(monthSheet, SUBMITTER_HANDLE_COLUMN_IN_MONTHLY_SCORE);
  const monthFinalScoreColIndex = getRequiredColumnIndexByName(monthSheet, GRADE_FINAL_SCORE_COLUMN);

  return { monthColumnIndex, monthDiscordColIndex, monthFinalScoreColIndex };
}

/**
 * Gets final scores from the month sheet.
 * @param {Sheet} monthSheet - The month sheet
 * @param {number} monthDiscordColIndex - Discord column index in month sheet
 * @param {number} monthFinalScoreColIndex - Final score column index in month sheet
 * @param {string} monthSheetName - The month sheet name for logging
 * @returns {Array} Array of score objects with handle and score properties
 */
function getFinalScoresFromMonthSheet(monthSheet, monthDiscordColIndex, monthFinalScoreColIndex, monthSheetName) {
  const finalScores = monthSheet
    .getRange(COMPLIANCE_FIRST_DATA_ROW, 1, monthSheet.getLastRow() - 1, monthSheet.getLastColumn())
    .getValues()
    .map((row) => ({
      handle: row[monthDiscordColIndex - 1],
      score: row[monthFinalScoreColIndex - 1],
    }));

  Logger.log(`Retrieved ${finalScores.length} scores from "${monthSheetName}" sheet.`);
  return finalScores;
}

/**
 * Copies final scores to the overall score sheet by matching discord handles.
 * @param {Array} finalScores - Array of score objects
 * @param {Sheet} overallScoreSheet - The Overall Score sheet
 * @param {number} monthColumnIndex - The month column index in overall score sheet
 */
function copyScoresToOverallSheet(finalScores, overallScoreSheet, monthColumnIndex) {
  const overallSheetDiscordColumn = getRequiredColumnIndexByName(overallScoreSheet, AMBASSADOR_DISCORD_HANDLE_COLUMN);
  const overallHandles = overallScoreSheet
    .getRange(COMPLIANCE_FIRST_DATA_ROW, overallSheetDiscordColumn, overallScoreSheet.getLastRow() - 1, 1)
    .getValues()
    .flat();

  finalScores.forEach(({ handle, score }) => {
    const rowIndex =
      overallHandles.findIndex(
        (overallHandle) => normalizeDiscordHandle(overallHandle) === normalizeDiscordHandle(handle)
      ) + 2;
    if (rowIndex > 1 && score !== '') {
      overallScoreSheet.getRange(rowIndex, monthColumnIndex).setValue(score);
      Logger.log(`Copied score for handle ${handle} to row ${rowIndex} in Overall score sheet.`);
    }
  });
}

/**
 * Copies Final Score from the current month sheet to the current month column in the Overall score sheet.
 * Note: Even if Evaluations came late, they anyway helpful for accountability, while those evaluators are subject to fine.
 */
function copyFinalScoresToOverallScore() {
  try {
    Logger.log('Starting copy of Final Scores to Overall Score sheet.');

    // Get overall score sheet
    const overallScoreSheet = getOverallScoreSheet();
    if (!overallScoreSheet) {
      const errorMsg = `Sheet "${OVERALL_SCORE_SHEET_NAME}" not found.`;
      alertAndLog(errorMsg);
      throw new Error(errorMsg);
    }

    // Get reporting month data and month sheet
    const { currentMonthDate, monthSheetName, monthSheet } = getReportingMonthForScoreCopy();

    // Get column indices for score copying
    const { monthColumnIndex, monthDiscordColIndex, monthFinalScoreColIndex } = getScoreCopyColumnIndices(
      overallScoreSheet,
      monthSheet,
      currentMonthDate,
      monthSheetName
    );

    // Extract final scores from month sheet
    const finalScores = getFinalScoresFromMonthSheet(
      monthSheet,
      monthDiscordColIndex,
      monthFinalScoreColIndex,
      monthSheetName
    );

    // Copy scores to overall sheet
    copyScoresToOverallSheet(finalScores, overallScoreSheet, monthColumnIndex);

    Logger.log('Copy of Final Scores to Overall Score sheet completed.');
  } catch (error) {
    alertAndLog(`Error in copyFinalScoresToOverallScore: ${error}`);
    throw error;
  }
}

// Function to check and create "Penalty Points" and "Max 6-Month PP" columns in the correct order
function checkAndCreateColumns() {
  const scoresSpreadsheet = getScoresSpreadsheet();
  const overallScoresSheet = scoresSpreadsheet.getSheetByName(OVERALL_SCORE_SHEET_NAME);

  // Find the index of the "Average Score" column
  const averageScoreColIndex = getColumnIndexByName(overallScoresSheet, SCORE_AVERAGE_SCORE_COLUMN);
  let nextColIndex = averageScoreColIndex; // Start position for the next column

  // Check if "Penalty Points" column exists

  let penaltyPointsColIndex = getColumnIndexByName(overallScoresSheet, SCORE_PENALTY_POINTS_COLUMN);

  if (penaltyPointsColIndex === -1) {
    nextColIndex += 1; // Next column after "Average Score"
    overallScoresSheet.insertColumnAfter(averageScoreColIndex);
    overallScoresSheet.getRange(COMPLIANCE_HEADER_ROW, nextColIndex).setValue(SCORE_PENALTY_POINTS_COLUMN);
    Logger.log('Created "Penalty Points" column.');
    penaltyPointsColIndex = nextColIndex; // Update index for the newly created column
  }

  // Check if "Max 6-Month PP" column exists
  let maxPenaltyPointsColIndex = getColumnIndexByName(overallScoresSheet, SCORE_MAX_6M_PP_COLUMN);
  if (maxPenaltyPointsColIndex === -1) {
    nextColIndex = penaltyPointsColIndex + 1; // Next column after "Penalty Points"
    overallScoresSheet.insertColumnAfter(penaltyPointsColIndex);
    overallScoresSheet.getRange(COMPLIANCE_HEADER_ROW, nextColIndex).setValue(SCORE_MAX_6M_PP_COLUMN);
    Logger.log('Created "Max 6-Month PP" column.');
    maxPenaltyPointsColIndex = nextColIndex;
  }

  // Check if "Inadequate Contribution Count" column exists
  let inadequateContributionColIndex = getColumnIndexByName(overallScoresSheet, SCORE_INADEQUATE_CONTRIBUTION_COLUMN);
  if (inadequateContributionColIndex === -1) {
    nextColIndex = maxPenaltyPointsColIndex + 1; // Next column after "Max 6-Month PP"
    overallScoresSheet.insertColumnAfter(maxPenaltyPointsColIndex);
    overallScoresSheet.getRange(COMPLIANCE_HEADER_ROW, nextColIndex).setValue(SCORE_INADEQUATE_CONTRIBUTION_COLUMN);
    Logger.log('Created "Inadequate Contribution Count" column.');
    inadequateContributionColIndex = nextColIndex;
  }

  // Check if "CRT Referral History" column exists
  let crtReferralHistoryColIndex = getColumnIndexByName(overallScoresSheet, SCORE_CRT_REFERRAL_HISTORY_COLUMN);
  if (crtReferralHistoryColIndex === -1) {
    nextColIndex = inadequateContributionColIndex + 1; // Next column after "Inadequate Contribution Count"
    overallScoresSheet.insertColumnAfter(inadequateContributionColIndex);
    overallScoresSheet.getRange(COMPLIANCE_HEADER_ROW, nextColIndex).setValue(SCORE_CRT_REFERRAL_HISTORY_COLUMN);
    Logger.log('Created "CRT Referral History" column.');
  }
}

/**
 * Initializes all data needed for penalty points calculation.
 * @returns {Object} Object containing sheets, column indices, reporting data, and ambassador data
 */
function initializePenaltyCalculationData() {
  Logger.log('Initializing penalty calculation data.');

  // Open necessary sheets
  const registrySheet = getRegistrySheet();
  const scoresSpreadsheet = getScoresSpreadsheet();
  const overallScoresSheet = scoresSpreadsheet.getSheetByName(OVERALL_SCORE_SHEET_NAME);
  const reviewLogSheet = getReviewLogSheet();
  const evaluationResponsesSheet = getEvaluationResponsesSheet();
  const submissionsSheet = getSubmissionResponsesSheet();

  if (!registrySheet || !overallScoresSheet || !reviewLogSheet || !evaluationResponsesSheet) {
    const errorMsg = 'Error: One or more required sheets not found.';
    Logger.log(errorMsg);
    throw new Error(errorMsg);
  }

  // Use the latest Evaluation request to determine the reporting month
  const reportingMonth = getReportingMonthFromRequestLog('Evaluation');
  if (!reportingMonth) {
    const errorMsg = 'Error: Could not determine reporting month from Request Log.';
    alertAndLog(errorMsg);
    throw new Error(errorMsg);
  }
  const currentReportingMonth = reportingMonth.firstDayDate;

  // Get submission request matching the reporting month
  const submissionRequest = getRequestByTypeMonthAndYear('Submission', reportingMonth.month, reportingMonth.year);
  if (!submissionRequest) {
    const errorMsg = `Error: No submission request found for ${reportingMonth.month} ${reportingMonth.year}`;
    Logger.log(errorMsg);
    throw new Error(errorMsg);
  }
  Logger.log(
    `Using Submission request for ${submissionRequest.month} ${submissionRequest.year} (window: ${submissionRequest.requestDateTime} to ${submissionRequest.windowEndDateTime})`
  );

  // Get headers and indices
  const headersRange = overallScoresSheet
    .getRange(COMPLIANCE_HEADER_ROW, 1, 1, overallScoresSheet.getLastColumn())
    .getValues()[0];
  const penaltyPointsColIndex = getRequiredColumnIndexByName(overallScoresSheet, SCORE_PENALTY_POINTS_COLUMN);
  const currentMonthColIndex =
    headersRange.findIndex((header) => header instanceof Date && header.getTime() === currentReportingMonth.getTime()) +
    1;

  if (currentMonthColIndex === 0) {
    const errorMsg = 'Error: Current reporting month column not found.';
    Logger.log(errorMsg);
    throw new Error(errorMsg);
  }
  Logger.log(`Current reporting month column index: ${currentMonthColIndex}`);

  // Get valid submitters and evaluators (pass submission window dates)
  const validSubmitters = getValidSubmissionEmails(
    submissionsSheet,
    submissionRequest.requestDateTime,
    submissionRequest.windowEndDateTime
  );
  const validEvaluators = getValidEvaluationEmails(evaluationResponsesSheet);
  Logger.log(`Valid submitters: ${validSubmitters.join(', ')}`);
  Logger.log(`Valid evaluators: ${validEvaluators.join(', ')}`);

  // Get assignments from Review Log (who evaluates whom)
  const assignments = getReviewLogAssignments();

  // Fetch data from Registry and filter non-expelled ambassadors
  const registryData = registrySheet
    .getRange(COMPLIANCE_FIRST_DATA_ROW, 1, registrySheet.getLastRow() - 1, registrySheet.getLastColumn())
    .getValues();
  const registryEmailColumn = getRequiredColumnIndexByName(registrySheet, AMBASSADOR_EMAIL_COLUMN) - 1;
  const registryDiscordColumn = getRequiredColumnIndexByName(registrySheet, AMBASSADOR_DISCORD_HANDLE_COLUMN) - 1;
  const registryStatusColumn = getRequiredColumnIndexByName(registrySheet, AMBASSADOR_STATUS_COLUMN) - 1;

  const ambassadorData = registryData
    .filter((row) => isActiveAmbassador(row, registryEmailColumn, registryStatusColumn))
    .map((row) => ({
      email: normalizeEmail(row[registryEmailColumn]),
      discordHandle: normalizeDiscordHandle(row[registryDiscordColumn]),
    }));

  Logger.log(`Filtered ambassadors: ${ambassadorData.length} valid rows`);

  // Collect month column indices in chronological order
  const monthColumns = [];
  for (let i = 0; i < headersRange.length; i++) {
    if (isDate(headersRange[i])) {
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
  const recentMonths = monthColumns.slice(-Math.min(COMPLIANCE_PERIOD_MONTHS, monthColumns.length));
  Logger.log(`Using ${recentMonths.length} most recent months for penalty calculation`);

  const inadequateContributionColIndex = getRequiredColumnIndexByName(
    overallScoresSheet,
    SCORE_INADEQUATE_CONTRIBUTION_COLUMN
  );

  return {
    overallScoresSheet,
    penaltyPointsColIndex,
    currentMonthColIndex,
    validSubmitters,
    validEvaluators,
    assignments,
    ambassadorData,
    recentMonths,
    inadequateContributionColIndex,
  };
}

/**
 * Applies color coding to current month cell based on missed activities.
 * @param {Object} params - Parameters object
 */
function applyCurrentMonthColorCoding(params) {
  const {
    email,
    discordHandle,
    overallScoresSheet,
    rowInScores,
    currentMonthColIndex,
    validSubmitters,
    validEvaluators,
    assignments,
  } = params;

  Logger.log(`Processing current month column (${currentMonthColIndex}) for ${discordHandle}`);
  const missedSubmission = didNotSubmitContribution(email, validSubmitters);
  const missedEvaluation = wasAssignedButDidNotEvaluate(email, assignments, validEvaluators);

  const currentCell = overallScoresSheet.getRange(rowInScores, currentMonthColIndex);

  if (missedSubmission && missedEvaluation) {
    currentCell.setBackground(COLOR_MISSED_SUBM_AND_EVAL);
    Logger.log(
      `Added ${COMPLIANCE_PENALTY_POINT_MISSED_BOTH} penalty points for ${discordHandle} (missed submission and evaluation).`
    );
  } else if (missedSubmission) {
    currentCell.setBackground(COLOR_MISSED_SUBMISSION);
    Logger.log(
      `Added ${COMPLIANCE_PENALTY_POINT_MISSED_SUBMISSION} penalty point for ${discordHandle} (missed submission).`
    );
  } else if (missedEvaluation) {
    currentCell.setBackground(COLOR_MISSED_EVALUATION);
    Logger.log(
      `Added ${COMPLIANCE_PENALTY_POINT_MISSED_EVALUATION} penalty point for ${discordHandle} (missed evaluation).`
    );
  }
}

/**
 * Calculates penalty points from a cell's background color.
 * @param {string} backgroundColor - Cell background color (lowercase)
 * @returns {number} Penalty points for the background color
 */
function getPenaltyPointsFromBackgroundColor(backgroundColor) {
  if (backgroundColor === COLOR_MISSED_SUBMISSION.toLowerCase()) {
    return COMPLIANCE_PENALTY_POINT_MISSED_SUBMISSION;
  } else if (backgroundColor === COLOR_MISSED_EVALUATION.toLowerCase()) {
    return COMPLIANCE_PENALTY_POINT_MISSED_EVALUATION;
  } else if (backgroundColor === COLOR_MISSED_SUBM_AND_EVAL.toLowerCase()) {
    return COMPLIANCE_PENALTY_POINT_MISSED_BOTH;
  }
  return 0;
}

/**
 * Processes a single month column to calculate penalty points and inadequate contributions.
 * @param {Sheet} overallScoresSheet - The Overall Score sheet
 * @param {number} rowInScores - Row number for the ambassador
 * @param {number} colIndex - Column index to process
 * @param {number} currentMonthColIndex - Current month column index
 * @param {string} email - Ambassador email
 * @param {string} discordHandle - Ambassador discord handle
 * @param {Array} validSubmitters - Valid submitter emails
 * @param {Array} validEvaluators - Valid evaluator emails
 * @param {Object} assignments - Assignment mappings
 * @returns {Object} Object containing penalty points and inadequate contribution count for this month
 */
function processMonthColumn(
  overallScoresSheet,
  rowInScores,
  colIndex,
  currentMonthColIndex,
  email,
  discordHandle,
  validSubmitters,
  validEvaluators,
  assignments
) {
  // Apply color coding if this is the current month
  if (colIndex === currentMonthColIndex) {
    applyCurrentMonthColorCoding({
      email,
      discordHandle,
      overallScoresSheet,
      rowInScores,
      currentMonthColIndex,
      validSubmitters,
      validEvaluators,
      assignments,
    });
  }

  const cell = overallScoresSheet.getRange(rowInScores, colIndex);
  const backgroundColor = cell.getBackground().toLowerCase();
  const scoreValue = cell.getValue();

  return {
    penaltyPoints: getPenaltyPointsFromBackgroundColor(backgroundColor),
    inadequateContribution: isInadequateContributionScore(scoreValue) ? 1 : 0,
  };
}

/**
 * Updates penalty points in the spreadsheet.
 * @param {Sheet} overallScoresSheet - The Overall Score sheet
 * @param {number} rowInScores - Row number for the ambassador
 * @param {number} penaltyPointsColIndex - Penalty points column index
 * @param {number} totalPenaltyPoints - Total penalty points to set
 * @param {string} discordHandle - Ambassador discord handle for logging
 */
function updatePenaltyPoints(
  overallScoresSheet,
  rowInScores,
  penaltyPointsColIndex,
  totalPenaltyPoints,
  discordHandle
) {
  overallScoresSheet.getRange(rowInScores, penaltyPointsColIndex).setValue(totalPenaltyPoints);
  Logger.log(`Updated penalty points for ${discordHandle} to ${totalPenaltyPoints}`);
}

/**
 * Updates inadequate contribution count in the spreadsheet.
 * @param {Sheet} overallScoresSheet - The Overall Score sheet
 * @param {number} rowInScores - Row number for the ambassador
 * @param {number} inadequateContributionColIndex - Inadequate contribution column index
 * @param {number} inadequateContributionCount - Count to set
 * @param {string} discordHandle - Ambassador discord handle for logging
 */
function updateInadequateContributionCount(
  overallScoresSheet,
  rowInScores,
  inadequateContributionColIndex,
  inadequateContributionCount,
  discordHandle
) {
  overallScoresSheet.getRange(rowInScores, inadequateContributionColIndex).setValue(inadequateContributionCount);
  Logger.log(`Updated Inadequate Contribution Count for ${discordHandle} to ${inadequateContributionCount}`);
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

  // Initialize all calculation data
  const {
    overallScoresSheet,
    penaltyPointsColIndex,
    currentMonthColIndex,
    validSubmitters,
    validEvaluators,
    assignments,
    ambassadorData,
    recentMonths,
    inadequateContributionColIndex,
  } = initializePenaltyCalculationData();

  // Process each ambassador
  ambassadorData.forEach((ambassador) => {
    const { email, discordHandle } = ambassador;

    const rowInScores = overallScoresSheet.createTextFinder(discordHandle).findNext()?.getRow();
    if (!rowInScores) {
      alertAndLog(`Discord handle not found in Overall Scores: ${discordHandle}`);
      return;
    }

    // Reset penalty points and recalculate based only on the last 6 months
    let totalPenaltyPoints = 0;
    let inadequateContributionCount = 0;

    // Process each recent month
    for (const colIndex of recentMonths) {
      const monthResult = processMonthColumn(
        overallScoresSheet,
        rowInScores,
        colIndex,
        currentMonthColIndex,
        email,
        discordHandle,
        validSubmitters,
        validEvaluators,
        assignments
      );

      totalPenaltyPoints += monthResult.penaltyPoints;
      inadequateContributionCount += monthResult.inadequateContribution;
    }

    // Update penalty points
    updatePenaltyPoints(overallScoresSheet, rowInScores, penaltyPointsColIndex, totalPenaltyPoints, discordHandle);

    // Update inadequate contribution count
    updateInadequateContributionCount(
      overallScoresSheet,
      rowInScores,
      inadequateContributionColIndex,
      inadequateContributionCount,
      discordHandle
    );

    // Use smart CRT referral logic to prevent duplicate referrals
    smartCRTReferralCheck(overallScoresSheet, rowInScores, recentMonths, discordHandle, inadequateContributionCount);
  });

  Logger.log('Penalty points calculation for submissions and evaluations completed.');
}

/**
 * Gets month columns from the overall scores sheet.
 * @param {Sheet} overallScoresSheet - The Overall Score sheet
 * @returns {Array} Array of month column indices
 */
function getMonthColumnsForPenaltyCalculation(overallScoresSheet) {
  const lastColumn = overallScoresSheet.getLastColumn();
  const spreadsheetTimeZone = getProjectTimeZone();
  const headers = overallScoresSheet.getRange(COMPLIANCE_HEADER_ROW, 1, 1, lastColumn).getValues()[0];
  const monthColumns = [];

  for (let col = 1; col <= lastColumn; col++) {
    const cellValue = headers[col - 1];
    if (isDate(cellValue)) {
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

  return monthColumns;
}

/**
 * Calculates penalty points for a specific period based on background colors.
 * @param {Array} backgroundColors - Array of background colors for the row
 * @param {number} startIndex - Start index of the period
 * @param {number} periodLength - Length of the period
 * @param {Array} monthColumns - Array of month column indices
 * @param {number} row - Row number for logging
 * @returns {number} Total penalty points for the period
 */
function calculatePenaltyPointsForPeriod(backgroundColors, startIndex, periodLength, monthColumns, row) {
  let periodTotal = 0;

  for (let j = startIndex; j < startIndex + periodLength; j++) {
    const cellBackgroundColor = backgroundColors[j].toLowerCase();

    switch (cellBackgroundColor) {
      case COLOR_MISSED_SUBMISSION:
        periodTotal += COMPLIANCE_PENALTY_POINT_MISSED_SUBMISSION;
        Logger.log(
          `Row ${row}: Adding ${COMPLIANCE_PENALTY_POINT_MISSED_SUBMISSION} point for missed submission at column ${monthColumns[j]}.`
        );
        break;
      case COLOR_MISSED_EVALUATION:
        periodTotal += COMPLIANCE_PENALTY_POINT_MISSED_EVALUATION;
        Logger.log(
          `Row ${row}: Adding ${COMPLIANCE_PENALTY_POINT_MISSED_EVALUATION} point for missed evaluation at column ${monthColumns[j]}.`
        );
        break;
      case COLOR_MISSED_SUBM_AND_EVAL:
        periodTotal += COMPLIANCE_PENALTY_POINT_MISSED_BOTH;
        Logger.log(
          `Row ${row}: Adding ${COMPLIANCE_PENALTY_POINT_MISSED_BOTH} points for missed submission and evaluation at column ${monthColumns[j]}.`
        );
        break;
      default:
        // No penalty for other colors
        break;
    }
  }

  return periodTotal;
}

/**
 * Processes a single ambassador row to calculate their max penalty points.
 * @param {Sheet} overallScoresSheet - The Overall Score sheet
 * @param {number} row - Row number to process
 * @param {Array} monthColumns - Array of month column indices
 * @param {number} periodLength - Length of the period for calculation
 * @param {number} maxPPCol - Column index for Max 6-Month PP
 */
function processAmbassadorRowForMaxPenalty(overallScoresSheet, row, monthColumns, periodLength, maxPPCol) {
  let maxPP = 0;
  Logger.log(`\nCalculating Max 6-Month Penalty Points for row ${row}.`);

  // Collect background colors for month columns
  const backgroundColors = overallScoresSheet
    .getRange(row, monthColumns[0], 1, monthColumns.length)
    .getBackgrounds()[0];

  Logger.log(`Row ${row}: Collected background colors for all month columns.`);

  // Iterate over all possible periods
  for (let i = 0; i <= monthColumns.length - periodLength; i++) {
    const periodTotal = calculatePenaltyPointsForPeriod(backgroundColors, i, periodLength, monthColumns, row);
    maxPP = Math.max(maxPP, periodTotal);
  }

  // Write the maximum penalty points to the "Max 6-Month PP" column
  const maxPPCell = overallScoresSheet.getRange(row, maxPPCol);
  maxPPCell.setValue(maxPP);
  maxPPCell.setHorizontalAlignment('center');

  // Set cell background to red if maxPP >= expulsion threshold
  if (maxPP >= MAX_PENALTY_POINTS_TO_EXPEL) {
    maxPPCell.setBackground(COLOR_EXPELLED);
    Logger.log(
      `Row ${row}: Max PP >= ${MAX_PENALTY_POINTS_TO_EXPEL}. Setting red background in Max 6-Month PP column.`
    );
  }

  Logger.log(`Row ${row}: Max 6-Month Penalty Points finalized as ${maxPP}.`);
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

  // Get overall score sheet and setup
  const overallScoresSheet = getOverallScoreSheet();
  const maxPPCol = getRequiredColumnIndexByName(overallScoresSheet, SCORE_MAX_6M_PP_COLUMN);
  const lastRow = overallScoresSheet.getLastRow();

  // Get month columns for penalty calculation
  const monthColumns = getMonthColumnsForPenaltyCalculation(overallScoresSheet);

  // Set the period length to the minimum of 6 or available months
  const periodLength = Math.min(COMPLIANCE_PERIOD_MONTHS, monthColumns.length);
  Logger.log(`Period length for calculation: ${periodLength} months.`);

  // Process each ambassador row
  for (let row = COMPLIANCE_FIRST_DATA_ROW; row <= lastRow; row++) {
    processAmbassadorRowForMaxPenalty(overallScoresSheet, row, monthColumns, periodLength, maxPPCol);
  }

  Logger.log('Completed calculating Max Penalty Points for all rows.');
}

/**
 * Gets ambassadors who exceed the penalty point threshold for expulsion.
 * @param {Sheet} overallScoresSheet - The Overall Score sheet
 * @param {number} scorePenaltiesColIndex - Column index for penalty points
 * @param {number} scoreDiscordHandleColIndex - Column index for discord handle in scores sheet
 * @returns {Array} Array of ambassador data objects with discord handles and penalty points
 */
function getAmbassadorsEligibleForExpulsion(overallScoresSheet, scorePenaltiesColIndex, scoreDiscordHandleColIndex) {
  const scoreData = overallScoresSheet
    .getRange(COMPLIANCE_FIRST_DATA_ROW, 1, overallScoresSheet.getLastRow() - 1, overallScoresSheet.getLastColumn())
    .getValues()
    .filter((row) => row[scorePenaltiesColIndex - 1] >= MAX_PENALTY_POINTS_TO_EXPEL)
    .map((row) => ({
      discordHandle: row[scoreDiscordHandleColIndex - 1],
      penaltyPoints: row[scorePenaltiesColIndex - 1],
    }));

  Logger.log(`Found ${scoreData.length} ambassadors eligible for expulsion`);
  return scoreData;
}

/**
 * Finds the registry row index for a given discord handle.
 * @param {Sheet} registrySheet - The Registry sheet
 * @param {string} discordHandle - Discord handle to search for
 * @param {number} registryDiscordHandleColIndex - Column index for discord handle in registry
 * @returns {number} Registry row index (1-based), or 0 if not found
 */
function findRegistryRowByDiscordHandle(registrySheet, discordHandle, registryDiscordHandleColIndex) {
  const registryData = registrySheet
    .getRange(COMPLIANCE_FIRST_DATA_ROW, registryDiscordHandleColIndex, registrySheet.getLastRow() - 1, 1)
    .getValues();

  const normalizedDiscordHandle = normalizeDiscordHandle(discordHandle);
  const rowIndex = registryData.findIndex((regRow) => normalizeDiscordHandle(regRow[0]) === normalizedDiscordHandle);
  return rowIndex >= 0 ? rowIndex + 2 : 0; // +2 to adjust for headers and 0-based index
}

/**
 * Processes a single ambassador for expulsion if not already expelled.
 * @param {Object} ambassadorData - Ambassador data with discordHandle and penaltyPoints
 * @param {Sheet} registrySheet - The Registry sheet
 * @param {number} registryDiscordHandleColIndex - Discord handle column index
 * @param {number} registryStatusColIndex - Status column index
 * @returns {boolean} True if ambassador was newly expelled, false otherwise
 */
function processAmbassadorForExpulsion(
  ambassadorData,
  registrySheet,
  registryDiscordHandleColIndex,
  registryStatusColIndex
) {
  const { discordHandle } = ambassadorData;

  const registryRowIndex = findRegistryRowByDiscordHandle(registrySheet, discordHandle, registryDiscordHandleColIndex);

  if (registryRowIndex <= 1) {
    Logger.log(`Error: Ambassador with discord handle: ${discordHandle} not found in the registry.`);
    return false;
  }

  const currentStatus = registrySheet.getRange(registryRowIndex, registryStatusColIndex).getValue();

  if (isAlreadyExpelled(currentStatus)) {
    Logger.log(`Notice: Ambassador ${discordHandle} is already marked as expelled.`);
    return false;
  }

  // Mark as expelled
  const currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yy');
  const updatedStatus = `${currentStatus} | Expelled [${currentDate}].`;
  registrySheet.getRange(registryRowIndex, registryStatusColIndex).setValue(updatedStatus);

  Logger.log(`Ambassador ${discordHandle} status updated to: "${updatedStatus}"`);
  return true;
}

/**
 * Expels ambassadors based on Max 6-Month PP. Tracks and notifies only newly expelled ambassadors.
 * If an ambassador already has status containing "Expelled", they are assumed to have already been notified and will be skipped.
 */
function expelAmbassadors() {
  Logger.log('Starting expelAmbassadors process.');

  const registrySheet = getRegistrySheet();
  const overallScoresSheet = getOverallScoreSheet();
  const scorePenaltiesColIndex = getRequiredColumnIndexByName(overallScoresSheet, SCORE_PENALTY_POINTS_COLUMN);
  const scoreDiscordHandleColIndex = getRequiredColumnIndexByName(overallScoresSheet, AMBASSADOR_DISCORD_HANDLE_COLUMN);
  const registryDiscordHandleColIndex = getRequiredColumnIndexByName(registrySheet, AMBASSADOR_DISCORD_HANDLE_COLUMN);
  const registryStatusColIndex = getRequiredColumnIndexByName(registrySheet, AMBASSADOR_STATUS_COLUMN);

  // Get ambassadors eligible for expulsion
  const eligibleAmbassadors = getAmbassadorsEligibleForExpulsion(
    overallScoresSheet,
    scorePenaltiesColIndex,
    scoreDiscordHandleColIndex
  );

  // Process each ambassador and collect newly expelled ones
  const newlyExpelled = [];
  eligibleAmbassadors.forEach((ambassadorData) => {
    const wasNewlyExpelled = processAmbassadorForExpulsion(
      ambassadorData,
      registrySheet,
      registryDiscordHandleColIndex,
      registryStatusColIndex
    );

    if (wasNewlyExpelled) {
      newlyExpelled.push(ambassadorData.discordHandle);
    }
  });

  // Send notifications to newly expelled ambassadors
  newlyExpelled.forEach((discordHandle) => {
    sendExpulsionNotifications(discordHandle);
  });

  Logger.log(`expelAmbassadors process completed. ${newlyExpelled.length} ambassadors newly expelled.`);
}

/**
 * Creates expulsion email body by replacing template tokens.
 * @param {string} discordHandle - Ambassador's discord handle
 * @param {string} expulsionDate - Formatted expulsion date
 * @returns {string} Email body with replaced tokens
 */
function createExpulsionEmailBody(discordHandle, expulsionDate) {
  const startDate = 'your start date'; // TODO: Add start date tracking to registry

  return EXPULSION_EMAIL_TEMPLATE.replace(/{Discord Handle}/g, discordHandle)
    .replace(/{Expulsion Date}/g, expulsionDate)
    .replace(/{Start Date}/g, startDate)
    .replace(/{Sponsor Email}/g, SPONSOR_EMAIL);
}

/**
 * Sends expulsion notifications to the expelled ambassador and sponsor.
 * @param {string} discordHandle - The ambassador's discord handle to look up the email.
 */
function sendExpulsionNotifications(discordHandle) {
  Logger.log(`Sending expulsion notifications for ambassador with discord handle: ${discordHandle}`);

  // Use existing utility to get ambassador email and discord handle
  const ambassadorInfo = lookupEmailAndDiscord(discordHandle);
  if (!ambassadorInfo || !isValidEmail(ambassadorInfo.email)) {
    Logger.log(
      `Skipping notification for ambassador with discord handle: ${discordHandle} due to missing or invalid email.`
    );
    return;
  }

  const { email } = ambassadorInfo;
  const subject = 'Expulsion from the Program';

  // Prepare expulsion date
  const currentDate = new Date();
  const timeZone = getProjectTimeZone();
  const expulsionDate = Utilities.formatDate(currentDate, timeZone, 'MMMM dd, yyyy');

  // Create email bodies
  const expulsionEmailBody = createExpulsionEmailBody(discordHandle, expulsionDate);
  const sponsorBody = `Ambassador ${email} (${discordHandle}) has been expelled from the program for Failure to Participate according to Article 2, Section 10 of the Bylaws.`;

  // Send notifications
  sendEmailNotification(email, subject, expulsionEmailBody);
  Logger.log(`Expulsion email sent to ${email}.`);

  sendEmailNotification(SPONSOR_EMAIL, subject, sponsorBody);
  Logger.log(`Notification sent to sponsor for expelled ambassador: ${email} (${discordHandle}).`);
}

/**
 * Parses a CRT referral history string into an array of referral entries.
 * Format: "months:date|months:date" where months is comma-separated like "2024-01,2024-02"
 * @param {string} historyString - The CRT referral history string
 * @returns {Array} Array of referral entries: [{months: ['2024-01', '2024-02'], date: '2024-04-15'}]
 */
function parseCRTReferralHistory(historyString) {
  if (!historyString || historyString.trim() === '') {
    return [];
  }

  return historyString.split('|').map((entry) => {
    const [monthsStr, dateStr] = entry.split(':');
    return {
      months: monthsStr.split(',').map((m) => m.trim()),
      date: dateStr.trim(),
    };
  });
}

/**
 * Formats CRT referral history entries into a string for storage.
 * @param {Array} referralEntries - Array of referral entries
 * @returns {string} Formatted history string
 */
function formatCRTReferralHistory(referralEntries) {
  return referralEntries.map((entry) => `${entry.months.join(',')}:${entry.date}`).join('|');
}

/**
 * Gets the month string in YYYY-MM format for a given date.
 * @param {Date} date - The date to convert
 * @returns {string} Month string in YYYY-MM format
 */
function getMonthString(date) {
  const year = date.getFullYear();
  const month = (date.getMonth() + 1).toString().padStart(2, '0');
  return `${year}-${month}`;
}

/**
 * Gets inadequate contribution months for an ambassador from the overall scores sheet.
 * @param {Sheet} overallScoresSheet - The Overall Score sheet
 * @param {number} rowIndex - Row index for the ambassador
 * @param {Array} recentMonths - Array of recent month column indices
 * @returns {Array} Array of month strings where ambassador scored < 3
 */
function getInadequateContributionMonths(overallScoresSheet, rowIndex, recentMonths) {
  const headers = overallScoresSheet
    .getRange(COMPLIANCE_HEADER_ROW, 1, 1, overallScoresSheet.getLastColumn())
    .getValues()[0];
  const inadequateMonths = [];

  for (const colIndex of recentMonths) {
    const cellValue = overallScoresSheet.getRange(rowIndex, colIndex).getValue();
    if (isInadequateContributionScore(cellValue)) {
      const headerDate = headers[colIndex - 1];
      if (isDate(headerDate)) {
        inadequateMonths.push(getMonthString(headerDate));
      }
    }
  }

  return inadequateMonths;
}

/**
 * Checks if there are new inadequate contribution months that haven't been referred to CRT.
 * @param {Array} currentInadequateMonths - Current months with inadequate scores
 * @param {Array} referralHistory - Previous CRT referral entries
 * @returns {Array} New inadequate months not previously referred
 */
function getNewInadequateMonths(currentInadequateMonths, referralHistory) {
  const previouslyReferredMonths = new Set();

  // Collect all months that were previously referred
  referralHistory.forEach((entry) => {
    entry.months.forEach((month) => previouslyReferredMonths.add(month));
  });

  // Return only months that haven't been referred before
  return currentInadequateMonths.filter((month) => !previouslyReferredMonths.has(month));
}

/**
 * Updates the CRT referral history for an ambassador.
 * @param {Sheet} overallScoresSheet - The Overall Score sheet
 * @param {number} rowIndex - Row index for the ambassador
 * @param {number} historyColIndex - Column index for CRT referral history
 * @param {Array} newInadequateMonths - New inadequate months to record
 * @param {string} discordHandle - Ambassador discord handle for logging
 */
function updateCRTReferralHistory(overallScoresSheet, rowIndex, historyColIndex, newInadequateMonths, discordHandle) {
  const currentHistoryString = overallScoresSheet.getRange(rowIndex, historyColIndex).getValue() || '';
  const referralHistory = parseCRTReferralHistory(currentHistoryString);

  // Add new referral entry
  const currentDate = Utilities.formatDate(new Date(), getProjectTimeZone(), 'yyyy-MM-dd');
  referralHistory.push({
    months: newInadequateMonths,
    date: currentDate,
  });

  // Update the cell
  const updatedHistoryString = formatCRTReferralHistory(referralHistory);
  overallScoresSheet.getRange(rowIndex, historyColIndex).setValue(updatedHistoryString);

  Logger.log(`Updated CRT referral history for ${discordHandle}: ${updatedHistoryString}`);
}

/**
 * Smart CRT referral function that only refers ambassadors for new inadequate contribution months.
 * Prevents multiple referrals for the same poor scoring months.
 * @param {Sheet} overallScoresSheet - The Overall Score sheet
 * @param {number} rowIndex - Row index for the ambassador
 * @param {Array} recentMonths - Array of recent month column indices
 * @param {string} discordHandle - Ambassador discord handle
 * @param {number} inadequateContributionCount - Number of inadequate contributions
 */
function smartCRTReferralCheck(overallScoresSheet, rowIndex, recentMonths, discordHandle, inadequateContributionCount) {
  // Only proceed if threshold is met
  if (inadequateContributionCount < MAX_INADEQUATE_CONTRIBUTION_COUNT_TO_REFER) {
    return;
  }

  // Get CRT referral history column
  const crtHistoryColIndex = getRequiredColumnIndexByName(overallScoresSheet, SCORE_CRT_REFERRAL_HISTORY_COLUMN);
  const currentHistoryString = overallScoresSheet.getRange(rowIndex, crtHistoryColIndex).getValue() || '';
  const referralHistory = parseCRTReferralHistory(currentHistoryString);

  // Get current inadequate contribution months
  const currentInadequateMonths = getInadequateContributionMonths(overallScoresSheet, rowIndex, recentMonths);

  // Find new inadequate months that haven't been referred before
  const newInadequateMonths = getNewInadequateMonths(currentInadequateMonths, referralHistory);

  // Only refer if there are new inadequate months
  if (newInadequateMonths.length > 0) {
    Logger.log(
      `${discordHandle}: Found ${newInadequateMonths.length} new inadequate months: ${newInadequateMonths.join(', ')}`
    );

    // Update referral history
    updateCRTReferralHistory(overallScoresSheet, rowIndex, crtHistoryColIndex, newInadequateMonths, discordHandle);

    // Send CRT referral
    referInadequateContributionToCRT(discordHandle, inadequateContributionCount, newInadequateMonths);
  } else {
    Logger.log(
      `${discordHandle}: All inadequate months (${currentInadequateMonths.join(', ')}) have already been referred to CRT. Skipping referral.`
    );
  }
}

/**
 * Refers an ambassador to the CRT for inadequate contribution.
 * Sends a notification email directly to the subject ambassador, copying the sponsor.
 * @param {string} discordHandle - The Discord handle of the ambassador being referred.
 * @param {number} inadequateContributionCount - The number of times the ambassador scored below the inadequate contribution threshold in the last 6 months.
 * @param {Array} newInadequateMonths - Optional array of new inadequate months being referred
 */
function referInadequateContributionToCRT(discordHandle, inadequateContributionCount, newInadequateMonths = []) {
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
    const deadlineDate = getBusinessDaysFromToday(COMPLIANCE_BUSINESS_DAYS_DEADLINE);

    // Compose the email using the template
    let emailBody = INADEQUATE_CONTRIBUTION_NOTIFICATION_EMAIL_TEMPLATE.replaceAll(
      '{monthName}',
      currentMonthName
    ).replaceAll('{deadlineDate}', deadlineDate);

    // If specific months are provided, add them to the email
    if (newInadequateMonths && newInadequateMonths.length > 0) {
      const monthDetails = `\n\nSpecifically, this referral is for inadequate contributions in: ${newInadequateMonths.join(', ')}`;
      emailBody += monthDetails;
    }

    const subject = `Autonomys AmbasasadorOS CRT complaint - Inadequate Contribution`;

    // Send to ambassador, CC sponsor
    sendEmailNotification(ambassadorEmail, subject, emailBody, '', SPONSOR_EMAIL);

    const monthsInfo =
      newInadequateMonths && newInadequateMonths.length > 0 ? ` for months: ${newInadequateMonths.join(', ')}` : '';
    Logger.log(
      `Sent inadequate contribution notification to ${ambassadorEmail} (${ambassadorDiscord})${monthsInfo}, CC: ${SPONSOR_EMAIL}`
    );
  } catch (e) {
    Logger.log('Error in referInadequateContributionToCRT: ' + e);
  }
}

// ===== ANONYMOUS SCORES PUBLISHING FUNCTIONS =====

/**
 * Publishes anonymous audit scores to a Google Sheet.
 * Creates a new sheet tab for the current reporting month.
 */
function publishAnonymousScoresToGoogleSheet() {
  if (!ANONYMOUS_SCORES_SPREADSHEET_ID) {
    Logger.log('Anonymous scores spreadsheet not configured. Skipping score publishing.');
    return;
  }

  Logger.log('Starting anonymous score publishing to Google Sheets.');

  try {
    // Get reporting month data using existing helper function
    const reportingMonth = getReportingMonthFromRequestLog('Evaluation');
    if (!reportingMonth) {
      throw new Error('Could not determine reporting month from Request Log.');
    }

    const currentMonthDate = reportingMonth.firstDayDate;
    const spreadsheetTimeZone = getProjectTimeZone();
    const monthName = Utilities.formatDate(currentMonthDate, spreadsheetTimeZone, 'MMMM yyyy');

    // Get submission window times for the reporting month from Request Log
    // Use the specific month/year that matches the reporting month, not the latest request
    const submissionRequest = getRequestByTypeMonthAndYear('Submission', reportingMonth.month, reportingMonth.year);
    if (!submissionRequest) {
      throw new Error(
        `Could not find Submission request for ${reportingMonth.month} ${reportingMonth.year} in Request Log.`
      );
    }

    Logger.log(
      `Using Submission request for ${submissionRequest.month} ${submissionRequest.year} to match reporting month (${reportingMonth.month} ${reportingMonth.year})`
    );

    const submissionWindowStart = submissionRequest.requestDateTime;
    const submissionWindowEnd = submissionRequest.windowEndDateTime;
    Logger.log(`Using submission window: ${submissionWindowStart} to ${submissionWindowEnd}`);

    // Collect anonymous score data
    const anonymousScores = collectAnonymousScoreData(monthName, submissionWindowStart, submissionWindowEnd);

    if (anonymousScores.length === 0) {
      Logger.log('No scores found to publish.');
      return;
    }

    // Create sheet tab and publish data
    createAnonymousScoresSheet(monthName, anonymousScores);

    Logger.log(`Successfully published ${anonymousScores.length} anonymous scores to sheet for ${monthName}.`);
  } catch (error) {
    Logger.log(`Error in publishAnonymousScoresToGoogleSheet: ${error}`);
    throw error;
  }
}

/**
 * Collects anonymous score data from the current month's evaluation results.
 * Includes primary team and contribution details for each submitter.
 * @param {string} monthName - The reporting month name (e.g., "January 2025")
 * @param {Date} submissionWindowStart - Start of submission window for the reporting month
 * @param {Date} submissionWindowEnd - End of submission window for the reporting month
 * @returns {Array} Array of anonymous score objects
 */
function collectAnonymousScoreData(monthName, submissionWindowStart, submissionWindowEnd) {
  Logger.log(`Collecting anonymous score data for ${monthName}.`);

  try {
    const scoresSpreadsheet = getScoresSpreadsheet();
    const monthSheet = scoresSpreadsheet.getSheetByName(monthName);

    if (!monthSheet) {
      Logger.log(`Month sheet "${monthName}" not found. Cannot collect scores.`);
      return [];
    }

    // Get column indices using existing helper functions and constants
    const submitterColIndex = getRequiredColumnIndexByName(monthSheet, SUBMITTER_HANDLE_COLUMN_IN_MONTHLY_SCORE);
    const score1ColIndex = getColumnIndexByName(monthSheet, GRADE_EVAL_1_SCORE_COLUMN);
    const remarks1ColIndex = getColumnIndexByName(monthSheet, GRADE_EVAL_1_REMARKS_COLUMN);
    const score2ColIndex = getColumnIndexByName(monthSheet, GRADE_EVAL_2_SCORE_COLUMN);
    const remarks2ColIndex = getColumnIndexByName(monthSheet, GRADE_EVAL_2_REMARKS_COLUMN);
    const score3ColIndex = getColumnIndexByName(monthSheet, GRADE_EVAL_3_SCORE_COLUMN);
    const remarks3ColIndex = getColumnIndexByName(monthSheet, GRADE_EVAL_3_REMARKS_COLUMN);
    const finalScoreColIndex = getColumnIndexByName(monthSheet, GRADE_FINAL_SCORE_COLUMN);

    // Get all data from the sheet
    const data = monthSheet
      .getRange(COMPLIANCE_FIRST_DATA_ROW, 1, monthSheet.getLastRow() - 1, monthSheet.getLastColumn())
      .getValues();

    const anonymousScores = [];

    data.forEach((row, index) => {
      const submitterHandle = row[submitterColIndex - 1];

      // Skip empty rows
      if (!submitterHandle || submitterHandle.toString().trim() === '') {
        return;
      }

      // Get email from discord handle to fetch primary team and contributions
      const submitterInfo = lookupEmailAndDiscord(submitterHandle);
      const submitterEmail = submitterInfo ? submitterInfo.email : '';

      // Get primary team
      let primaryTeam = '';
      if (submitterEmail) {
        try {
          primaryTeam = getAmbassadorPrimaryTeam(submitterEmail);
        } catch (error) {
          Logger.log(`Error getting primary team for ${submitterEmail}: ${error}`);
        }
      }

      // Get contribution details
      let contributionDetails = '';
      if (submitterEmail) {
        try {
          contributionDetails = getContributionDetailsByEmail(
            submitterEmail,
            submissionWindowStart,
            submissionWindowEnd
          );
          // Remove HTML tags for cleaner display in spreadsheet
          contributionDetails = stripHtmlTags(contributionDetails);
        } catch (error) {
          Logger.log(`Error getting contributions for ${submitterEmail}: ${error}`);
          contributionDetails = 'No contribution details found.';
        }
      }

      const scoreRow = {
        Submitter: submitterHandle,
        'Primary Team': primaryTeam || '',
        Contributions: contributionDetails || '',
        'Score-1': score1ColIndex > 0 ? row[score1ColIndex - 1] || '' : '',
        'Remarks-1': remarks1ColIndex > 0 ? row[remarks1ColIndex - 1] || '' : '',
        'Score-2': score2ColIndex > 0 ? row[score2ColIndex - 1] || '' : '',
        'Remarks-2': remarks2ColIndex > 0 ? row[remarks2ColIndex - 1] || '' : '',
        'Score-3': score3ColIndex > 0 ? row[score3ColIndex - 1] || '' : '',
        'Remarks-3': remarks3ColIndex > 0 ? row[remarks3ColIndex - 1] || '' : '',
        'Final Score': finalScoreColIndex > 0 ? row[finalScoreColIndex - 1] || '' : '',
      };

      anonymousScores.push(scoreRow);
    });

    Logger.log(`Collected ${anonymousScores.length} score records for ${monthName}.`);
    return anonymousScores;
  } catch (error) {
    Logger.log(`Error collecting anonymous score data: ${error}`);
    throw error;
  }
}

/**
 * Creates a new sheet tab in the anonymous scores spreadsheet with score data.
 * Includes Primary Team and Contributions columns.
 * @param {string} monthName - The reporting month name
 * @param {Array} anonymousScores - Array of anonymous score objects
 */
function createAnonymousScoresSheet(monthName, anonymousScores) {
  Logger.log(`Creating anonymous scores sheet for ${monthName}.`);

  try {
    // Get the anonymous scores spreadsheet
    const spreadsheet = SpreadsheetApp.openById(ANONYMOUS_SCORES_SPREADSHEET_ID);

    // Format sheet name: "Scores - {Month} {Year}" or "TEST - Scores - {Month} {Year}"
    const sheetName = TESTING ? `TEST - Scores - ${monthName}` : `Scores - ${monthName}`;

    // Check if sheet already exists
    let sheet = spreadsheet.getSheetByName(sheetName);
    if (sheet) {
      Logger.log(`Sheet "${sheetName}" already exists. Clearing existing data.`);
      sheet.clear();
    } else {
      Logger.log(`Creating new sheet: "${sheetName}"`);
      sheet = spreadsheet.insertSheet(sheetName);
    }

    // Set up headers with Primary Team and Contributions columns
    const headers = [
      'Submitter',
      'Primary Team',
      'Contributions',
      'Score-1',
      'Remarks-1',
      'Score-2',
      'Remarks-2',
      'Score-3',
      'Remarks-3',
      'Final Score',
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Format header row
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');

    // Insert score data
    if (anonymousScores.length > 0) {
      const dataRows = anonymousScores.map((score) => [
        score.Submitter || '',
        score['Primary Team'] || '',
        score.Contributions || '',
        score['Score-1'] || '',
        score['Remarks-1'] || '',
        score['Score-2'] || '',
        score['Remarks-2'] || '',
        score['Score-3'] || '',
        score['Remarks-3'] || '',
        score['Final Score'] || '',
      ]);

      sheet.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);
      Logger.log(`Inserted ${dataRows.length} score records into sheet "${sheetName}".`);
    }

    // Auto-resize columns first
    for (let i = 1; i <= headers.length; i++) {
      sheet.autoResizeColumn(i);
    }

    // Set specific column widths for Contributions and Remarks columns
    sheet.setColumnWidth(3, 500); // Contributions column (C)
    sheet.setColumnWidth(5, 200); // Remarks-1 column (E)
    sheet.setColumnWidth(7, 200); // Remarks-2 column (G)
    sheet.setColumnWidth(9, 200); // Remarks-3 column (I)

    // Enable word wrap for Contributions and Remarks columns
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).setWrap(true); // Contributions (C)
      sheet.getRange(2, 5, sheet.getLastRow() - 1, 1).setWrap(true); // Remarks-1 (E)
      sheet.getRange(2, 7, sheet.getLastRow() - 1, 1).setWrap(true); // Remarks-2 (G)
      sheet.getRange(2, 9, sheet.getLastRow() - 1, 1).setWrap(true); // Remarks-3 (I)
    }

    // Freeze header row
    sheet.setFrozenRows(1);

    Logger.log(`Successfully created anonymous scores sheet "${sheetName}" with ${anonymousScores.length} records.`);
  } catch (error) {
    Logger.log(`Error creating anonymous scores sheet: ${error}`);
    throw error;
  }
}

// ===== Predicate Functions for Complex Conditionals =====

/**
 * Strips HTML tags from a string through repeated application to prevent bypass via nested tags.
 * This prevents HTML injection vulnerabilities by ensuring all tags are removed, even if nested.
 * @param {string} str - String potentially containing HTML tags
 * @returns {string} Sanitized string with all HTML tags removed
 */
function stripHtmlTags(str) {
  if (!str) return '';
  let previous;
  let sanitized = str;
  do {
    previous = sanitized;
    sanitized = sanitized.replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, '');
  } while (sanitized !== previous);
  return sanitized;
}

/**
 * Checks if an ambassador did not submit their monthly contribution.
 * @param {string} email - Ambassador email
 * @param {Array} validSubmitters - Array of valid submitter emails
 * @returns {boolean} True if ambassador did not submit contribution
 */
function didNotSubmitContribution(email, validSubmitters) {
  return !validSubmitters.map(normalizeEmail).includes(normalizeEmail(email));
}

/**
 * Checks if an ambassador was assigned to evaluate but did not submit evaluation.
 * @param {string} email - Ambassador email
 * @param {Object} assignments - Assignment object mapping submitters to evaluators
 * @param {Array} validEvaluators - Array of valid evaluator emails
 * @returns {boolean} True if ambassador was assigned but did not evaluate
 */
function wasAssignedButDidNotEvaluate(email, assignments, validEvaluators) {
  return Object.values(assignments).some(
    (evaluators) =>
      evaluators.map(normalizeEmail).includes(normalizeEmail(email)) &&
      !validEvaluators.map(normalizeEmail).includes(normalizeEmail(email))
  );
}

/**
 * Checks if a score value is below the inadequate contribution threshold.
 * @param {*} scoreValue - Score value to check
 * @returns {boolean} True if score is a number below threshold
 */
function isInadequateContributionScore(scoreValue) {
  return typeof scoreValue === 'number' && scoreValue < INADEQUATE_CONTRIBUTION_SCORE_THRESHOLD;
}
