// (( Global Variables ))
// Declare & initialize global variables; these will be updated by the setProductionVariables() or setTestVariables() functions

// testing constant will be used to load production vs. test values for the global variables
const testing = true; // Set to true for testing (logs instead of sending emails, uses test sheets and forms)
var SEND_EMAIL; // Will control whether emails are sent - must be true for production; may be true or false for testing depending on testing needs.

// Provide the actual Id of the google sheet for the registry and scoreing sheets in EnvironmentVariables[Prod|Test].js:
var AMBASSADOR_REGISTRY_SPREADSHEET_ID = ''; //"Ambassador Registry"
var AMBASSADORS_SCORES_SPREADSHEET_ID = ''; // "Ambassadors' Scores"
var AMBASSADORS_SUBMISSIONS_SPREADSHEET_ID = ''; // "Ambassador Submission Responses"
var EVALUATION_RESPONSES_SPREADSHEET_ID = ''; // "Evaluation Responses"

// Provide the actual Id and submission URL for the submission and evaluation google forms in EnvironmentVariables[Prod|Test].js:
var SUBMISSION_FORM_ID = ''; // ID for Submission form
var EVALUATION_FORM_ID = ''; // ID for Evaluation form
var SUBMISSION_FORM_URL = ''; // Submission Form URL for mailing
var EVALUATION_FORM_URL = ''; // Evaluation Form URL for mailing

// Provide the actual sheet names for the registry, review log, CRT, and overall score sheets in EnvironmentVariables[Prod|Test].js:
var REGISTRY_SHEET_NAME = '';
var FORM_RESPONSES_SHEET_NAME = '';
var REVIEW_LOG_SHEET_NAME = '';
var CONFLICT_RESOLUTION_TEAM_SHEET_NAME = '';
var OVERALL_SCORE_SHEET_NAME = ''; // Overall score sheet in Ambassadors' Scores
var EVAL_FORM_RESPONSES_SHEET_NAME = ''; // Evaluation Form responses sheet

// Columns in Registry Sheet:
// set the actual values in EnvironmentVariables[Prod|Test].js
var AMBASSADOR_EMAIL_COLUMN = '';
var AMBASSADOR_DISCORD_HANDLE_COLUMN = '';
var AMBASSADOR_STATUS_COLUMN = '';

// Sponsor Email (for notifications when ambassadors are expelled)
// set the actual values in EnvironmentVariables[Prod|Test].js
var SPONSOR_EMAIL = ''; // Sponsor's email

// Color variables .The color hex string must be in lowercase!
// set the actual values in EnvironmentVariables[Prod|Test].js
var COLOR_MISSED_SUBMISSION = '';
var COLOR_MISSED_EVALUATION = '';
var COLOR_EXPELLED = '';
var COLOR_MISSED_SUBM_AND_EVAL = '';
var COLOR_OLD_MISSED_SUBMISSION = '';

// Triggers and Delays
// These values will set the due date and reminder schedule for Submissions and Evaluations.
// The Submission or Evaluation will be due after the relevant WINDOW_MINUTES,
// and each ambassador will receive a reminder after the relevant WINDOW_REMINDER_MINUTES.
// specifies as days * hours * minutes
// set the actual values in EnvironmentVariables[Prod|Test].js
var SUBMISSION_WINDOW_MINUTES = '';
var SUBMISSION_WINDOW_REMINDER_MINUTES = ''; // how many minutes after Submission Requests sent to remind
var EVALUATION_WINDOW_MINUTES = '';
var EVALUATION_WINDOW_REMINDER_MINUTES = ''; // how many minutes after Evaluation Requests sent to remind

const ButtonSet = {
  OK: 'OK',
  OK_CANCEL: 'OK_CANCEL',
  YES_NO: 'YES_NO',
  YES_NO_CANCEL: 'YES_NO_CANCEL',
};

const ButtonResponse = {
  OK: 'ok',
  CANCEL: 'cancel',
  YES: 'yes',
  NO: 'no',
};

if (testing) {
  setTestVariables();
} else {
  setProductionVariables();
}

//		 ======= 	 Email Content Templates 	=======

// Request Submission Email Template
let REQUEST_SUBMISSION_EMAIL_TEMPLATE = `
<p>Dear {AmbassadorDiscordHandle},</p>

<p>Please submit your deliverables for {Month} {Year} using the link below:</p>
<p><a href="{SubmissionFormURL}">Submission Form</a></p>

<p>The deadline is {SUBMISSION_DEADLINE_DATE}.</p>

<p>Thank you,<br>
Ambassador Program Team</p>
`;

// Request Evaluation Email Template
let REQUEST_EVALUATION_EMAIL_TEMPLATE = `
<p>Dear {AmbassadorDiscordHandle},</p>
<p>Please review the following deliverables for the month of {Month} by {AmbassadorSubmitter}:</p>
<p>{SubmissionsList}</p>

<p>Assign a grade using the form:</p>
<p><a href="{EvaluationFormURL}">Evaluation Form</a></p>

<p>The deadline is {EVALUATION_DEADLINE_DATE}.</p>

<p>Thank you,<br>Ambassador Program Team</p>
`;

// Reminder Email Template
let REMINDER_EMAIL_TEMPLATE = `
Hi there! Just a friendly reminder that we’re still waiting for your response to the Request for Submission/Evaluation. Please respond soon to avoid any penalties. Thank you!.
`;

// Penalty Warning Email Template
let PENALTY_WARNING_EMAIL_TEMPLATE = `
Dear Ambassador,
You have been assessed one penalty point for failing to meet Submission or Evaluation deadlines. Further penalties may result in expulsion from the program. Please be vigillant.
`;

// Expulsion Email Template
let EXPULSION_EMAIL_TEMPLATE = `
Dear Ambassador,
We regret to inform you that you have been expelled from the program due to multiple missed deadlines.
`;

// Notify Upcoming Peer Review Email Template
let NOTIFY_UPCOMING_PEER_REVIEW = `
Dear Ambassador,
By this we notify you about upcoming Peer Review mailing, please be vigilant!
`;

//"CRT Selecting Notification" email template
let CRT_SELECTING_NOTIFICATION_TEMPLATE = `
Dear Ambassador, you have been selected for next Conflict Resolution Team ...
`;

// EXEMPTION FROM EVALUATION e-mail template
let EXEMPTION_FROM_EVALUATION_TEMPLATE = `
Dear Ambassador, you have been relieved of the obligation to evaluate your colleagues this month.
`;

//    On Open Menu

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  Logger.log('Initializing menu on spreadsheet open.');
  ui.createMenu('Ambassador Program')
    .addItem('Request Submissions', 'requestSubmissionsModule') // Request Submissions
    .addItem('Request Evaluations', 'requestEvaluationsModule') // Request Evaluations
    .addItem('Compliance Audit', 'runComplianceAudit') // Process Scores and Penalties
    .addItem('Notify Upcoming Peer Review', 'notifyUpcomingPeerReview') // Peer Review notifications
    .addItem('Select CRT members', 'selectCRTMembers') // CRT
    .addItem('🔧️Force Authorization', 'forceAuthorization') // Authorization trigger
    .addItem('🔧️Check missing Emails/DiscorHandles', 'check_missing_data') // Checks Registry sheet for completeness of data. Recommended to run every cycle if multiple changes were made.
    .addItem('🔧️Delete Existing Triggers', 'deleteExistingTriggers') // Optional item
    .addItem('🔧️Create/Sync Columns', 'syncRegistryColumnsToOverallScore') // creates Ambassador Status column in Overall score sheet; Syncs Ambassadors' Discord Handles and Ambassador Status columns between Registry and Overall score.
    .addToUi();
  Logger.log('Menu initialized.');
}

// ⚠️ functions to prevent Google using of cached outdated variables when run from GUI
function refreshScriptState() {
  Logger.log('Starting Refresh Script State');
  refreshGlobalVariables();
  Logger.log('Script state refreshed: cache cleared, variables refreshed, and flush completed.');
}

function refreshGlobalVariables() {
  Logger.log('Refreshing global variables.');

  if (testing) {
    setTestVariables();
  } else {
    setProductionVariables();
  }

  // Log the reinitialization of templates
  Logger.log('Templates and constants reinitialized to ensure accurate processing.');

  Logger.log('Color for missed submission: ' + COLOR_MISSED_SUBMISSION);
  Logger.log('Color for missed submission: ' + COLOR_OLD_MISSED_SUBMISSION);
  Logger.log('Color for missed evaluation: ' + COLOR_MISSED_EVALUATION);
  Logger.log('Color for missed submiss and eval: ' + COLOR_MISSED_SUBM_AND_EVAL);
  Logger.log('Color for expelled: ' + COLOR_EXPELLED);
}

// ===== Generic function to send email =====
function sendEmailNotification(recipientEmail, subject, body) {
  try {
    if (SEND_EMAIL) {
      MailApp.sendEmail(recipientEmail, subject, body);
      Logger.log(`Email sent to: ${recipientEmail}, Subject: ${subject}`);
    } else {
      if (!testing) {
        Logger.log(
          `WARNING: Production mode with email disabled. Email logged but NOT SENT to: ${recipientEmail}, Subject: ${subject}`
        );
      } else {
        Logger.log(`Test mode with email disabled: Email to be sent to: ${recipientEmail}, Subject: ${subject}`);
      }
    }
  } catch (error) {
    Logger.log(`Failed to send email to ${recipientEmail}: ${error.message}`);
  }
}

//     Submission/Evaluation WINDOW TIME SET/GET

// Save the submission window start time (in PST)
function setSubmissionWindowStart(time) {
  const formattedTime = Utilities.formatDate(time, getProjectTimeZone(), 'yyyy-MM-dd HH:mm:ss z'); // Format the time
  PropertiesService.getScriptProperties().setProperty('submissionWindowStart', formattedTime); // Save the formatted time
  Logger.log(`Submission window start time saved: ${formattedTime}`);
}

function getSubmissionWindowStart() {
  Logger.log('Getting submission window start time.');
  const scriptProperties = PropertiesService.getScriptProperties();
  const startDateStr = scriptProperties.getProperty('submissionWindowStart');

  if (!startDateStr) {
    Logger.log('Submission window start date not found!');
    return null;
  }

  const startDate = new Date(startDateStr);
  Logger.log(`Submission window started at: ${startDate}`);
  return startDate;
}

function getSubmissionWindowTimes() {
  const submissionWindowStartStr = PropertiesService.getScriptProperties().getProperty('submissionWindowStart');
  if (!submissionWindowStartStr) {
    throw new Error('Evaluation window start time not found.');
  }
  const submissionWindowStart = new Date(submissionWindowStartStr);
  const submissionWindowEnd = new Date(submissionWindowStart.getTime() + EVALUATION_WINDOW_MINUTES * 60 * 1000);
  return { submissionWindowStart, submissionWindowEnd };
}

// Save the evaluation window start time (in PST)
function setEvaluationWindowStart(time) {
  const formattedTime = Utilities.formatDate(time, getProjectTimeZone(), 'yyyy-MM-dd HH:mm:ss z'); // Format the time
  PropertiesService.getScriptProperties().setProperty('evaluationWindowStart', formattedTime); // Save the formatted time
  Logger.log(`Evaluation window start time saved: ${formattedTime}`);
}

function getEvaluationWindowTimes() {
  const evaluationWindowStartStr = PropertiesService.getScriptProperties().getProperty('evaluationWindowStart');
  if (!evaluationWindowStartStr) {
    throw new Error('Evaluation window start time not found.');
  }
  const evaluationWindowStart = new Date(evaluationWindowStartStr);
  const evaluationWindowEnd = new Date(evaluationWindowStart.getTime() + EVALUATION_WINDOW_MINUTES * 60 * 1000);
  return { evaluationWindowStart, evaluationWindowEnd };
}

/**
 * Extracts the list of valid submission emails from the request log sheet within the submission time window.
 * Assumes the review log is for the current / correct period, and that the submitter email is in the first column for the request log sheet.
 * @param {Sheet} submissionSheet - The sheet containing submissions.
 * @returns {Array} - A list of valid submission emails within the submission time window.
 */
function getValidSubmissionEmails(submissionSheet) {
  Logger.log('Extracting valid submission emails.');

  const lastRow = submissionSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('No submission requests found.');
    return [];
  }

  const { submissionWindowStart, submissionWindowEnd } = getSubmissionWindowTimes();

  // Extract valid requests within the submission time window
  // only grabs the firts column, assumes that is the submitter's email address
  const validSubmitters = submissionSheet
    .getRange(2, 1, lastRow - 1, 6)
    .getValues()
    .filter((row) => {
      const submissionTimestamp = new Date(row[0]); // Assuming the first column is the timestamp
      // Check if the response is within the evaluation time window
      const isWithinWindow = submissionTimestamp >= submissionWindowStart && submissionTimestamp <= submissionWindowEnd;
      if (isWithinWindow) {
        Logger.log(`Valid submission found: ${row[1]} (Response time: ${row[0]})`);
        return true;
      }
      Logger.log(`Invalid submission found - outside sumbission window: ${row[1]} (Response time: ${row[0]})`);
      return false;
    })
    //TODO: get column for Email instead of assuming
    .map((row) => row[1]?.trim().toLowerCase()); // Extract and clean submitter emails

  Logger.log(`Valid submitters (within time window): ${validSubmitters.join(', ')}`);
  return validSubmitters;
}

/**
 * Extracts the list of valid evaluator emails from the evaluation responses sheet within the evaluation time window.
 * @param {Sheet} evaluationResponsesSheet - The sheet containing evaluation form responses.
 * @returns {Array} - A list of valid evaluator emails within the evaluation time window.
 */
function getValidEvaluationEmails(evaluationResponsesSheet) {
  Logger.log('Extracting valid evaluation emails.');

  const lastRow = evaluationResponsesSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('No evaluation responses found.');
    return [];
  }

  // Get the evaluation time window from the stored properties
  const { evaluationWindowStart, evaluationWindowEnd } = getEvaluationWindowTimes();

  // Extract valid responses within the evaluation time window
  //TODO: get column from 'Email' header
  const validEvaluators = evaluationResponsesSheet
    .getRange(2, 1, lastRow - 1, 3)
    .getValues()
    .filter((row) => {
      const responseTimestamp = new Date(row[0]); // Assuming the first column is the timestamp
      // Check if the response is within the evaluation time window
      const isWithinWindow = responseTimestamp >= evaluationWindowStart && responseTimestamp <= evaluationWindowEnd;
      if (isWithinWindow) {
        Logger.log(`Valid evaluator found: ${row[2]} (Response time: ${row[0]})`);
        return true;
      }
      return false;
    })
    .map((row) => row[2].trim().toLowerCase()); // Extracting the evaluator email

  Logger.log(`Valid evaluators (within time window): ${validEvaluators.join(', ')}`);
  return validEvaluators;
}

/**
 * Fetches and returns the submitter-evaluator assignments from the Review Log
 * @returns list of { submitter: [evaluators] }
 */
function getReviewLogAssignments() {
  Logger.log('Fetching submitter-evaluator assignments from Review Log.');

  const reviewLogSheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(
    REVIEW_LOG_SHEET_NAME
  );
  const lastRow = reviewLogSheet.getLastRow();

  if (lastRow < 2) {
    Logger.log('No data found in Review Log sheet.');
    return {};
  }

  const reviewData = reviewLogSheet.getRange(2, 1, lastRow - 1, 4).getValues(); // Fetches submitter and evaluator data

  // Structure the data as { submitter: [evaluators] }
  const assignments = {};
  reviewData.forEach((row) => {
    const submitterEmail = row[0];
    const evaluators = [row[1], row[2], row[3]].filter((email) => email); // Collect evaluators' emails
    assignments[submitterEmail] = evaluators;
  });

  Logger.log(`Review Log assignments: ${JSON.stringify(assignments)}`);
  return assignments;
}

/**
 * Returns a list of eligible ambassador emails from the Registry sheet.
 * Excludes ambassadors with "Expelled" status.
 */
function getEligibleAmbassadorsEmails() {
  try {
    Logger.log('Fetching eligible ambassador emails.');

    const registrySheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(
      REGISTRY_SHEET_NAME
    );
    if (!registrySheet) {
      Logger.log('Registry sheet not found.');
      return [];
    }

    const registryData = registrySheet.getRange(2, 1, registrySheet.getLastRow() - 1, 3).getValues(); // Columns: Email, Discord Handle, Status
    const eligibleEmails = registryData
      .filter((row) => !row[2].includes('Expelled')) // Exclude those marked as 'Expelled'
      .map((row) => row[0]); // Extract emails

    Logger.log(`Eligible emails (excluding 'Expelled'): ${JSON.stringify(eligibleEmails)}`);
    return eligibleEmails;
  } catch (error) {
    Logger.log(`Error in getEligibleAmbassadorsEmails: ${error}`);
    return [];
  }
}

//
//
//////////// DATE UTILITS

///**
// * Get the time zone of the script (all spreadsheets).
function getProjectTimeZone() {
  return Session.getScriptTimeZone(); // Using Project's time zone
}

/**
 * Get the formatted month name for a given date and time zone.
 * @param {Date} date - The date object.
 * @param {string} timeZone - The time zone for formatting.
 * @returns {string} - Formatted string for the month and year (e.g., "September 2024").
 */
function getMonthNameForDate(date, timeZone) {
  return Utilities.formatDate(date, timeZone, 'MMMM yyyy');
}

/**
 * A helper function to retrieve the first day of the previous month based on time zone.
 * This function returns a Date object that represents the first day of the previous month.
 * The time is set to 7:00 (feel free to adjust as needed) to match the time set in the previous columns.
 * @param {string} timeZone - The time zone of the table.
 * @returns {Date} - The date of the first day of the previous month.
 */
function getPreviousMonthDate(timeZone) {
  const now = new Date();

  // Get the current year and month in the time zone of the table
  const formattedYear = Utilities.formatDate(now, timeZone, 'yyyy');
  const formattedMonth = Utilities.formatDate(now, timeZone, 'MM');

  let prevMonth = parseInt(formattedMonth) - 1;
  let prevYear = parseInt(formattedYear);
  if (prevMonth === 0) {
    prevMonth = 12;
    prevYear -= 1;
  }

  // Create a Date object for the first day of the previous month at 7:00 UTC (to match the time of the previous columns)
  //TODO: fix timezone assumption
  const targetDate = new Date(Date.UTC(prevYear, prevMonth - 1, 1, 7, 0, 0, 0));
  Logger.log(`Calculated date of the previous month: ${targetDate} (ISO: ${targetDate.toISOString()})`);

  return targetDate;
}

/**
 * A helper function to retrieve the first day of the previous month based on time zone and a starting date.
 * This function returns a Date object that represents the first day of the month before the starting date.
 * The time is set to 7:00 (feel free to adjust as needed) to match the time set in the previous columns.
 * @param {string} timeZone - The time zone of the table.
 * @param {Date} startingDate - The starting date for the calculation.
 * @returns {Date} - The date of the first day of the previous month.
 */
function getStartOfPriorMonth(timeZone, startingDate) {
  const now = new Date();

  // Get the current year and month in the time zone of the table
  const formattedYear = Utilities.formatDate(startingDate, timeZone, 'yyyy');
  const formattedMonth = Utilities.formatDate(startingDate, timeZone, 'MM');

  let prevMonth = parseInt(formattedMonth) - 1;
  let prevYear = parseInt(formattedYear);
  if (prevMonth === 0) {
    prevMonth = 12;
    prevYear -= 1;
  }

  // Create a Date object for the first day of the previous month at 7:00 UTC (to match the time of the previous columns)
  //TODO: fix timezone assumption
  const targetDate = new Date(Date.UTC(prevYear, prevMonth - 1, 1, 7, 0, 0, 0));
  Logger.log(`Calculated date of the previous month: ${targetDate} (ISO: ${targetDate.toISOString()})`);

  return targetDate;
}

//
//
//
// ======= email-Discord Handle Converters =======

function getDiscordHandleFromEmail(email) {
  const registrySheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(REGISTRY_SHEET_NAME);
  const emailColumn = registrySheet
    .getRange(2, 1, registrySheet.getLastRow() - 1, 1)
    .getValues()
    .flat();
  const discordColumn = registrySheet
    .getRange(2, 2, registrySheet.getLastRow() - 1, 1)
    .getValues()
    .flat();

  const index = emailColumn.indexOf(email);
  return index !== -1 ? discordColumn[index] : null;
}

//    FORMS' TITLES
//
// Main function to update the form titles based on the current reporting month
function updateFormTitlesWithCurrentReportingMonth() {
  // Retrieve the reporting month in "MMMM yyyy" format, e.g., "August 2024"
  const reportingMonth = Utilities.formatDate(
    getPreviousMonthDate(Session.getScriptTimeZone()),
    Session.getScriptTimeZone(),
    'MMMM yyyy'
  );

  // Open each form by its ID
  const submissionForm = FormApp.openById(SUBMISSION_FORM_ID);
  const evaluationForm = FormApp.openById(EVALUATION_FORM_ID);

  // Define the titles based on the reporting month
  const newSubmissionTitle = `Your Contributions in ${reportingMonth}`;
  const newEvaluationTitle = `Submitter's ScoreCard - ${reportingMonth}`;

  // Update the form titles
  submissionForm.setTitle(newSubmissionTitle);
  evaluationForm.setTitle(newEvaluationTitle);

  Logger.log(`Updated Submission Form title to: ${newSubmissionTitle}`);
  Logger.log(`Updated Evaluation Form title to: ${newEvaluationTitle}`);
}

/////////////  INDEX UTILITIES for COLUMNS and SHEETS

function getColumnIndexByName(sheet, columnName) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  headers.forEach((header, index) => {
    Logger.log(`Header at column ${index + 1}: ${header}`);
  });
  return headers.indexOf(columnName) + 1;
}

/**
 * This function helps to determine if a column already exists for a given month to avoid duplicate columns in the "Overall score" sheet.
 * @param {Array} existingColumns - An array of existing column names.
 * @param {Date} targetDate - Target date (first day of the month).
 * @param {string} timeZone - Time zone of the table.
 * @returns {boolean} - Returns true if the column exists, otherwise false.
 */
function doesColumnExist(existingColumns, targetDate, timeZone) {
  const targetMonthName = Utilities.formatDate(targetDate, timeZone, 'MMMM yyyy');
  Logger.log(`Checking the existence of the column: "${targetMonthName}"`);
  return existingColumns.some((column) => {
    if (column instanceof Date) {
      const columnMonthName = Utilities.formatDate(column, timeZone, 'MMMM yyyy');
      return columnMonthName === targetMonthName;
    }
    return false;
  });
}

/**
 * Finds the index for inserting a new column after the last existing month column.
 * Check every column, including the first two, to make sure that the new column is added after the most recent existing column of the month.
 * @param {Array} existingColumns - Array of existing column names.
 * @returns {number} - Column index for insertion.
 */
function findInsertIndexForMonth(existingColumns) {
  let insertIndex = 1; // Starting with the first column

  existingColumns.forEach((columnName, index) => {
    if (isMonthColumn(columnName)) {
      insertIndex = index + 1; // Updating the index after the found month-column
      Logger.log(`Month-column found on the index ${index + 1}: "${columnName}"`);
    }
  });

  Logger.log(`Final insertion index: ${insertIndex}`);
  return insertIndex;
}

/**
 * Checks if the cell value is a month-column value.
 * This function is used to check if the cell value in "Overall score" is the date corresponding to the first day of the month.
 * @param {string|Date} columnValue - Field value.
 * @returns {boolean} - Returns true, if this is a month-column, otherwise false.
 */
function isMonthColumn(columnValue) {
  if (!(columnValue instanceof Date)) {
    Logger.log(`Value "${columnValue}" isn't Date object.`);
    return false;
  }

  const date = new Date(columnValue);
  const isFirstDay = date.getUTCDate() === 1;

  Logger.log(`Checking if "${columnValue}" is a month-column: isFirstDay=${isFirstDay}`);

  return isFirstDay;
}

// Function to explicitly get the "Form Responses" sheet based on the defined name
function getSubmissionFormResponseSheet() {
  Logger.log('Fetching "Submissions Form Responses" sheet.');
  const ss = SpreadsheetApp.openById(AMBASSADORS_SUBMISSIONS_SPREADSHEET_ID); // Open the "Ambassador Registry" spreadsheet
  const formResponseSheet = ss.getSheetByName(FORM_RESPONSES_SHEET_NAME); // Get the "Form Responses" sheet by name
  if (formResponseSheet) {
    Logger.log('"Form Responses" sheet found.');
  } else {
    Logger.log('Error: "Form Responses" sheet not found.');
  }
  return formResponseSheet;
}

/**
 * Function to find an index to insert a new month sheet before all existing month sheets.
 * This function goes through all the sheets in the table and determines where the new sheet should be inserted.
 * New sheets are always inserted before existing sheets that are months.
 * @param {Spreadsheet} scoresSpreadsheet - Ambassadors' Scores table.
 * @returns {number} - Index to insert a new month sheet.
 */
function findInsertIndexForMonthSheet(scoresSpreadsheet) {
  const sheets = scoresSpreadsheet.getSheets();
  let firstMonthSheetIndex = sheets.length; // Starting with the last index

  sheets.forEach((sheet, index) => {
    const sheetName = sheet.getName();
    if (isMonthSheet(sheetName)) {
      firstMonthSheetIndex = Math.min(firstMonthSheetIndex, index); // Updating the index to insert before the first month sheet
      Logger.log(`Found a month sheet on the index ${index + 1}: "${sheetName}"`);
    }
  });

  Logger.log(`Final month sheet insertion index: ${firstMonthSheetIndex}`);
  return firstMonthSheetIndex;
}

/**
 * Checks whether the sheet name is a month sheet (e.g. "September 2024").
 * This function is used to determine if a table sheet is a month sheet.
 * @param {string} sheetName - Name of the sheet.
 * @returns {boolean} - Returns true if it is a month sheet, otherwise false.
 */
function isMonthSheet(sheetName) {
  const monthYearPattern = /^[A-Za-z]+ \d{4}$/; // Template for checking the "Month Year" format
  return monthYearPattern.test(sheetName);
}

function findRowByDiscordHandle(discordHandle) {
  const overallScoresSheet = SpreadsheetApp.openById(AMBASSADORS_SCORES_SPREADSHEET_ID).getSheetByName(
    OVERALL_SCORE_SHEET_NAME
  );
  const handlesColumn = overallScoresSheet
    .getRange(2, 1, overallScoresSheet.getLastRow() - 1, 1)
    .getValues()
    .flat(); // Assuming Discord handle is in the first column

  const rowIndex = handlesColumn.findIndex((handle) => handle === discordHandle);

  // Since findIndex returns 0-based index, we add 2 to get the actual row in the sheet (1-based index, plus header row).
  return rowIndex !== -1 ? rowIndex + 2 : null;
}

//      TRIGGERS

// Function to delete all existing triggers
function deleteExistingTriggers() {
  try {
    Logger.log('Deleting existing triggers.');
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach((trigger) => {
      ScriptApp.deleteTrigger(trigger);
      Logger.log(`Deleted trigger: ${trigger.getHandlerFunction()}`);
    });
  } catch (error) {
    Logger.log(`Error in deleteExistingTriggers: ${error}`);
  }
}
// Lists all triggers:
function listAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((trigger) => {
    Logger.log(
      `Trigger function: ${trigger.getHandlerFunction()}, type: ${trigger.getEventType()}, next run: ${trigger.getTriggerSourceId()}`
    );
  });
}

//       Force re-authorization
/**
 * Triggers the Google Apps Script authorization dialog by attempting to access a protected service.
 * Useful when permissions need to be granted before using the script.
 */
function forceAuthorization() {
  try {
    // Access a protected service to prompt authorization. DriveApp is used as an example.
    DriveApp.getRootFolder();
    Logger.log('Authorization confirmed.');
  } catch (e) {
    Logger.log('Authorization required. Please reauthorize the script.');
  }
}

//    Helper functions to show a pop-up alert, or if running in debugger, write a log message
//    https://developers.google.com/apps-script/reference/base/button-set
/**
 * Shows an alert on the spreadsheet app ui with the given message and logs it.
 */
function alertAndLog(message) {
  Logger.log(message);
  try {
    // Attempt to get the UI
    const ui = SpreadsheetApp.getUi();
    ui.alert(message);
  } catch (e) {
    // If an error occurs, we don't have a UI; just catch and continue
  }
}

/**
 * Shows a prompt message and logs; if no UI is available, returns an affirmative response - YES or OK, depending on buttonSet.
 */
function promptAndLog(title, message, buttonSet = ButtonSet.OK) {
  Logger.log(`${title}: ${message}`);
  try {
    const ui = SpreadsheetApp.getUi();
    let response;

    switch (buttonSet) {
      case ButtonSet.OK:
        response = ui.alert(title, message, ui.ButtonSet.OK);
        return response == ui.Button.OK ? ButtonResponse.OK : null;
      case ButtonSet.OK_CANCEL:
        response = ui.alert(title, message, ui.ButtonSet.OK_CANCEL);
        return response == ui.Button.OK ? ButtonResponse.OK : ButtonResponse.CANCEL;
      case ButtonSet.YES_NO:
        response = ui.alert(title, message, ui.ButtonSet.YES_NO);
        return response == ui.Button.YES ? ButtonResponse.YES : ButtonResponse.NO;
      case ButtonSet.YES_NO_CANCEL:
        response = ui.alert(title, message, ui.ButtonSet.YES_NO_CANCEL);
        if (response == ui.Button.YES) return ButtonResponse.YES;
        if (response == ui.Button.NO) return ButtonResponse.NO;
        return ButtonResponse.CANCEL;
      default:
        Logger.log('Unknown button set');
        return null;
    }
  } catch (e) {
    Logger.log(message);
    Logger.log('UI not available; assuming YES or OK response');
    // Return default responses if no UI is available
    switch (buttonSet) {
      case ButtonSet.YES_NO:
      case ButtonSet.YES_NO_CANCEL:
        return ButtonResponse.YES;
      case ButtonSet.OK_CANCEL:
      default:
        return ButtonResponse.OK;
    }
  }
}
