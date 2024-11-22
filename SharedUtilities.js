// (( Global Variables ))
// Declare & initialize global variables; these will be updated by the setProductionVariables() or setTestVariables() functions

// testing constant will be used to load production vs. test values for the global variables
const testing = true; // Set to true for testing (logs instead of sending emails, uses test sheets and forms)
// var SEND_EMAIL = false;// (DUPLICATE!) // Will control whether emails are sent - must be true for production; may be true or false for testing depending on testing needs.

// Provide the Id of the google sheet for the registry and scoreing sheets:
var AMBASSADOR_REGISTRY_SPREADSHEET_ID = ''; //"Ambassador Registry"
var AMBASSADORS_SCORES_SPREADSHEET_ID = ''; // "Ambassadors' Scores"
var AMBASSADORS_SUBMISSIONS_SPREADSHEET_ID = ''; // "Ambassador Submission Responses"
var EVALUATION_RESPONSES_SPREADSHEET_ID = ''; // "Evaluation Responses"

// Provide the Id and submission URL for the submission and evaluation google forms:
var SUBMISSION_FORM_ID = ''; // ID for Submission form
var EVALUATION_FORM_ID = ''; // ID for Evaluation form
var SUBMISSION_FORM_URL = ''; // Submission Form URL for mailing
var EVALUATION_FORM_URL = ''; // Evaluation Form URL for mailing

// Sheet names
var REGISTRY_SHEET_NAME = 'Registry';
var FORM_RESPONSES_SHEET_NAME = '';
var REVIEW_LOG_SHEET_NAME = 'Review Log';
var CONFLICT_RESOLUTION_TEAM_SHEET_NAME = 'Conflict Resolution Team';
var OVERALL_SCORE_SHEET_NAME = 'Overall score'; // Overall score sheet in Ambassadors' Scores
var EVAL_FORM_RESPONSES_SHEET_NAME = ''; // Evaluation Form responses sheet

// Columns in Registry Sheet:
var AMBASSADOR_EMAIL_COLUMN = 'Ambassador Email Address';
var AMBASSADOR_DISCORD_HANDLE_COLUMN = 'Ambassador Discord Handle';
var AMBASSADOR_STATUS_COLUMN = 'Ambassador Status';

// Sponsor Email (for notifications when ambassadors are expelled)
var SPONSOR_EMAIL = ''; // Sponsor's email

// Color variables
var COLOR_MISSED_SUBMISSION = '#f5eee6';
var COLOR_MISSED_EVALUATION = '#e6d6c1';
var COLOR_EXPELLED = '#FF0000';
var COLOR_MISSED_SUBM_AND_EVAL = '#ceae83';
var COLOR_OLD_MISSED_SUBMISSION = '#f5eee6';

// Triggers and Delays
// These values will set the due date and reminder schedule for Submissions and Evaluations.
// The Submission or Evaluation will be due after the relevant WINDOW_MINUTES,
// and each ambassador will receive a reminder after the relevant WINDOW_REMINDER_MINUTES.
// specifies as days * hours * minutes
var SUBMISSION_WINDOW_MINUTES = '';
var SUBMISSION_WINDOW_REMINDER_MINUTES = ''; // how many minutes after Submission Requests sent to remind
var EVALUATION_WINDOW_MINUTES = '';
var EVALUATION_WINDOW_REMINDER_MINUTES = ''; // how many minutes after Evaluation Requests sent to remind


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

// On Open Menu

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
    .addItem('🔧️Delete Existing Triggers', 'deleteExistingTriggers') // Optional item
    .addItem('🔧️Refresh Script State', 'refreshScriptState') // Add this for easy access
    .addToUi();
  Logger.log('Menu initialized.');
}

// ⚠️ functions to prevent Google's use of cached outdated variables when run from GUI.
function refreshScriptState() {
  Logger.log('Starting Refresh Script State');

  clearCache();
  refreshGlobalVariables();
  SpreadsheetApp.flush();
  Logger.log('Script state refreshed: cache cleared, variables refreshed, and flush completed.');
}

function clearCache() {
  const cache = CacheService.getScriptCache();
  cache.removeAll([]);
  Logger.log('Cache cleared.');
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

///////// Submission/Evaluation WINDOW TIME

function setSubmissionWindowStart(time) {
  PropertiesService.getScriptProperties().setProperty('submissionWindowStart', time.toISOString());
}
function setEvaluationWindowStart(time) {
  PropertiesService.getScriptProperties().setProperty('evaluationWindowStart', time.toISOString());
}
function getEvaluationWindowStart() {
  const timeStr = PropertiesService.getScriptProperties().getProperty('evaluationWindowStart');
  return timeStr ? new Date(timeStr) : null;
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

////// VALID RESPONSES

function setValidSubmissionResponses(emails) {
  PropertiesService.getScriptProperties().setProperty('validSubmissionResponses', JSON.stringify(emails));
}
function getValidSubmissionResponses() {
  const emailsStr = PropertiesService.getScriptProperties().getProperty('validSubmissionResponses');
  return emailsStr ? JSON.parse(emailsStr) : [];
}

function setValidEvaluationResponses(emails) {
  PropertiesService.getScriptProperties().setProperty('validEvaluationResponses', JSON.stringify(emails));
}
function getValidEvaluationResponses() {
  const emailsStr = PropertiesService.getScriptProperties().getProperty('validEvaluationResponses');
  return emailsStr ? JSON.parse(emailsStr) : [];
}
//function getReviewLogData() {
// const reviewLogSheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(REVIEW_LOG_SHEET_NAME);
//  return reviewLogSheet.getRange(2, 1, reviewLogSheet.getLastRow() - 1, reviewLogSheet.getLastColumn()).getValues();
//}

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
  const evaluationWindowStart = getEvaluationWindowStart();

  if (!evaluationWindowStart) {
    Logger.log('Error: Evaluation window start time not found.');
    return [];
  }

  const evaluationWindowEnd = new Date(evaluationWindowStart.getTime() + EVALUATION_WINDOW_MINUTES * 60 * 1000);

  Logger.log(`Evaluation window: ${evaluationWindowStart} - ${evaluationWindowEnd}`);

  // Extract valid responses within the evaluation time window
  const validEvaluators = evaluationResponsesSheet
    .getRange(2, 1, lastRow - 1, 2)
    .getValues()
    .filter((row) => {
      const responseTimestamp = new Date(row[0]); // Assuming the first column is the timestamp
      const email = row[1].trim().toLowerCase(); // Assuming the second column is the evaluator's email

      // Check if the response is within the evaluation time window
      const isWithinWindow = responseTimestamp >= evaluationWindowStart && responseTimestamp <= evaluationWindowEnd;

      if (isWithinWindow) {
        Logger.log(`Valid evaluator found: ${email} (Response time: ${responseTimestamp})`);
        return true;
      }
      return false;
    })
    .map((row) => row[1].trim().toLowerCase()); // Extracting the evaluator email

  Logger.log(`Valid evaluators (within time window): ${validEvaluators.join(', ')}`);
  return validEvaluators;
}

// Fetches and returns the submitter-evaluator assignments from the Review Log
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

//////////// DATE UTILITS

/**
 * Get the time zone of the given spreadsheet.
 * @param {Spreadsheet} spreadsheet - The Spreadsheet instance.
 * @returns {string} - Time zone of the spreadsheet.
 */
function getSpreadsheetTimeZone(spreadsheet) {
  return spreadsheet.getSpreadsheetTimeZone();
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
  const targetDate = new Date(Date.UTC(prevYear, prevMonth - 1, 1, 7, 0, 0, 0));
  Logger.log(`Calculated date of the previous month: ${targetDate} (ISO: ${targetDate.toISOString()})`);

  return targetDate;
}

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

////// TRIGGERS

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

//// Force re-authorization
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
