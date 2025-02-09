// (( Global Variables ))
// Declare & initialize global variables; these will be updated by the setProductionVariables() or setTestVariables() functions

// testing constant will be used to load production vs. test values for the global variables
const testing = false; // Set to true for testing (logs instead of sending emails, uses test sheets and forms)
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
var AMBASSADOR_ID_COLUMN = '';
var AMBASSADOR_EMAIL_COLUMN = '';
var AMBASSADOR_DISCORD_HANDLE_COLUMN = '';
var AMBASSADOR_STATUS_COLUMN = '';
var GOOGLE_FORM_TIMESTAMP_COLUMN = '';
var GOOGLE_FORM_CONTRIBUTION_DETAILS_COLUMN = '';
var GOOGLE_FORM_CONTRIBUTION_LINKS_COLUMN = '';
var SUBM_FORM_USER_PROVIDED_EMAIL_COLUMN = '';
var EVAL_FORM_USER_PROVIDED_EMAIL_COLUMN = '';
var GOOGLE_FORM_REAL_EMAIL_COLUMN = '';
var GOOGLE_FORM_EVALUATION_HANDLE_COLUMN = '';
var GOOGLE_FORM_EVALUATION_GRADE_COLUMN = '';
var GOOGLE_FORM_EVALUATION_REMARKS_COLUMN = '';
var SCORE_PENALTY_POINTS_COLUMN = '';
var SCORE_AVERAGE_SCORE_COLUMN = '';
var SCORE_MAX_6M_PP_COLUMN = '';
var GRADE_SUBMITTER_COLUMN = '';
var GRADE_FINAL_SCORE_COLUMN = '';

// Sponsor Email (for notifications when ambassadors are expelled)
// set the actual values in EnvironmentVariables[Prod|Test].js
var SPONSOR_EMAIL = ''; // Sponsor's email

// Penalty Points threshold - if > or = this number for the past 6 months, ambassador will be expelled
var MAX_PENALTY_POINTS_TO_EXPEL = '';

// Color variables .The color hex string must be in lowercase!
// set the actual values in EnvironmentVariables[Prod|Test].js
var COLOR_MISSED_SUBMISSION = '';
var COLOR_MISSED_EVALUATION = '';
var COLOR_EXPELLED = '';
var COLOR_MISSED_SUBM_AND_EVAL = '';
//var COLOR_OLD_MISSED_SUBMISSION = '';

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
    .addItem('🔧️Batch process scores', 'batchProcessEvaluationResponses') //Re-runs score responses
    .addItem('🔧️Create/Sync Columns', 'syncRegistryColumnsToOverallScore') // creates Ambassador Status column in Overall score sheet; Syncs Ambassadors' Discord Handles and Ambassador Status columns between Registry and Overall score.
    .addItem('🔧️Check Emails in Submission Form responses', 'validateEmailsInSubmissionForm') // Checks completance of emails in 'Your Email Address' field of Submission Form. Recommended to run before Evaluation Requests to avoid errors caused by users' typo.
    .addItem('🔧️Delete Existing Triggers', 'deleteExistingTriggers') // Optional item
    .addItem('🔧️Force Authorization', 'forceAuthorization') // Authorization trigger
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
  //Logger.log('Color for missed submission: ' + COLOR_OLD_MISSED_SUBMISSION);
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

  // date parsing without time zone shifts
  const startDate = new Date(startDateStr);
  const timeZone = getProjectTimeZone();
  Logger.log(`Submission window start at: ${Utilities.formatDate(startDate, timeZone, 'yyyy-MM-dd HH:mm:ss z')}`);
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
 * Filters out emails not present in Registry.
 * Uses dynamic column indexing for the submitter email and timestamp columns.
 * @param {Sheet} submissionSheet - The sheet containing submissions.
 * @returns {Array} - A list of unique valid submission emails within the submission time window.
 */
function getValidSubmissionEmails(submissionSheet) {
  Logger.log('Extracting valid submission emails.');

  const registrySheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(REGISTRY_SHEET_NAME);
  const registryEmails = registrySheet
    .getRange(
      2,
      getRequiredColumnIndexByName(registrySheet, AMBASSADOR_EMAIL_COLUMN),
      registrySheet.getLastRow() - 1,
      1
    )
    .getValues()
    .flat()
    .map((email) => email.trim().toLowerCase());

  const lastRow = submissionSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('No submission requests found.');
    return [];
  }

  const { submissionWindowStart, submissionWindowEnd } = getSubmissionWindowTimes();
  const submitterEmailColumnIndex = getRequiredColumnIndexByName(submissionSheet, SUBM_FORM_USER_PROVIDED_EMAIL_COLUMN);
  const submitterTimestampColumnIndex = getRequiredColumnIndexByName(submissionSheet, GOOGLE_FORM_TIMESTAMP_COLUMN);

  Logger.log(`Submission time window: ${submissionWindowStart} - ${submissionWindowEnd}`);

  const validSubmitters = submissionSheet
    .getRange(2, 1, lastRow - 1, submissionSheet.getLastColumn())
    .getValues()
    .filter((row, index) => {
      const submissionTimestamp = new Date(row[submitterTimestampColumnIndex - 1]);
      const submitterEmail = row[submitterEmailColumnIndex - 1]?.trim().toLowerCase();

      if (!submissionTimestamp || !submitterEmail) {
        Logger.log(`Row ${index + 2}: Missing timestamp or email.`);
        return false;
      }

      const isWithinWindow = submissionTimestamp >= submissionWindowStart && submissionTimestamp <= submissionWindowEnd;

      if (!isWithinWindow) {
        Logger.log(
          `Row ${index + 2}: Submission at ${submissionTimestamp} outside time window ${submissionWindowStart} to ${submissionWindowEnd}.`
        );
        return false;
      }

      if (!registryEmails.includes(submitterEmail)) {
        Logger.log(`Row ${index + 2}: Email ${submitterEmail} not found in Registry. Skipping.`);
        return false;
      }

      return true;
    })
    .map((row) => row[submitterEmailColumnIndex - 1]?.trim().toLowerCase());

  // Remove duplicates
  const uniqueValidSubmitters = [...new Set(validSubmitters)];

  Logger.log(`Unique valid submitters: ${uniqueValidSubmitters.join(', ')}`);
  return uniqueValidSubmitters;
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

  // Dynamically retrieve column indices using headers
  const evalEmailColumnIndex = getRequiredColumnIndexByName(
    evaluationResponsesSheet,
    EVAL_FORM_USER_PROVIDED_EMAIL_COLUMN
  );
  const evalTimestampColumnIndex = getRequiredColumnIndexByName(evaluationResponsesSheet, GOOGLE_FORM_TIMESTAMP_COLUMN);

  // Extract valid responses within the evaluation time window
  const validEvaluators = evaluationResponsesSheet
    .getRange(2, 1, lastRow - 1, evaluationResponsesSheet.getLastColumn())
    .getValues()
    .filter((row, index) => {
      const responseTimestamp = new Date(row[evalTimestampColumnIndex - 1]); // Use correct column index for timestamp
      const evaluatorEmail = row[evalEmailColumnIndex - 1]?.trim().toLowerCase(); // Use correct column index for email

      if (!responseTimestamp || !evaluatorEmail) {
        Logger.log(`Row ${index + 2}: Missing timestamp or email.`);
        return false;
      }

      const isWithinWindow = responseTimestamp >= evaluationWindowStart && responseTimestamp <= evaluationWindowEnd;

      if (!isWithinWindow) {
        Logger.log(`Row ${index + 2}: Response outside evaluation time window.`);
        return false;
      }

      Logger.log(`Row ${index + 2}: Valid response found for email: ${evaluatorEmail}`);
      return true;
    })
    .map((row) => row[evalEmailColumnIndex - 1]?.trim().toLowerCase()); // Extract evaluator email

  Logger.log(`Valid evaluators (within time window): ${validEvaluators.join(', ')}`);
  return validEvaluators;
}

/**
 * Fetches and returns the submitter-evaluator assignments from the Review Log.
 * Dynamically determines column indices based on header names to avoid hardcoded indices.
 * @returns {Object} - A map of submitter emails to a list of evaluator emails.
 */
function getReviewLogAssignments() {
  Logger.log('Fetching submitter-evaluator assignments from Review Log.');

  const reviewLogSheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(
    REVIEW_LOG_SHEET_NAME
  );

  if (!reviewLogSheet) {
    Logger.log('Error: Review Log sheet not found.');
    return {};
  }

  const lastRow = reviewLogSheet.getLastRow();
  const lastColumn = reviewLogSheet.getLastColumn();

  if (lastRow < 2) {
    Logger.log('No data found in Review Log sheet.');
    return {};
  }

  // Get header row to determine column indices dynamically
  const headers = reviewLogSheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const submitterColIndex = getRequiredColumnIndexByName(reviewLogSheet, GRADE_SUBMITTER_COLUMN);
  const evaluatorCols = ['Reviewer 1', 'Reviewer 2', 'Reviewer 3'].map((header) => headers.indexOf(header) + 1);

  if (evaluatorCols.some((index) => index === 0)) {
    Logger.log('Error: Required columns (Submitter or Reviewer columns) not found in Review Log sheet.');
    return {};
  }

  // Fetch data for the entire sheet
  const reviewData = reviewLogSheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();

  // Structure the data as { submitter: [evaluators] }
  const assignments = {};
  reviewData.forEach((row) => {
    const submitterEmail = row[submitterColIndex - 1];
    const evaluators = evaluatorCols.map((colIndex) => row[colIndex - 1]).filter((email) => email); // Collect evaluators' emails
    if (submitterEmail) {
      assignments[submitterEmail] = evaluators;
    }
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
    const registryAmbassadorStatusColumnIndex = getRequiredColumnIndexByName(registrySheet, AMBASSADOR_STATUS_COLUMN);
    const registryAmbassadorEmailColumnIndex = getRequiredColumnIndexByName(registrySheet, AMBASSADOR_EMAIL_COLUMN);
    const registryData = registrySheet
      .getRange(2, 1, registrySheet.getLastRow() - 1, registrySheet.getLastColumn())
      .getValues(); // Columns: Email, Discord Handle, Status
    const eligibleEmails = registryData
      .filter((row) => !row[registryAmbassadorStatusColumnIndex - 1].includes('Expelled')) // Exclude those marked as 'Expelled'
      .map((row) => row[registryAmbassadorEmailColumnIndex - 1]); // Extract emails

    Logger.log(`Eligible emails (excluding 'Expelled'): ${JSON.stringify(eligibleEmails)}`);
    return eligibleEmails;
  } catch (error) {
    Logger.log(`Error in getEligibleAmbassadorsEmails: ${error}`);
    return [];
  }
}

//       DATE UTILITS

// Get the time zone of the script (all spreadsheets).
function getProjectTimeZone() {
  return Session.getScriptTimeZone(); // Using Project's time zone
}

/**
 * A helper function to retrieve the first day of the previous month based on the project time zone.
 * Returns a Date object that represents the first day of the previous month.
 * @returns {Date} - The date of the first day of the previous month.
 */
function getPreviousMonthDate() {
  Logger.log('Calculating the first day of the previous month.');

  // Retrieve the project time zone
  const timeZone = Session.getScriptTimeZone();
  Logger.log(`Using project time zone: ${timeZone}`);

  const now = new Date();

  // Get the current year and month in the project time zone
  const formattedYear = Utilities.formatDate(now, timeZone, 'yyyy');
  const formattedMonth = Utilities.formatDate(now, timeZone, 'MM');

  let prevMonth = parseInt(formattedMonth) - 1;
  let prevYear = parseInt(formattedYear);
  if (prevMonth === 0) {
    prevMonth = 12;
    prevYear -= 1;
  }

  // Create a Date object for the first day of the previous month at 00:00:00 (Pacific Time)
  const targetDate = new Date(prevYear, prevMonth - 1, 1, 0, 0, 0, 0);
  Logger.log(`Calculated date of the previous month: ${targetDate} (ISO: ${targetDate.toISOString()})`);

  return targetDate;
}

/**
 * Returns the first day of the month prior to the given date in the "submissionWindowStart" property.
 * Uses the time zone set in the Google Apps Script Project settings.
 * @returns {Date} - Date object representing the first day of the previous month.
 */
/**
 * Returns the first day of the previous month based on the submission window start date.
 * The date returned is the local time, ignoring UTC shifts.
 */
function getFirstDayOfReportingMonth() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const submissionWindowStart = scriptProperties.getProperty('submissionWindowStart');

    if (!submissionWindowStart) {
      throw new Error('submissionWindowStart is not defined in Script Properties.');
    }

    // Parsing the stored date string as a Date object in local time (not UTC)
    const startDate = new Date(submissionWindowStart);
    if (isNaN(startDate)) {
      throw new Error('Invalid date format in submissionWindowStart.');
    }

    const timeZone = getProjectTimeZone();
    Logger.log(`Using project time zone: ${timeZone}`);

    // Calculate the first day of the previous month with local time only
    const previousMonth = new Date(startDate.getFullYear(), startDate.getMonth() - 1, 1);
    Logger.log(
      `First day of the previous month (Local Time): ${Utilities.formatDate(previousMonth, timeZone, 'yyyy-MM-dd HH:mm:ss z')}`
    );

    return previousMonth;
  } catch (error) {
    Logger.log(`Error in getFirstDayOfReportingMonth: ${error.message}`);
    return null;
  }
}

// ======= email-Discord Handle Converters =======

function getDiscordHandleFromEmail(email) {
  const registrySheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(REGISTRY_SHEET_NAME);
  const emailColumnIndex = getRequiredColumnIndexByName(registrySheet, AMBASSADOR_EMAIL_COLUMN);
  const discordColumnIndex = getRequiredColumnIndexByName(registrySheet, AMBASSADOR_DISCORD_HANDLE_COLUMN);
  const emailColumn = registrySheet
    .getRange(2, emailColumnIndex, registrySheet.getLastRow() - 1, 1)
    .getValues()
    .flat();
  const discordColumn = registrySheet
    .getRange(2, discordColumnIndex, registrySheet.getLastRow() - 1, 1)
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
  const reportingMonth = Utilities.formatDate(getPreviousMonthDate(), Session.getScriptTimeZone(), 'MMMM yyyy');

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

//        INDEX UTILITIES for COLUMNS and SHEETS
/**
 * Wraps getColumnIndexByName -fFinds the column index for a given column name in the header row of the sheet, or throws exception.
 * @param {Sheet} sheet - The sheet to search.
 * @param {string} columnName - The name of the column to find.
 * @returns {number} The column index (1-based), or throws an error if not found
 */
function getRequiredColumnIndexByName(sheet, columnName) {
  const index = getColumnIndexByName(sheet, columnName);
  if (index == -1) {
    alertAndLog(`Expected Column "${columnName}" not found in sheet "${sheet.getName()}" header row: "${header}".`);
    throw new Error('Required column not found');
  }
  return index;
}

/**
 * Finds the column index for a given column name in the header row of the sheet.
 * @param {Sheet} sheet - The sheet to search.
 * @param {string} columnName - The name of the column to find.
 * @returns {number} The column index (1-based), or -1 if not found.
 */
function getColumnIndexByName(sheet, columnName) {
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const index = header.findIndex((h) => (h?.trim?.() ?? '') === columnName.trim());
  if (index == -1) {
    Logger.log(`Expected Column "${columnName}" not found in sheet "${sheet.getName()}" header row: "${header}".`);
    return -1;
  }
  return index + 1; // Convert to 1-based index
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
  const discordHandleColIndex = getRequiredColumnIndexByName(overallScoresSheet, AMBASSADOR_DISCORD_HANDLE_COLUMN);
  const handlesColumn = overallScoresSheet
    .getRange(2, discordHandleColIndex, overallScoresSheet.getLastRow() - 1, 1)
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

/**
 * Generates an MD5 hash from the given string.
 */
function generateMD5Hash(input) {
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, input);
  return rawHash.map((byte) => ('0' + (byte & 0xff).toString(16)).slice(-2)).join('');
}

/**
 *
 * @param {sheetName} string sheetName used for logging to explain what sheet is being evaluated
 * @param {headers[]} headers array of the entries in the first row from the sheet
 * @param {expectedHeaders[]} expectedHeaders array of the expected headers - simple string array of column names we expect to find
 * @returns true if all expectedHeaders are in headers, in the same order from left to right.
 */

function validateHeaders(sheetName, headers, expectedHeaders) {
  // Check if headers match the expected headers
  // only checks one-way - that expected headers are present in the sheet headers
  // does not care if there are more headers in the sheet than expected,
  // but expects the order matches from left (first column) to right.
  for (let i = 0; i < expectedHeaders.length; i++) {
    if (headers[i] !== expectedHeaders[i]) {
      alertAndLog(
        `Error: Unexpected column heading in ${sheetName} at index ${i}. Found: ${headers[i]}. Expected: ${expectedHeaders[i]}`
      );
      return false;
    }
  }
  return true;
}
