// ======= Global Variables =======
const testing = false;  // Set to true for testing (logs instead of sending emails); set to false to send emails

// Google Forms IDs:
const SUBMISSION_FORM_ID = '1SV5rJbzPv6BQgDZkC_xgrauWgoKPcEmtk3aKY6f4ZC8';  // ID for Submission form
const EVALUATION_FORM_ID = '15UXnrpOOoZPO7XCP2TV7mwezewHY6UIsYAU_W_aoMwo';  // ID for Evaluation form
const SUBMISSION_FORM_URL = 'https://forms.gle/beZrwuP9Zs1HvUY49';          // Submission Form URL for mailing
const EVALUATION_FORM_URL = 'https://forms.gle/kndReXQqXT6JyKX68';          // Evaluation Form URL for mailing

// Spreadsheets:
const AMBASSADOR_REGISTRY_SPREADSHEET_ID = '1AIMD61YKfk-JyP6Aia3UWke9pW15bPYTvWc1C46ofkU';  //"Ambassador Registry"
const AMBASSADORS_SCORES_SPREADSHEET_ID = '1RJzCo1FgGWkx0UCYhaY3SIOR0sl7uD6vCUFw55BX0iQ';   // "Ambassadors' Scores"

// "Ambassador Registry" spreadsheet sheets' names:
const REGISTRY_SHEET_NAME = 'Registry';                                  // Registry sheet in Ambassador Registry
const FORM_RESPONSES_SHEET_NAME = 'Form Responses 14';                   // Explicit name for 'Form Responses' sheet
const REVIEW_LOG_SHEET_NAME = 'Review Log';                              // Review Log sheet for evaluations
const CONFLICT_RESOLUTION_TEAM_SHEET_NAME = 'Conflict Resolution Team';  // CRT sheet

// Columns in 'Registry' sheet:
const AMBASSADOR_EMAIL_COLUMN = 'Ambassador Email Address';              // Email addresses column, column A
const AMBASSADOR_DISCORD_HANDLE_COLUMN = 'Ambassador Discord Handle';    // Discord handles column, column B

// Columns in 'Form Responses *' sheet
const RESPONSE_TIMESTAMP_COLUMN = 'Timestamp'; // Column A
const RESPONSE_EMAIL_COLUMN = 'Email Address'; // Column B
const RESPONSE_DISCORD_HANDLE_COLUMN = 'Your Discord Handle'; // Column C
const RESPONSE_CONTRIBUTIONS_COLUMN = "Dear Ambassador,\
Please add text to your contributions during the month of August, 2024:"; // Column D
const RESPONSE_LINKS_COLUMN = "Dear Ambassador,\
Please add links your contributions during the month of August, 2024:"; // Column E

// Columns in 'Review Log' sheet:
const REVIEW_LOG_SUBMITTER_COLUMN = 'Submitter';
const REVIEW_LOG_REVIEWER1_COLUMN = 'Reviewer 1';
const REVIEW_LOG_REVIEWER2_COLUMN = 'Reviewer 2';
const REVIEW_LOG_REVIEWER3_COLUMN = 'Reviewer 3';

// Columns in 'Conflict Resolution Team' sheet:
const CRT_SELECTION_DATE = 'Selection Date';
const CRT_AMBASSADOR1_COLUMN = 'Ambassador 1';
const CRT_AMBASSADOR2_COLUMN = 'Ambassador 2';
const CRT_AMBASSADOR3_COLUMN = 'Ambassador 3';
const CRT_AMBASSADOR4_COLUMN = 'Ambassador 4';
const CRT_AMBASSADOR5_COLUMN = 'Ambassador 5';

// "Ambassadors' Scores" spreadsheet sheets' names:
const OVERALL_SCORE_SHEET_NAME = 'Overall score';  // Overall score sheet in Ambassadors' Scores

// Columns in 'Overall score' sheet
const AMBASSADORS_SCORES_DISCORD_HANDLES_COLUMN = "Ambassadors' Discord Handles";
const AMBASSADORS_SCORES_EMAIL_COLUMN = 'E-mail';
const AMBASSADORS_SCORES_PENALTY_POINTS_COLUMN = 'Penalty Points';  // Column for penalty points in the Overall score sheet

// Columns in monthly sheets:
const SUBMITTER_COLUMN = 'Submitter';  // Submitter column in monthly sheets
const EVALUATIONS_COLUMN1 = 'Score or Evaluator\'s Email (if left unevaluated)-1'; // evaluations as they come
const EVALUATIONS_COLUMN2 = 'Score or Evaluator\'s Email (if left unevaluated)-2'; // evaluations as they come
const EVALUATIONS_COLUMN3 = 'Score or Evaluator\'s Email (if left unevaluated)-3'; // evaluations as they come
const FINAL_SCORE_COLUMN = 'Final Score';  // Arithmetic mean of grades columns

// Sponsor Email (for notifications when ambassadors are expelled)
const SPONSOR_EMAIL = "economicsilver@starmail.net";  // Sponsor's email

// Triggers and Delays 
// NOTE: Date replaced with Minutes for testing purposes
const SUBMISSION_WINDOW_DAYS = 7;
const SUBMISSION_WINDOW_DAYS_SPACE = 1;   // at how many days before Submission ends reminder will be activated
const EVALUATION_WINDOW_DAYS = 7;
const EVALUATION_WINDOW_DAYS_SPACE = 1;   // at how many days before Evaluation ends reminder will be activated

//		 ======= 	 Email Content Templates 	=======

// Request Submission Email Template
const REQUEST_SUBMISSION_EMAIL_TEMPLATE = `
Dear {AmbassadorName},
Please submit your deliverables for {Month} {Year} using the link below:
{SubmissionFormURL}
The deadline is {Deadline}.
Thank you,
Program Team
`;

// Request Evaluation Email Template
const REQUEST_EVALUATION_EMAIL_TEMPLATE = `
Dear {AmbassadorName},
Please review the following deliverables for the month of {Month} by {AmbassadorSubmitter}::
{SubmissionsList}
Assign a grade using the form:
{EvaluationFormURL}
Best regards,
Program Team
`;

// Reminder Email Template
const REMINDER_EMAIL_TEMPLATE = `
Hi there! Just a friendly reminder that we’re still waiting for your response to the Request for Submission/Evaluation. Please respond soon to avoid any penalties. Thank you!.
`;

// Penalty Warning Email Template
const PENALTY_WARNING_EMAIL_TEMPLATE = `
Dear Ambassador,
You have been assessed one penalty point for failing to meet deadlines. Further penalties may result in expulsion from the program. Please be vigillant.
`;

// Expulsion Email Template
const EXPULSION_EMAIL_TEMPLATE = `
Dear Ambassador,
We regret to inform you that you have been expelled from the program due to multiple missed deadlines.
`;
// Notify Upcoming Peer Review Email Template
const NOTIFY_UPCOMING_PEER_REVIEW = `
Dear Ambassador,
By this we notify you about upcoming Peer Review mailing, please be vigilant!
`;

//"CRT Selecting Notification" email template
const CRT_SELECTING_NOTIFICATION_TEMPLATE = `
Dear Ambassador, you have been selected for next Conflict Resolution Team ...
`;

// EXEMPTION FROM EVALUATION e-mail template
const EXEMPTION_FROM_EVALUATION_TEMPLATE = `
Dear Ambassador, you have been relieved of the obligation to evaluate your colleagues this month.
`;

// @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ SECTION 1 @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ //

// ======= Menu On Spreadsheet Open =======
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  Logger.log('Initializing menu on spreadsheet open.');
  
  ui.createMenu('Ambassador Program')
    .addItem('Request Submissions', 'requestSubmissions')  // Branch for Request Submissions
    .addItem('Request Evaluations', 'requestEvaluationsBranch')  // Branch for Request Evaluations
    .addItem('Process Responses', 'processResponses')   // Process Scores
    .addItem('Notify Upcoming Peer Review', 'notifyPeerReview')  // only for Notification about upcoming Peer Review
    .addItem('Selecting CRT', 'selectCRT')  // CRT
    .addItem('Force Authorization', 'forceAuthorization') // for Force Authorization
    .addItem('Delete Existing Triggers', 'deleteExistingTriggers') // for Force Authorization
    .addToUi();
  Logger.log('Menu initialized.');
}

// ======= Functions for Menu Item 1 =======

// Request Submissions: sends emails, sets up the new mailing, and reminder trigger
function requestSubmissions() {
  Logger.log('Request Submissions started.');

  const scoresSpreadsheet = SpreadsheetApp.openById(AMBASSADORS_SCORES_SPREADSHEET_ID); // Open the "Ambassadors' Scores" spreadsheet
  Logger.log('Opened "Ambassadors\' Scores" spreadsheet.');

  const spreadsheetTimeZone = scoresSpreadsheet.getSpreadsheetTimeZone(); // Get the spreadsheet's time zone
  Logger.log(`Spreadsheet time zone: ${spreadsheetTimeZone}`);

  const currentMailingDate = new Date(); // Get the current date and time as the start of mailing
  Logger.log(`Current mailing date: ${currentMailingDate}`);

  const registrySheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(REGISTRY_SHEET_NAME); // Open the "Registry" sheet
  Logger.log('Opened "Registry" sheet from "Ambassador Registry" spreadsheet.');

  const emails = registrySheet.getRange(2, 1, registrySheet.getLastRow() - 1, 1).getValues().flat(); // Fetch all emails from the Registry sheet
  Logger.log(`Fetched emails from "Registry" sheet: ${emails}`);

  const deliverableDate = getPreviousMonthDate(spreadsheetTimeZone); // Get the previous month's date
  Logger.log(`Deliverable date: ${deliverableDate}`);

  //const deadline = getExpectedResponseDate(spreadsheetTimeZone); // Get the expected response deadline (7 days from now)
  //Logger.log(`Expected response deadline: ${deadline}`);

  const month = Utilities.formatDate(deliverableDate, spreadsheetTimeZone, 'MMMM'); // Format the deliverable date to get the month name
  const year = Utilities.formatDate(deliverableDate, spreadsheetTimeZone, 'yyyy'); // Format the deliverable date to get the year
  Logger.log(`Formatted month and year: ${month} ${year}`);

  emails.forEach((email, index) => {
    const discordHandle = registrySheet.getRange(index + 2, 2).getValue(); // Fetch Discord Handle from column B
    Logger.log(`Fetched Discord handle for email ${email}: ${discordHandle}`);

    // Create the email message using the template and replace placeholders with actual values
    const message = REQUEST_SUBMISSION_EMAIL_TEMPLATE
      .replace('{AmbassadorName}', discordHandle)
      .replace('{Month}', month)
      .replace('{Year}', year)
      .replace('{SubmissionFormURL}', SUBMISSION_FORM_URL)
      .replace('{Deadline}', SUBMISSION_WINDOW_DAYS);
    Logger.log(`Email message created for ${email}: ${message}`);

    if (!testing) {
      MailApp.sendEmail(email, '☑️Request for Submission', message); // Send the email to the ambassador
      Logger.log(`Email sent to ${email}`);
    } else {
      Logger.log(`Testing mode: Submission request email logged for ${email}`);
    }
  });

  // Save the submission window start time
  const submissionWindowStart = new Date();
  PropertiesService.getScriptProperties().setProperty('submissionWindowStart', submissionWindowStart.toISOString());
  Logger.log(`Submission window start time saved: ${submissionWindowStart}`);

  // Set a trigger to check for non-respondents and send reminders after 6 days
  setupSubmissionReminderTrigger(submissionWindowStart);

  Logger.log('Request Submissions completed.');
}

// Calculates the expected response deadline, 7 days from the current submission date.
//function getExpectedResponseDate(timeZone) {
//  Logger.log('Calculating expected response date.');
//  const deadlineDate = new Date(); // Create a new Date object for the current date
//  deadlineDate.setDate(deadlineDate.getDate() + SUBMISSION_WINDOW_DAYS); // Add 7 days to the current date to calculate the deadline
//  Logger.log(`Deadline date set to: ${deadlineDate}`);
//  return Utilities.formatDate(deadlineDate, timeZone, 'MMMM d, yyyy'); // Format the deadline date as 'MMMM d, yyyy'
//}

// Setup a time-based trigger for sending reminders and save submission window start time
function setupSubmissionReminderTrigger(submissionStartTime) {
  Logger.log('Setting up submission reminder trigger.');

  const triggerDate = new Date(submissionStartTime);
  triggerDate.setMinutes(triggerDate.getMinutes() + SUBMISSION_WINDOW_DAYS - SUBMISSION_WINDOW_DAYS_SPACE); // Set trigger for 6 days after start for reminder
  Logger.log(`Trigger date for reminder set to: ${triggerDate}`);

  ScriptApp.newTrigger('checkNonRespondents')
    .timeBased()
    .at(triggerDate)
    .create();
  Logger.log('Reminder trigger created.');
}

// Check for non-respondents by comparing 'Form Responses' and 'Registry' sheets based on the submission window
function checkNonRespondents() {
  Logger.log('Checking for non-respondents.');

  const submissionWindowStartStr = PropertiesService.getScriptProperties().getProperty('submissionWindowStart');
  if (!submissionWindowStartStr) {
    Logger.log('Submission window start time not found, aborting checkNonRespondents.');
    return;
  }
  const submissionWindowStart = new Date(submissionWindowStartStr);
  const submissionWindowEnd = new Date(submissionWindowStart);
  submissionWindowEnd.setMinutes(submissionWindowStart.getMinutes() + SUBMISSION_WINDOW_DAYS);
  Logger.log(`Submission window is from ${submissionWindowStart} to ${submissionWindowEnd}`);

  const scoresSpreadsheet = SpreadsheetApp.openById(AMBASSADORS_SCORES_SPREADSHEET_ID);
  Logger.log('Opened "Ambassadors\' Scores" spreadsheet.');

  const spreadsheetTimeZone = scoresSpreadsheet.getSpreadsheetTimeZone();
  Logger.log(`Spreadsheet time zone: ${spreadsheetTimeZone}`);

  const registrySheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(REGISTRY_SHEET_NAME);
  Logger.log('Opened "Registry" sheet.');

  const formResponseSheet = getFormResponseSheet();
  if (!formResponseSheet) {
    Logger.log("Error: Form Response sheet not found.");
    return;
  }
  Logger.log('Form Response sheet found.');

  const allEmails = registrySheet.getRange(2, 1, registrySheet.getLastRow() - 1, 1).getValues().flat();
  Logger.log(`All emails from registry: ${allEmails}`);

  const responseData = formResponseSheet.getRange(2, 1, formResponseSheet.getLastRow() - 1, formResponseSheet.getLastColumn()).getValues();
  Logger.log(`Response data fetched from form: ${responseData.length} rows`);

  // Filter responses within submission window
  const validResponses = responseData.filter(row => {
    const timestamp = new Date(row[0]);
    return timestamp >= submissionWindowStart && timestamp <= submissionWindowEnd;
  });
  Logger.log(`Valid responses within submission window: ${validResponses.length}`);

  const respondedEmails = validResponses.map(row => row[1]); // Assuming email is in the second column
  Logger.log(`Responded emails: ${respondedEmails}`);

  // Find non-respondents by comparing 'Registry' emails with emails of valid responses
  const nonRespondents = allEmails.filter(email => !respondedEmails.includes(email));
  Logger.log(`Non-respondents: ${nonRespondents}`);

  // Send reminder emails to non-respondents
  sendReminderEmails(nonRespondents);
}

// Function to explicitly get the "Form Responses" sheet based on the defined name
function getFormResponseSheet() {
  Logger.log('Fetching "Form Responses" sheet.');
  const ss = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID); // Open the "Ambassador Registry" spreadsheet
  const formResponseSheet = ss.getSheetByName(FORM_RESPONSES_SHEET_NAME); // Get the "Form Responses" sheet by name
  if (formResponseSheet) {
    Logger.log('"Form Responses" sheet found.');
  } else {
    Logger.log('Error: "Form Responses" sheet not found.');
  }
  return formResponseSheet;
}

// Function for sending reminder emails with logging
function sendReminderEmails(nonRespondents) {
  Logger.log('Sending reminder emails.');
  const registrySheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(REGISTRY_SHEET_NAME); // Open the "Registry" sheet
  Logger.log('Opened "Registry" sheet.');

  if (!nonRespondents || nonRespondents.length === 0) {
    Logger.log('No non-respondents found.');
    return; // Exit if there are no non-respondents
  }

  nonRespondents.forEach(email => {
    const result = registrySheet.createTextFinder(email).findNext(); // Find the row with the given email
    if (result) {
      const row = result.getRow(); // Get the row number
      Logger.log(`Non-respondent found at row: ${row}`);
      const discordHandle = registrySheet.getRange(row, 2).getValue(); // Fetch Discord Handle from column B
      Logger.log(`Discord handle found for ${email}: ${discordHandle}`);
  
      // Create the reminder email message
      const message = REMINDER_EMAIL_TEMPLATE.replace('{AmbassadorName}', discordHandle);
  
      if (!testing) {
        MailApp.sendEmail(email, '⏰Reminder to Submit', message); // Send the reminder email
        Logger.log(`Reminder email sent to: ${email} (Discord: ${discordHandle})`);
      } else {
        Logger.log(`Testing mode: Reminder email logged for ${email}`);
      }
    } else {
      Logger.log(`Error: Could not find the ambassador with email ${email}`);
    }
  });
}

// Collect data from the form and check within a 7-day window
function collectSubmissions() {
  Logger.log('Collecting submissions.');

  const submissionWindowStart = getSubmissionWindowStart();
  if (!submissionWindowStart) {
    Logger.log('Submission window start not found, aborting submission collection.');
    return;
  }
  const submissionWindowEnd = new Date(submissionWindowStart);
  submissionWindowEnd.setMinutes(submissionWindowStart.getMinutes() + SUBMISSION_WINDOW_DAYS);
  Logger.log(`Submission window is from ${submissionWindowStart} to ${submissionWindowEnd}`);

  const scoresSpreadsheet = SpreadsheetApp.openById(AMBASSADORS_SCORES_SPREADSHEET_ID);
  Logger.log('Opened "Ambassadors\' Scores" spreadsheet.');

  const spreadsheetTimeZone = scoresSpreadsheet.getSpreadsheetTimeZone();
  Logger.log(`Spreadsheet time zone: ${spreadsheetTimeZone}`);

  const formResponseSheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(FORM_RESPONSES_SHEET_NAME);
  Logger.log('Opened "Form Responses" sheet.');

  const formData = formResponseSheet.getRange(2, 1, formResponseSheet.getLastRow() - 1, formResponseSheet.getLastColumn()).getValues();
  Logger.log(`Fetched form data: ${formData.length} rows`);

  formData.forEach(row => {
    const timestamp = new Date(row[0]);
    Logger.log(`Processing submission with timestamp: ${timestamp}`);
    if (timestamp >= submissionWindowStart && timestamp <= submissionWindowEnd) {
      Logger.log(`Valid submission received: ${timestamp}`);
      // Logic for processing valid responses goes here
    } else {
      Logger.log(`Submission outside window: ${timestamp}`);
    }
  });
}

// Get the start time of the submission window from script properties
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


// @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@  SECTION 2  @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ //


// Basic function for Request Evaluations menu item processing 
function requestEvaluationsBranch() {
  const evaluationStartTime = new Date(); // Capture start time
  PropertiesService.getScriptProperties().setProperty('evaluationStartTime', evaluationStartTime.toISOString());  
  
  // Step 1: Create a month sheet and column in the Overall Scores
  createMonthSheetAndOverallColumn();
  
  // Step 2: Generating the review matrix (submitters и evaluators)
  generateReviewMatrix();
  
  // Step 3: Sending evaluation requests
  sendEvaluationRequests();
  
  // Step 4: Filling out the Discord handle evaluators month sheet
  populateMonthSheetWithEvaluators();
  
  // Setting triggers
  setupEvaluationResponseTrigger();                 // Setting the onFormSubmit trigger to process evaluation responses
  setupEvaluationTriggers(evaluationStartTime);     // Setting triggers for reminders and closures
}


/**
 * Basic function to create a month sheet and the corresponding column in Overall Scores.
 * An approach for working with dates based on the time zone of the table is used.
 * This function is responsible for creating a new sheet for the current reporting month, and for adding a new column in the Overall Scores sheet.
 */
function createMonthSheetAndOverallColumn() {
  try {
    Logger.log('Execution started');

    // Opening "Ambassadors' Scores" spreadsheet and "Overall Scores" sheet
    const scoresSpreadsheet = SpreadsheetApp.openById(AMBASSADORS_SCORES_SPREADSHEET_ID);
    const overallScoreSheet = scoresSpreadsheet.getSheetByName(OVERALL_SCORE_SHEET_NAME);

    if (!overallScoreSheet) {
      Logger.log(`Sheet "${OVERALL_SCORE_SHEET_NAME}" isn't found in "Ambassadors' Scores" sheet.`);
      return;
    }

    // Get the time zone of the table
    const spreadsheetTimeZone = scoresSpreadsheet.getSpreadsheetTimeZone();
    Logger.log(`Time zone of the table: ${spreadsheetTimeZone}`);

    // Get the date of the first day of the previous month
    const deliverableMonthDate = getPreviousMonthDate(spreadsheetTimeZone);
    Logger.log(`Previous month date: ${deliverableMonthDate} (ISO: ${deliverableMonthDate.toISOString()})`);

    // Form the name of the month, e.g. 'September 2024'
    const deliverableMonthName = Utilities.formatDate(deliverableMonthDate, spreadsheetTimeZone, 'MMMM yyyy');
    Logger.log(`Month name: "${deliverableMonthName}"`);

    // Create a month sheet if it does not exist
    let monthSheet = scoresSpreadsheet.getSheetByName(deliverableMonthName);
    if (!monthSheet) {
      // Define the index for inserting a new sheet in front of all existing month sheets
      const sheetIndex = findInsertIndexForMonthSheet(scoresSpreadsheet);
      monthSheet = scoresSpreadsheet.insertSheet(deliverableMonthName, sheetIndex);
      // Adding headers to a new sheet
      monthSheet.getRange(1, 1).setValue(SUBMITTER_COLUMN);
      monthSheet.getRange(1, 2).setValue(EVALUATIONS_COLUMN1);
      monthSheet.getRange(1, 3).setValue(EVALUATIONS_COLUMN2);
      monthSheet.getRange(1, 4).setValue(EVALUATIONS_COLUMN3);
      monthSheet.getRange(1, 5).setValue(FINAL_SCORE_COLUMN);
      Logger.log(`New sheet created: "${deliverableMonthName}".`);
    } else {
      Logger.log(`Sheet "${deliverableMonthName}" already exists.`);
    }

    // Get the existing columns in the sheet "Overall Scores"
    const lastColumn = overallScoreSheet.getLastColumn();
    const existingColumns = overallScoreSheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    Logger.log(`Existing columns in "Overall Scores": ${existingColumns.join(', ')}`);

    // Check if there is a column for the previous month
    const columnExists = doesColumnExist(existingColumns, deliverableMonthDate, spreadsheetTimeZone);
    if (columnExists) {
      Logger.log(`Column for "${deliverableMonthName}" already exists in "Overall Scores". Skipping creation.`);
      return;
    }

    // Find the index for inserting a new column after the last existing month-column
    const insertIndex = findInsertIndexForMonth(existingColumns);
    Logger.log(`New column insertion index: ${insertIndex}`);

    // Insert a new column after the found index
    overallScoreSheet.insertColumnAfter(insertIndex);
    const newHeaderCell = overallScoreSheet.getRange(1, insertIndex + 1);
    Logger.log(`Insert date in cell: Column ${insertIndex + 1}, String 1`);

    // Set the value of the header cell as a Date object (with the same time as the existing columns)
    const safeDate = new Date(deliverableMonthDate.getTime());
    safeDate.setUTCHours(7, 0, 0, 0); // Set the time to 7:00 by UTC () to match the previous columns
    Logger.log(`Before setting the value: ${safeDate.toISOString()}`);
    newHeaderCell.setValue(safeDate);
    Logger.log(`The value of the cell is set: ${safeDate.toISOString()}`);

    // Set the cell format as 'MMMM yyyyy' to display only the month and year
    newHeaderCell.setNumberFormat('MMMM yyyy');
    Logger.log(`The cell format is set: 'MMMM yyyy'`);

    Logger.log(`Column for "${deliverableMonthName}" succefully added to "Overall Scores".`);

  } catch (error) {
    Logger.log(`Error in createMonthSheetAndOverallColumn: ${error}`);
  }
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
 * This function is used to determine if a table sheet is a month sheet...
 * @param {string} sheetName - SheetName.
 * @returns {boolean} - Returns true if it is a month sheet, otherwise false.
 */
function isMonthSheet(sheetName) {
  const monthYearPattern = /^[A-Za-z]+ \d{4}$/; // Template for checking the "Month Year" format
  return monthYearPattern.test(sheetName);
}

/**
 * A helper function to retrieve the first day of the previous month based on time zone.
 * This function returns a Date object that represents the first day of the previous month. 
 * The time is set to 14:00 (feel free to adjust as needed) to match the time set in the previous columns.
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

  // Create a Date object for the first day of the previous month at 14:00 UTC (to match the time of the previous columns)
  const targetDate = new Date(Date.UTC(prevYear, prevMonth - 1, 1, 14, 0, 0, 0));
  Logger.log(`Calculated date of the previous month: ${targetDate} (ISO: ${targetDate.toISOString()})`);

  return targetDate;
}

/**
 * Checks if a column exists for a given date.
 * This function helps to determine if a column already exists for a given month to avoid duplicate columns in the "Overall Scores" sheet.
 * @param {Array} existingColumns - An array of existing column names.
 * @param {Date} targetDate - Target date (first day of the month).
 * @param {string} timeZone - Time zone of the table.
 * @returns {boolean} - Returns true if the column exists, otherwise false.
 */
function doesColumnExist(existingColumns, targetDate, timeZone) {
  const targetMonthName = Utilities.formatDate(targetDate, timeZone, 'MMMM yyyy');
  Logger.log(`Checking the existence of the column: "${targetMonthName}"`);
  return existingColumns.some(column => {
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
 * This function is used to check if the cell value in "Overall Scores" is the date corresponding to the first day of the month. 
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


/**
 * Generates the review matrix by assigning evaluators to submitters.
 * Considers only valid submissions received within the 7-day submission window (adjust SUBMISSION_WINDOW_DAYS as needed).
 */
function generateReviewMatrix() {
  Logger.log('Starting generateReviewMatrix.');

  const registrySpreadsheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID);
  const registrySheet = registrySpreadsheet.getSheetByName(REGISTRY_SHEET_NAME);
  const formResponseSheet = registrySpreadsheet.getSheetByName(FORM_RESPONSES_SHEET_NAME);
  const reviewLogSheet = registrySpreadsheet.getSheetByName(REVIEW_LOG_SHEET_NAME);
  const spreadsheetTimeZone = registrySpreadsheet.getSpreadsheetTimeZone();

  Logger.log('Accessed Registry, Form Responses, and Review Log sheets.');

  // Retrieve submission window start time
  const submissionWindowStartStr = PropertiesService.getScriptProperties().getProperty('submissionWindowStart');
  if (!submissionWindowStartStr) {
    Logger.log('Submission window start time not found. Exiting generateReviewMatrix.');
    return;
  }
  const submissionWindowStart = new Date(submissionWindowStartStr);
  const submissionWindowEnd = new Date(submissionWindowStart);
  submissionWindowEnd.setMinutes(submissionWindowStart.getMinutes() + SUBMISSION_WINDOW_DAYS);

  Logger.log(`Submission window is from ${submissionWindowStart} to ${submissionWindowEnd}`);

  // Get responses from the Submission Form
  const lastRow = formResponseSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('No submissions found in Form Responses sheet.');
    return;
  }

  const responseData = formResponseSheet.getRange(2, 1, lastRow - 1, formResponseSheet.getLastColumn()).getValues();
  Logger.log(`Retrieved ${responseData.length} form responses.`);

  // Filter valid responses within the submission window
  const validResponses = responseData.filter(row => {
    const timestamp = new Date(row[0]);
    return timestamp >= submissionWindowStart && timestamp <= submissionWindowEnd;
  });

  Logger.log(`Found ${validResponses.length} valid submissions within the submission window.`);

  if (validResponses.length === 0) {
    Logger.log('No valid submissions found within the submission window.');
    return;
  }

  // Get email addresses of submitters
  const submittersEmails = validResponses.map(row => row[1]); // Assuming email is in the 2nd column
  Logger.log(`Submitters Emails: ${JSON.stringify(submittersEmails)}`);

  // Get all ambassadors' emails from Registry
  const allAmbassadorsEmails = registrySheet.getRange(2, 1, registrySheet.getLastRow() - 1, 1).getValues().flat();
  Logger.log(`All Ambassadors Emails: ${JSON.stringify(allAmbassadorsEmails)}`);

  // Initialize Review Log
  reviewLogSheet.clear();
  reviewLogSheet.getRange(1, 1).setValue('Submitter');
  reviewLogSheet.getRange(1, 2).setValue('Reviewer 1');
  reviewLogSheet.getRange(1, 3).setValue('Reviewer 2');
  reviewLogSheet.getRange(1, 4).setValue('Reviewer 3');
  Logger.log('Initialized Review Log sheet.');

  // Assign evaluators to each submitter
  const evaluatorQueue = [...allAmbassadorsEmails];
  evaluatorQueue.sort(() => Math.random() - 0.5); // Shuffle evaluators pool
  Logger.log('Shuffled evaluator pool.');

  const assignments = [];

  submittersEmails.forEach(submitter => {
    // Filter out the current submitter to prevent self-evaluation
    const availableEvaluators = evaluatorQueue.filter(email => email !== submitter);
    Logger.log(`Available evaluators for ${submitter}: ${JSON.stringify(availableEvaluators)}`);

    // Assign up to 3 unique evaluators
    const reviewers = [];
    for (let i = 0; i < 3; i++) {
      if (availableEvaluators.length === 0) {
        reviewers.push('Has No Evaluator');
      } else {
        const evaluator = availableEvaluators.shift(); // Take an evaluator from the front
        reviewers.push(evaluator);
        evaluatorQueue.push(evaluator); // Return evaluator to the queue
      }
    }

    // Add the assignment to the list
    assignments.push({ submitter, reviewers });
  });

  Logger.log(`Final evaluator assignments: ${JSON.stringify(assignments)}`);

  // Fill the Review Log sheet
  assignments.forEach((assignment, index) => {
    reviewLogSheet.getRange(index + 2, 1).setValue(assignment.submitter);
    assignment.reviewers.forEach((reviewer, idx) => {
      reviewLogSheet.getRange(index + 2, idx + 2).setValue(reviewer || 'Has No Evaluator');
    });
  });

  Logger.log('generateReviewMatrix completed.');
}

// ===== Sending Requests for Evaluation ======
function sendEvaluationRequests() {
  try {
    // Opening Review Log sheet
    const reviewLogSheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(REVIEW_LOG_SHEET_NAME);
    Logger.log(`Opened sheet: ${REVIEW_LOG_SHEET_NAME}`);

    // Getting spreadsheet Time Zone
    const scoresSheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID);
    const spreadsheetTimeZone = scoresSheet.getSpreadsheetTimeZone();
    Logger.log(`Spreadsheet TimeZone: ${spreadsheetTimeZone}`);

    const lastRow = reviewLogSheet.getLastRow();
    const lastColumn = reviewLogSheet.getLastColumn();
    Logger.log(`Review Log Sheet - Last Row: ${lastRow}, Last Column: ${lastColumn}`);

    // Проверяем, есть ли данные для обработки
    if (lastRow < 2) {
      Logger.log('No data in Review Log sheet. Exiting sendEvaluationRequests.');
      return;
    }

    // Get data from the sheet (starting from the second row, first through fourth columns)
    const reviewData = reviewLogSheet.getRange(2, 1, lastRow - 1, 4).getValues(); // Evaluations matrix
    Logger.log(`Retrieved ${reviewData.length} rows of data for the review.`);

    // Get the name of the previous month for sending requests
    const deliverableMonthDate = getPreviousMonthDate(spreadsheetTimeZone);
    const deliverableMonthName = Utilities.formatDate(deliverableMonthDate, spreadsheetTimeZone, 'MMMM yyyy');
    Logger.log(`Name of previous month: ${deliverableMonthName}`);

    reviewData.forEach((row, rowIndex) => {
      const submitterEmail = row[0]; // Email submitter'а
      const reviewersEmails = [row[1], row[2], row[3]].filter(email => email); // Email оценщиков

      Logger.log(`String processing ${rowIndex + 2}: Email Submitter: ${submitterEmail}, Email Evaluators: ${reviewersEmails.join(', ')}`);

      // Getting Discord handle submitter'а
      const submitterDiscordHandle = getDiscordHandleFromEmail(submitterEmail);
      Logger.log(`Discord Submitter: ${submitterDiscordHandle}`);

      // Getting the details of the contribution
      const contributionDetails = getContributionDetailsByEmail(submitterEmail, spreadsheetTimeZone);
      Logger.log(`Contribution details: ${contributionDetails}`);

      reviewersEmails.forEach(reviewerEmail => {
        try {
          const evaluatorDiscordHandle = getDiscordHandleFromEmail(reviewerEmail);
          Logger.log(`Discord Evaluator: ${evaluatorDiscordHandle}`);

          // Forming a message for evaluation
          const message = REQUEST_EVALUATION_EMAIL_TEMPLATE
            .replace('{AmbassadorName}', evaluatorDiscordHandle)
            .replace('{Month}', deliverableMonthName) // Use string name of the month
            .replace('{AmbassadorSubmitter}', submitterDiscordHandle)
            .replace('{SubmissionsList}', contributionDetails)
            .replace('{EvaluationFormURL}', EVALUATION_FORM_URL);

          if (!testing) {
            MailApp.sendEmail(reviewerEmail, '⚖️Request for Evaluation', message);
            Logger.log(`Evaluation request sent to ${reviewerEmail} (Discord: ${evaluatorDiscordHandle}) for submitter: ${submitterDiscordHandle}`);
          } else {
            Logger.log(`Test mode: The evaluation request must be sent to ${reviewerEmail}`);
          }
        } catch (error) {
          Logger.log(`Error sending evaluation request to ${reviewerEmail}: ${error}`);
        }
      });
    });
  } catch (error) {
    Logger.log(`Error in sendEvaluationRequests: ${error}`);
  }
}

// Function to get contribution details by email within the submission window
function getContributionDetailsByEmail(email) {
  const formResponseSheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(FORM_RESPONSES_SHEET_NAME);

  // Retrieve submission window start and end times
  const submissionWindowStartStr = PropertiesService.getScriptProperties().getProperty('submissionWindowStart');
  if (!submissionWindowStartStr) {
    Logger.log('Submission window start time not found.');
    return 'No contribution details found for this submitter.';
  }
  const submissionWindowStart = new Date(submissionWindowStartStr);
  const submissionWindowEnd = new Date(submissionWindowStart);
  submissionWindowEnd.setMinutes(submissionWindowStart.getMinutes() + SUBMISSION_WINDOW_DAYS);

  // Get form responses
  const formData = formResponseSheet.getRange(2, 1, formResponseSheet.getLastRow() - 1, formResponseSheet.getLastColumn()).getValues();

  // Find the corresponding response within the submission window
  for (let row of formData) {
    const timestamp = new Date(row[0]);
    if (timestamp >= submissionWindowStart && timestamp <= submissionWindowEnd) {
      const respondentEmail = row[1]; // Assuming Email is in the 2nd column
      if (respondentEmail === email) {
        const contributionText = row[3]; // Contribution details in the 4th column
        const contributionLinks = row[4]; // Links in the 5th column
        return `Contribution Details: ${contributionText}\nLinks: ${contributionLinks}`;
      }
    }
  }

  return 'No contribution details found for this submitter.';
}

// Function for filling in the month sheet with evaluators' Discord handles and grades
function populateMonthSheetWithEvaluators() {
  try {
    Logger.log('Populating month sheet with evaluators.');

    const scoresSheet = SpreadsheetApp.openById(AMBASSADORS_SCORES_SPREADSHEET_ID);
    const spreadsheetTimeZone = scoresSheet.getSpreadsheetTimeZone();
    const monthSheetName = Utilities.formatDate(getPreviousMonthDate(spreadsheetTimeZone), spreadsheetTimeZone, 'MMMM yyyy');
    const monthSheet = scoresSheet.getSheetByName(monthSheetName);

    if (!monthSheet) {
      Logger.log(`Month sheet ${monthSheetName} not found.`);
      return;
    }

    const reviewLogSheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(REVIEW_LOG_SHEET_NAME);
    const lastRow = reviewLogSheet.getLastRow();

    if (lastRow < 2) {
      Logger.log('No data found in Review Log sheet.');
      return;
    }

    const reviewData = reviewLogSheet.getRange(2, 1, lastRow - 1, 4).getValues();

    reviewData.forEach((row, index) => {
      const submitterEmail = row[0];
      const submitterDiscordHandle = getDiscordHandleFromEmail(submitterEmail);

      if (!submitterDiscordHandle) {
        Logger.log(`Discord handle not found for submitter email: ${submitterEmail}`);
        return;
      }

      Logger.log(`Row ${index + 2}: Submitter Discord Handle: ${submitterDiscordHandle}`);

      // Populate submitter handle in the Month Sheet
      monthSheet.getRange(index + 2, 1).setValue(submitterDiscordHandle);

      // Get evaluators' Discord handles
      const evaluatorsEmails = [row[1], row[2], row[3]].filter(email => email);
      const evaluatorsDiscordHandles = evaluatorsEmails.map(email => {
        const handle = getDiscordHandleFromEmail(email);
        if (!handle) {
          Logger.log(`Discord handle not found for evaluator email: ${email}`);
        } else {
          Logger.log(`Evaluator Discord Handle for ${email}: ${handle}`);
        }
        return handle || 'Unknown Evaluator';
      });

      // Populate evaluator columns
      evaluatorsDiscordHandles.forEach((handle, idx) => {
        monthSheet.getRange(index + 2, idx + 2).setValue(handle);
      });
    });

    Logger.log(`Evaluators populated in month sheet ${monthSheetName}.`);

  } catch (error) {
    Logger.log(`Error in populateMonthSheetWithEvaluators: ${error}`);
  }
}


/**
 * Function to process evaluation responses from the Google Form submissions.
 * It extracts evaluator's email, submitter's Discord handle, and the grade,
 * then updates the month sheet accordingly.
 */
function processEvaluationResponse(e) {
  try {
    Logger.log('processEvaluationResponse triggered.');

    const scoresSpreadsheet = SpreadsheetApp.openById(AMBASSADORS_SCORES_SPREADSHEET_ID);
    const spreadsheetTimeZone = scoresSpreadsheet.getSpreadsheetTimeZone();

    if (!e || !e.response) {
      Logger.log('Error: Event parameter is missing or does not have a response.');
      return;
    }

    // Extract the FormResponse object
    const formResponse = e.response;
    Logger.log('Form response received.');

    // Get the evaluator's email
    const evaluatorEmail = formResponse.getRespondentEmail();
    Logger.log(`Evaluator Email: ${evaluatorEmail}`);

    // Get the item responses
    const itemResponses = formResponse.getItemResponses();
    Logger.log(`Number of item responses: ${itemResponses.length}`);

    let submitterDiscordHandle = '';
    let grade = NaN;

    // Loop through the item responses to find the answers
    itemResponses.forEach(itemResponse => {
      const question = itemResponse.getItem().getTitle();
      const answer = itemResponse.getResponse();
      Logger.log(`Question: ${question}, Answer: ${answer}, Type of answer: ${typeof answer}`);

      // Match the question titles to extract the correct answers
      if (question === 'Discord handle of the ambassador you are evaluating?') {
        submitterDiscordHandle = String(answer).trim();
      } else if (question === 'Please assign a grade on a scale of 0 to 5') {
        const answerText = String(answer);
        Logger.log(`Grade answer text: ${answerText}`);

        const gradeMatch = answerText.match(/\d+/);
        if (gradeMatch) {
          grade = parseFloat(gradeMatch[0]);
        } else {
          Logger.log(`Unable to parse grade from answer: ${answerText}`);
        }
      }
    });

    Logger.log(`Submitter Discord Handle (provided): ${submitterDiscordHandle}`);
    Logger.log(`Grade: ${grade}`);

    if (!evaluatorEmail || !submitterDiscordHandle || isNaN(grade)) {
      Logger.log('Missing required data. Exiting processEvaluationResponse.');
      return;
    }

    // Get the evaluation start and end times
    const evaluationStartTimeStr = PropertiesService.getScriptProperties().getProperty('evaluationStartTime');
    if (!evaluationStartTimeStr) {
      Logger.log('Error: Evaluation start time is not defined.');
      return;
    }
    const evaluationStartTime = new Date(evaluationStartTimeStr);
    const evaluationEndTime = new Date(evaluationStartTime);
    evaluationEndTime.setUTCMinutes(evaluationStartTime.getMinutes() + EVALUATION_WINDOW_DAYS);

    // Get the form submission timestamp
    const timestamp = formResponse.getTimestamp();
    Logger.log(`Form submission timestamp: ${timestamp}`);

    // Check if the response is within the valid evaluation period
    if (timestamp < evaluationStartTime || timestamp > evaluationEndTime) {
      Logger.log('Response is outside the valid evaluation period. Ignoring.');
      return;
    }

    // Get evaluator's expected submitters from the Review Log sheet
    const expectedSubmitters = getExpectedSubmittersForEvaluator(evaluatorEmail);
    if (!expectedSubmitters || expectedSubmitters.length === 0) {
      Logger.log(`No expected submitters found for evaluator: ${evaluatorEmail}`);
      return;
    }
    Logger.log(`Expected submitters for evaluator ${evaluatorEmail}: ${expectedSubmitters.join(', ')}`);

    // Use bruteforceDiscordHandle to find the best match
    const correctedDiscordHandle = bruteforceDiscordHandle(submitterDiscordHandle, expectedSubmitters);

    if (!correctedDiscordHandle) {
      Logger.log(`Could not match Discord handle: ${submitterDiscordHandle} for evaluator: ${evaluatorEmail}`);
      return;
    }

    Logger.log(`Corrected Discord Handle: ${correctedDiscordHandle}`);
    submitterDiscordHandle = correctedDiscordHandle;

    // Get evaluator's Discord handle from email (necessary if evaluator's Discord handle is used in the month sheet)
    const evaluatorDiscordHandle = getDiscordHandleFromEmail(evaluatorEmail);
    if (!evaluatorDiscordHandle) {
      Logger.log(`Discord handle not found for evaluator email: ${evaluatorEmail}`);
      return;
    }
    Logger.log(`Evaluator Discord Handle: ${evaluatorDiscordHandle}`);

    // Get the month sheet to update grades
    const deliverableMonthDate = getPreviousMonthDate(spreadsheetTimeZone);
    const monthSheetName = Utilities.formatDate(deliverableMonthDate, spreadsheetTimeZone, 'MMMM yyyy');
    const monthSheet = scoresSpreadsheet.getSheetByName(monthSheetName);

    if (!monthSheet) {
      Logger.log(`Month sheet ${monthSheetName} not found.`);
      return;
    }

    // Find the row for the submitter in the month sheet
    const submitterFinder = monthSheet.createTextFinder(submitterDiscordHandle);
    const submitterCell = submitterFinder.findNext();
    if (!submitterCell) {
      Logger.log(`Submitter ${submitterDiscordHandle} not found in month sheet.`);
      return;
    }

    const row = submitterCell.getRow();
    Logger.log(`Submitter ${submitterDiscordHandle} found at row ${row}`);

    // Find the column for the evaluator and replace the Discord handle with the grade
    let gradeUpdated = false;
    for (let col = 2; col <= 4; col++) { // Evaluator columns are 2, 3, and 4
      const cellValue = monthSheet.getRange(row, col).getValue();
      Logger.log(`Checking cell at row ${row}, column ${col}: ${cellValue}`);
      if (cellValue === evaluatorDiscordHandle) {
        monthSheet.getRange(row, col).setValue(grade);
        Logger.log(`Updated grade for submitter ${submitterDiscordHandle} by evaluator ${evaluatorDiscordHandle}. Grade: ${grade}`);
        gradeUpdated = true;
        break;
      }
    }

    if (!gradeUpdated) {
      Logger.log(`Evaluator ${evaluatorDiscordHandle} not assigned to submitter ${submitterDiscordHandle} in month sheet.`);
    }

    // Update the final score for the submitter
    const gradesRange = monthSheet.getRange(row, 2, 1, 3); // Columns 2 to 4
    const grades = gradesRange.getValues()[0];
    const validGrades = grades.filter(value => typeof value === 'number' && !isNaN(value));

    if (validGrades.length > 0) {
      const finalScore = validGrades.reduce((sum, grade) => sum + grade, 0) / validGrades.length;
      monthSheet.getRange(row, 5).setValue(finalScore); // Column 5 is for final score
      Logger.log(`Final score updated for submitter ${submitterDiscordHandle}: ${finalScore}`);
    }

  } catch (error) {
    Logger.log(`Error in processEvaluationResponse: ${error}`);
  }
  updateOverallScores();
  Logger.log('processEvaluationResponse completed.');
}


// Function for setting up the evaluation response trigger with logging
function setupEvaluationResponseTrigger() {
  try {
    Logger.log('Setting up evaluation response trigger.');

    const form = FormApp.openById(EVALUATION_FORM_ID); // Open the form using the provided ID
    if (!form) {
      Logger.log('Error: Form not found with the given ID.');
      return;
    }

    // Delete existing triggers for this function to prevent duplicates
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'processEvaluationResponse') {
        ScriptApp.deleteTrigger(trigger);
        Logger.log('Deleted existing trigger: ' + trigger.getHandlerFunction());
      }
    });

    // Create a new trigger for form submissions
    ScriptApp.newTrigger('processEvaluationResponse')
      .forForm(form)
      .onFormSubmit()
      .create();

    Logger.log('Evaluation response trigger set up successfully.');
  } catch (error) {
    Logger.log(`Error in setupEvaluationResponseTrigger: ${error}`);
  }
}

// ======= Setup Evaluation Triggers Function =======
function setupEvaluationTriggers() {
  try {
    const evaluationStartTime = new Date();
    PropertiesService.getScriptProperties().setProperty('evaluationStartTime', evaluationStartTime.toISOString());
    Logger.log(`Evaluation start time set to: ${evaluationStartTime}`);

    setupEvaluationReminderTrigger(evaluationStartTime); // Set up the reminder trigger after 6 days
  } catch (error) {
    Logger.log(`Error in setupEvaluationTriggers: ${error}`);
  }
}


/**
 * Sets a reminder trigger to be sent 6 days after the start of the evaluation period.
 * @param {Date} evaluationStartTime - Evaluation start time.
 */
function setupEvaluationReminderTrigger(evaluationStartTime) {
  try {
    const reminderTime = new Date(evaluationStartTime);
    reminderTime.setMinutes(reminderTime.getMinutes() + 6); // Set reminder time to 6 days after evaluation start

    ScriptApp.newTrigger('sendEvaluationReminderEmails')
      .timeBased()
      .at(reminderTime)
      .create();

    Logger.log(`Reminder trigger for evaluation set for: ${reminderTime}`);
  } catch (error) {
    Logger.log(`Error in setupEvaluationReminderTrigger: ${error}`);
  }
}

/**
 * Function to update the Overall Scores sheet with the Final Scores from the month sheet.
 * Reads the Final Scores from the month sheet and updates the corresponding month column in the "Overall Scores" sheet for each ambassador.
 * Uses only Discord handles for mapping, as emails are not present in the Overall Scores sheet.
 * Since the order of ambassadors is fixed and matches the Registry sheet, the function matches ambassadors by Discord handle.
 */
function updateOverallScores() {
  Logger.log('Starting updateOverallScores.');

  // Open the "Ambassadors' Scores" spreadsheet and get the necessary sheets
  const scoresSpreadsheet = SpreadsheetApp.openById(AMBASSADORS_SCORES_SPREADSHEET_ID);
  Logger.log('Opened Ambassadors\' Scores spreadsheet.');

  const overallScoreSheet = scoresSpreadsheet.getSheetByName(OVERALL_SCORE_SHEET_NAME);
  Logger.log('Accessed Overall Scores sheet.');

  // Get the spreadsheet's time zone
  const spreadsheetTimeZone = scoresSpreadsheet.getSpreadsheetTimeZone();
  Logger.log(`Spreadsheet time zone: ${spreadsheetTimeZone}`);

  // Get the name of the current month sheet (e.g., "September 2024")
  const deliverableMonthDate = getPreviousMonthDate(spreadsheetTimeZone);
  Logger.log(`Deliverable month date: ${deliverableMonthDate}`);

  const monthSheetName = Utilities.formatDate(deliverableMonthDate, spreadsheetTimeZone, 'MMMM yyyy');
  Logger.log(`Month sheet name: ${monthSheetName}`);

  const monthSheet = scoresSpreadsheet.getSheetByName(monthSheetName);
  if (!monthSheet) {
    Logger.log(`Month sheet "${monthSheetName}" not found. Exiting updateOverallScores.`);
    return;
  }
  Logger.log(`Accessed month sheet: ${monthSheetName}`);

  // Find the column index for the current month in the "Overall Scores" sheet
  const lastColumn = overallScoreSheet.getLastColumn();
  Logger.log(`Last column in Overall Scores sheet: ${lastColumn}`);

  const headerValues = overallScoreSheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  Logger.log(`Header values: ${headerValues}`);

  let monthColumnIndex = null;
  for (let col = 1; col <= headerValues.length; col++) {
    const cellValue = headerValues[col - 1];
    Logger.log(`Checking column ${col}: ${cellValue}`);
    if (cellValue instanceof Date) {
      const cellMonthName = Utilities.formatDate(cellValue, spreadsheetTimeZone, 'MMMM yyyy');
      Logger.log(`Formatted date in column ${col}: ${cellMonthName}`);
      if (cellMonthName === monthSheetName) {
        monthColumnIndex = col;
        Logger.log(`Found month column "${monthSheetName}" at index ${monthColumnIndex}.`);
        break;
      }
    }
  }

  if (!monthColumnIndex) {
    Logger.log(`Month column "${monthSheetName}" not found in "Overall Scores" sheet. Exiting updateOverallScores.`);
    return;
  }

  // Create a mapping of Discord handles to row numbers in the "Overall Scores" sheet
  const overallScoreData = overallScoreSheet.getRange(2, 1, overallScoreSheet.getLastRow() - 1, 1).getValues(); // Only Discord handles
  Logger.log(`Retrieved Discord handles from Overall Scores sheet.`);

  const discordHandleToRowMap = {};
  for (let i = 0; i < overallScoreData.length; i++) {
    const discordHandle = overallScoreData[i][0];
    if (discordHandle) {
      discordHandleToRowMap[discordHandle.trim()] = i + 2; // Row numbers start from 2 (excluding header)
      Logger.log(`Mapped Discord handle "${discordHandle.trim()}" to row ${i + 2}`);
    }
  }
  Logger.log('Created mapping of Discord handles to rows in "Overall Scores" sheet.');

  // Iterate over the month sheet and update the "Overall Scores" sheet
  const monthData = monthSheet.getRange(2, 1, monthSheet.getLastRow() - 1, 5).getValues(); // Columns A to E
  Logger.log('Retrieved data from month sheet.');

  for (let i = 0; i < monthData.length; i++) {
    const row = monthData[i];
    const submitterDiscordHandle = row[0];
    const finalScore = row[4];

    Logger.log(`Processing row ${i + 2} in month sheet: Submitter "${submitterDiscordHandle}", Final Score: ${finalScore}`);

    if (!submitterDiscordHandle) {
      Logger.log(`Row ${i + 2}: Missing submitter Discord handle. Skipping.`);
      continue;
    }

    // No need to check if finalScore is a valid number as per your instruction

    const overallRow = discordHandleToRowMap[submitterDiscordHandle.trim()];
    if (!overallRow) {
      Logger.log(`Submitter "${submitterDiscordHandle}" not found in "Overall Scores" sheet. Skipping.`);
      continue;
    }

    // Update the final score in the corresponding month column
    overallScoreSheet.getRange(overallRow, monthColumnIndex).setValue(finalScore);
    Logger.log(`Updated final score for "${submitterDiscordHandle}" in "Overall Scores" sheet at row ${overallRow}, column ${monthColumnIndex}.`);
  }

  Logger.log('Completed updateOverallScores.');
}


// ----------------------------

// для получения Discord handle по email
function getDiscordHandleFromEmail(email) {
  Logger.log(`Looking up Discord handle for email: ${email}`);

  const registrySheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(REGISTRY_SHEET_NAME);
  const emailColumn = registrySheet.getRange(2, 1, registrySheet.getLastRow() - 1, 1).getValues().flat();
  const discordColumn = registrySheet.getRange(2, 2, registrySheet.getLastRow() - 1, 1).getValues().flat();

  const index = emailColumn.indexOf(email);
  if (index !== -1) {
    const discordHandle = discordColumn[index];
    Logger.log(`Found Discord handle: ${discordHandle} for email: ${email}`);
    return discordHandle;
  } else {
    Logger.log(`Discord handle not found for email: ${email}`);
    return null;
  }
}


function getExpectedSubmittersForEvaluator(evaluatorEmail) {
  const registrySpreadsheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID);
  const reviewLogSheet = registrySpreadsheet.getSheetByName(REVIEW_LOG_SHEET_NAME);

  const dataRange = reviewLogSheet.getDataRange();
  const data = dataRange.getValues();

  const expectedSubmitters = [];

  // Assuming the first row is the header
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const submitterEmail = row[0]; // Assuming submitter's email is in column A
    const evaluatorEmails = row.slice(1, 4); // Assuming evaluator emails are in columns B to D

    if (evaluatorEmails.includes(evaluatorEmail)) {
      // Get submitter's Discord handle from the registry
      const submitterDiscordHandle = getDiscordHandleFromEmail(submitterEmail);
      if (submitterDiscordHandle) {
        expectedSubmitters.push(submitterDiscordHandle.trim());
      }
    }
  }

  return expectedSubmitters;
}

function bruteforceDiscordHandle(providedHandle, expectedHandles) {
  let bestMatch = null;
  let lowestDistance = Infinity;

  providedHandle = providedHandle.toLowerCase();

  expectedHandles.forEach(expectedHandle => {
    const distance = levenshteinDistance(providedHandle, expectedHandle.toLowerCase());
    if (distance < lowestDistance) {
      lowestDistance = distance;
      bestMatch = expectedHandle;
    }
  });

  // Set a reasonable threshold for maximum acceptable distance
  const maxAcceptableDistance = 3; // Adjust as needed
  if (lowestDistance <= maxAcceptableDistance) {
    return bestMatch;
  } else {
    return null; // No suitable match found
  }
}

// Levenshtein Distance function remains the same

function levenshteinDistance(a, b) {
  const matrix = [];

  // Increment along the first column of each row
  let i;
  for (i = 0; i <= b.length; i++) {
    matrix[i] = [i];
  }

  // Increment each column in the first row
  let j;
  for (j = 0; j <= a.length; j++) {
    matrix[0][j] = j;
  }

  // Fill in the rest of the matrix
  for (i = 1; i <= b.length; i++) {
    for (j = 1; j <= a.length; j++) {
      if (b.charAt(i - 1) === a.charAt(j - 1)) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1, // substitution
          matrix[i][j - 1] + 1,     // insertion
          matrix[i - 1][j] + 1      // deletion
        );
      }
    }
  }

  return matrix[b.length][a.length];
}


// Sending reminder emails to those who didn't respond on Eval Form within 6 days
function sendEvaluationReminderEmails() {
  try {
    Logger.log('Starting to send evaluation reminder emails.');

    const formResponseSheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(FORM_RESPONSES_SHEET_NAME);
    const reviewLogSheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(REVIEW_LOG_SHEET_NAME);

    const lastRow = reviewLogSheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('No data found in Review Log sheet.');
      return;
    }

    // Get all evaluations (make a list of those who should have evaluated)
    const reviewData = reviewLogSheet.getRange(2, 1, lastRow - 1, 4).getValues();
    const allEvaluators = reviewData.flatMap(row => [row[1], row[2], row[3]]).filter(email => email); // Вытаскиваем всех оценщиков

    // Get the list of already responded evaluators from Form Responses
    const formResponses = formResponseSheet.getRange(2, 1, formResponseSheet.getLastRow() - 1, formResponseSheet.getLastColumn()).getValues();
    const respondedEvaluators = formResponses.map(row => row[1]); // Предполагается, что Email оценщика в 2-й колонке

    // Identifying those who didn't respond
    const nonRespondents = allEvaluators.filter(email => !respondedEvaluators.includes(email));
    Logger.log(`Non-respondents: ${JSON.stringify(nonRespondents)}`);

    // Use the function from branch 1 to send reminders
    sendReminderEmails(nonRespondents);
    Logger.log('Evaluation reminder emails sent successfully.');
  } catch (error) {
    Logger.log(`Error in sendEvaluationReminderEmails: ${error}`);
  }
}

// Function to prompt re-authorization
function forceAuthorization() {
  Logger.log('Prompting for authorization.');

  // Access services that require authorization
  const form = FormApp.openById(EVALUATION_FORM_ID);
  const formResponses = form.getResponses();
  formResponses.forEach(response => {
    const email = response.getRespondentEmail();
    Logger.log(`Evaluator Email: ${email}`);
  });

  const spreadsheet = SpreadsheetApp.openById(AMBASSADORS_SCORES_SPREADSHEET_ID);
  Logger.log(`Spreadsheet accessed: ${spreadsheet.getName()}`);
}

// ======= Delete Existing Triggers =======
function deleteExistingTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  
  triggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
    Logger.log('Deleted existing trigger: ' + trigger.getHandlerFunction());
  });
}