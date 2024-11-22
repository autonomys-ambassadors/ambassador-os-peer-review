// Module1.gs

// Request Submissions: sends emails, sets up the new mailing, and reminder trigger
function requestSubmissionsModule() {
  Logger.log('Request Submissions Module started.');

  // Update form titles with the current reporting month
  updateFormTitlesWithCurrentReportingMonth();
  Logger.log('Form titles updated with the current reporting month.');

  const scoresSpreadsheet = SpreadsheetApp.openById(AMBASSADORS_SCORES_SPREADSHEET_ID); // Open the "Ambassadors' Scores" spreadsheet
  Logger.log('Opened "Ambassadors\' Scores" spreadsheet.');

  const spreadsheetTimeZone = scoresSpreadsheet.getSpreadsheetTimeZone(); // Get the spreadsheet's time zone
  Logger.log(`Spreadsheet time zone: ${spreadsheetTimeZone}`);

  const registrySheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(REGISTRY_SHEET_NAME); // Open the "Registry" sheet
  Logger.log('Opened "Registry" sheet from "Ambassador Registry" spreadsheet.');

  // Fetch data from Registry sheet (Emails and Status)
  const registryData = registrySheet.getRange(2, 1, registrySheet.getLastRow() - 1, 3).getValues(); // Fetch Emails, Discord Handles, and Status
  Logger.log(`Fetched data from "Registry" sheet: ${JSON.stringify(registryData)}`);

  // Filter out ambassadors with 'Expelled' in their status
  const eligibleEmails = registryData
    .filter(row => !row[2].includes('Expelled')) // Exclude expelled ambassadors
    .map(row => row[0]); // Extract only emails
  Logger.log(`Eligible ambassadors emails: ${JSON.stringify(eligibleEmails)}`);

  const deliverableDate = getPreviousMonthDate(spreadsheetTimeZone); // Call from SharedUtilities.gs
  Logger.log(`Deliverable date: ${deliverableDate}`);

  const month = Utilities.formatDate(deliverableDate, spreadsheetTimeZone, 'MMMM'); // Format the deliverable date to get the month name
  const year = Utilities.formatDate(deliverableDate, spreadsheetTimeZone, 'yyyy'); // Format the deliverable date to get the year
  Logger.log(`Formatted month and year: ${month} ${year}`);

  // Calculate the exact deadline date based on submission window (e.g., in minutes or days)
  const submissionWindowStart = new Date();
  const submissionDeadline = new Date(submissionWindowStart.getTime() + SUBMISSION_WINDOW_MINUTES); // Adjust as needed
  const submissionDeadlineDate = Utilities.formatDate(submissionDeadline, spreadsheetTimeZone, 'MMMM dd, yyyy');

  eligibleEmails.forEach((email) => {
    const discordHandle = getDiscordHandleFromEmail(email); // Use helper function to get Discord handle
    if (!discordHandle) {
      Logger.log(`Error: Discord handle not found for email: ${email}`);
      return; // Skip this email if no Discord handle found
    }

    // Prepare the email message with the deadline date
    const message = REQUEST_SUBMISSION_EMAIL_TEMPLATE.replace('{AmbassadorDiscordHandle}', discordHandle)
      .replace('{Month}', month)
      .replace('{Year}', year)
      .replace('{SubmissionFormURL}', SUBMISSION_FORM_URL)
      .replace('{SUBMISSION_DEADLINE_DATE}', submissionDeadlineDate); // Insert the calculated deadline date

    Logger.log(`Email message created for ${email} with Discord handle: ${discordHandle}`);

    if (SEND_EMAIL) {
      MailApp.sendEmail({
        to: email,
        subject: '☑️Request for Submission',
        htmlBody: message, // Use htmlBody to ensure clickable link
      });
      Logger.log(`Email sent to ${email}`);
    } else {
      if (!testing) {
        Logger.log(
          `WARNING: Production mode with email disabled. Submission request email logged but NOT SENT for ${email}`
        );
      } else {
        Logger.log(`Testing mode: Submission request email logged for ${email}`);
      }
    }
  });

  // Save the submission window start time
  PropertiesService.getScriptProperties().setProperty('submissionWindowStart', submissionWindowStart.toISOString());
  Logger.log(`Submission window start time saved: ${submissionWindowStart}`);

  // Set a trigger to check for non-respondents and send reminders
  setupSubmissionReminderTrigger(submissionWindowStart);

  Logger.log('Request Submissions completed.');
}

// Function to set up submission reminder trigger
function setupSubmissionReminderTrigger(submissionStartTime) {
  Logger.log('Setting up submission reminder trigger.');

  const triggerDate = new Date(submissionStartTime);
  triggerDate.setMinutes(triggerDate.getMinutes() + SUBMISSION_WINDOW_REMINDER_MINUTES); // Setup reminder trigger

  if (isNaN(triggerDate.getTime())) {
    Logger.log('Invalid Date for trigger.');
    return;
  }

  Logger.log(`Trigger date for reminder set to: ${triggerDate}`);
  ScriptApp.newTrigger('checkNonRespondents').timeBased().at(triggerDate).create();
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
  submissionWindowEnd.setMinutes(submissionWindowStart.getMinutes() + SUBMISSION_WINDOW_MINUTES);
  Logger.log(`Submission window is from ${submissionWindowStart} to ${submissionWindowEnd}`);

  const registrySheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(REGISTRY_SHEET_NAME);
  Logger.log('Opened "Registry" sheet.');

  const formResponseSheet = getSubmissionFormResponseSheet(); // Call from SharedUtilities.gs
  if (!formResponseSheet) {
    Logger.log('Error: Form Response sheet not found.');
    return;
  }
  Logger.log('Form Response sheet found.');

  // Fetch emails and statuses from Registry
  const registryData = registrySheet.getRange(2, 1, registrySheet.getLastRow() - 1, 3).getValues(); // Columns: Email, Discord, Status
  Logger.log(`Registry data fetched: ${registryData.length} rows`);

  // Filter eligible ambassadors (exclude those with 'Expelled' in status)
  const eligibleEmails = registryData.filter(row => !row[2].includes('Expelled')).map(row => row[0]);
  Logger.log(`Eligible emails (excluding Expelled): ${eligibleEmails}`);

  const responseData = formResponseSheet
    .getRange(2, 1, formResponseSheet.getLastRow() - 1, formResponseSheet.getLastColumn())
    .getValues();
  Logger.log(`Response data fetched from form: ${responseData.length} rows`);

  // Filter responses within submission window
  const validResponses = responseData.filter((row) => {
    const timestamp = new Date(row[0]);
    return timestamp >= submissionWindowStart && timestamp <= submissionWindowEnd;
  });
  Logger.log(`Valid responses within submission window: ${validResponses.length}`);

  const respondedEmails = validResponses.map((row) => row[1]); // Assuming email is in the second column
  Logger.log(`Responded emails: ${respondedEmails}`);

  // Find non-respondents among eligible ambassadors by comparing 'Registry' emails with emails of valid responses
  const nonRespondents = eligibleEmails.filter((email) => !respondedEmails.includes(email));
  Logger.log(`Non-respondents (eligible only): ${nonRespondents}`);

  // Send reminder emails to non-respondents
  sendReminderEmails(nonRespondents); // Call from SharedUtilities.gs
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

  nonRespondents.forEach((email) => {
    const result = registrySheet.createTextFinder(email).findNext(); // Find the row with the given email
    if (result) {
      const row = result.getRow(); // Get the row number
      Logger.log(`Non-respondent found at row: ${row}`);
      const discordHandle = registrySheet.getRange(row, 2).getValue(); // Fetch Discord Handle from column B
      Logger.log(`Discord handle found for ${email}: ${discordHandle}`);

      // Create the reminder email message
      const message = REMINDER_EMAIL_TEMPLATE.replace('{AmbassadorDiscordHandle}', discordHandle);

      if (SEND_EMAIL) {
        MailApp.sendEmail(email, '🕚Reminder to Submit', message); // Send the reminder email
        Logger.log(`Reminder email sent to: ${email} (Discord: ${discordHandle})`);
      } else {
        if (!testing) {
          Logger.log(`WARNING: Production mode with email disabled. Reminder email logged but NOT SENT for ${email}`);
        } else {
          Logger.log(`Testing mode: Reminder email logged for ${email}`);
        }
      }
    } else {
      Logger.log(`Error: Could not find the ambassador with email ${email}`);
    }
  });
}
