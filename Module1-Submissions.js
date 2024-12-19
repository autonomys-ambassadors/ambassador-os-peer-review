// Module1.gs

// Request Submissions: sends emails, sets up the new mailing, and reminder trigger
function requestSubmissionsModule() {
  Logger.log('Request Submissions Module started.');

  // activating OnForm submit trigger for detecting edited responses
  setupFormSubmitTrigger();
  Logger.log('Form submission trigger set up.');

  // Update form titles with the current reporting month
  updateFormTitlesWithCurrentReportingMonth();
  Logger.log('Form titles updated with the current reporting month.');

  // Get the universal spreadsheet time zone
  const spreadsheetTimeZone = getProjectTimeZone();
  Logger.log(`Spreadsheet time zone: ${spreadsheetTimeZone}`);

  const registrySheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(REGISTRY_SHEET_NAME); // Open the "Registry" sheet
  Logger.log('Opened "Registry" sheet from "Ambassador Registry" spreadsheet.');

  // Fetch data from Registry sheet (Emails and Status)
  const registryEmailColIndex = getColumnIndexByName(registrySheet, AMBASSADOR_EMAIL_COLUMN);
  const registryAmbassadorStatus = getColumnIndexByName(registrySheet, AMBASSADOR_STATUS_COLUMN);
  const registryAmbassadorDiscordHandle = getColumnIndexByName(registrySheet, AMBASSADOR_DISCORD_HANDLE_COLUMN);
  const registryData = registrySheet
    .getRange(2, 1, registrySheet.getLastRow() - 1, registrySheet.getLastColumn())
    .getValues(); // Fetch Emails, Discord Handles, and Status
  Logger.log(`Fetched data from "Registry" sheet: ${JSON.stringify(registryData)}`);

  // Filter out ambassadors with 'Expelled' in their status
  const eligibleEmails = registryData
    .filter((row) => !row[registryAmbassadorStatus - 1].includes('Expelled')) // Exclude expelled ambassadors
    .map((row) => [row[registryEmailColIndex - 1], row[registryAmbassadorDiscordHandle - 1]]); // Extract only emails
  Logger.log(`Eligible ambassadors emails: ${JSON.stringify(eligibleEmails)}`);

  // Get deliverable date (previous month date)
  const deliverableDate = getPreviousMonthDate(spreadsheetTimeZone); // Call from SharedUtilities.gs
  Logger.log(`Deliverable date: ${deliverableDate}`);

  const month = Utilities.formatDate(deliverableDate, spreadsheetTimeZone, 'MMMM'); // Format the deliverable date to get the month name
  const year = Utilities.formatDate(deliverableDate, spreadsheetTimeZone, 'yyyy'); // Format the deliverable date to get the year
  Logger.log(`Formatted month and year: ${month} ${year}`);

  // Calculate the exact deadline date based on submission window
  const submissionWindowStart = new Date();
  const submissionDeadline = new Date(submissionWindowStart.getTime() + SUBMISSION_WINDOW_MINUTES * 60 * 1000); // Convert minutes to milliseconds
  const submissionDeadlineDate = Utilities.formatDate(submissionDeadline, spreadsheetTimeZone, 'MMMM dd, yyyy');

  eligibleEmails.forEach((row) => {
    const email = row[0]; //
    const discordHandle = row[1]; // Get Discord Handle from Registry

    // Validating email
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/; // Simple email regex
    if (!emailRegex.test(email)) {
      const warningMessage = `Warning: Invalid or missing email for Discord Handle "${discordHandle}". Skipping.`;
      Logger.log(warningMessage);
      return; // Skip invalid emails
    }

    if (!discordHandle) {
      Logger.log(`Error: Discord handle not found for email: ${email}`);
      return; // Skip if Discord Handle is missing
    }

    // Composing email body
    const message = REQUEST_SUBMISSION_EMAIL_TEMPLATE.replace('{AmbassadorDiscordHandle}', discordHandle)
      .replace('{Month}', month)
      .replace('{Year}', year)
      .replace('{SubmissionFormURL}', SUBMISSION_FORM_URL)
      .replace('{SUBMISSION_DEADLINE_DATE}', submissionDeadlineDate); // Insert deadline date

    Logger.log(`Email message created for ${email} with Discord handle: ${discordHandle}`);

    // Email sending logic
    if (SEND_EMAIL) {
      try {
        MailApp.sendEmail({
          to: email,
          subject: '☑️Request for Submission',
          htmlBody: message, // Use htmlBody to send HTML email
        });
        Logger.log(`Email sent to ${email}`);
      } catch (error) {
        Logger.log(`Failed to send email to ${email}. Error: ${error}`);
      }
    } else {
      Logger.log(`Testing mode: Submission request email logged for ${email}`);
    }
  });

  // Save the submission window start time in Los Angeles time zone format
  setSubmissionWindowStart(submissionWindowStart);

  // Set a trigger to check for non-respondents and send reminders
  setupSubmissionReminderTrigger(submissionWindowStart);

  Logger.log('Request Submissions completed.');
}

/**
 * Sets up a trigger for form submission to handle new responses.
 * Ensures that only one trigger is active for the 'handleNewResponses' function.
 */
function setupFormSubmitTrigger() {
  Logger.log('Setting up form submission trigger.');

  // Remove any existing trigger for 'handleNewResponses'
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((trigger) => {
    if (trigger.getHandlerFunction() === 'handleNewResponses') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('Existing trigger for handleNewResponses removed.');
    }
  });

  // Create a new trigger for 'handleNewResponses'
  ScriptApp.newTrigger('handleNewResponses')
    .forSpreadsheet(AMBASSADORS_SUBMISSIONS_SPREADSHEET_ID)
    .onFormSubmit()
    .create();

  Logger.log('Form submission trigger created.');
}

/**
 * Handles new form submissions, ensuring only the latest response per user (real email) is kept.
 * Removes older responses for the same email within the submission window.
 * The real email collected by Google Forms is used instead of a user-inputted email field.
 */
function handleNewResponses(e) {
  Logger.log('Processing new form submission.');

  // Open the Form Responses sheet
  const formResponsesSheet = SpreadsheetApp.openById(AMBASSADORS_SUBMISSIONS_SPREADSHEET_ID).getSheetByName(FORM_RESPONSES_SHEET_NAME);
  if (!formResponsesSheet) {
    Logger.log('Error: Form Responses sheet not found.');
    return;
  }

  // Get the headers to find the indices of "Timestamp" and "Email Address" columns
  const headers = formResponsesSheet.getRange(1, 1, 1, formResponsesSheet.getLastColumn()).getValues()[0];
  const timestampColIndex = headers.indexOf(GOOGLE_FORM_TIMESTAMP_COLUMN) + 1; // Convert to 1-based index
  const realEmailColIndex = headers.indexOf(GOOGLE_FORM_REAL_EMAIL_COLUMN) + 1; // Automatically collected email column

  if (timestampColIndex === 0 || realEmailColIndex === 0) {
    Logger.log('Error: Required columns (Timestamp or Real Email) not found in Form Responses Sheet.');
    return;
  }

  // Access the newly submitted row via the event object 'e'
  const newRow = e.range.getRow();
  const newTimestamp = formResponsesSheet.getRange(newRow, timestampColIndex).getValue();
  const newRealEmail = formResponsesSheet.getRange(newRow, realEmailColIndex).getValue();

  if (!newRealEmail || !newTimestamp) {
    Logger.log('New response missing real email or timestamp. No action taken.');
    return;
  }

  const submissionDate = new Date(newTimestamp).toDateString();
  Logger.log(`New response from ${newRealEmail} on ${submissionDate} at row ${newRow}. Checking for older duplicates...`);

  // Loop through all rows except the new one to find duplicates
  const lastRow = formResponsesSheet.getLastRow();
  for (let row = lastRow; row >= 2; row--) {
    if (row === newRow) continue; // Skip the newly submitted row

    const rowTimestamp = formResponsesSheet.getRange(row, timestampColIndex).getValue();
    const rowRealEmail = formResponsesSheet.getRange(row, realEmailColIndex).getValue();

    if (!rowRealEmail || !rowTimestamp) continue;

    if (
      rowRealEmail.trim().toLowerCase() === newRealEmail.trim().toLowerCase() &&
      new Date(rowTimestamp).toDateString() === submissionDate
    ) {
      // Delete older responses from the same email and date
      formResponsesSheet.deleteRow(row);
      Logger.log(`Deleted older response from ${newRealEmail} on ${submissionDate} at row ${row}.`);
    }
  }

  Logger.log('Form responses updated. Only the latest response for this email and date is kept.');
}

// Function to set up submission reminder trigger
function setupSubmissionReminderTrigger(submissionStartTime) {
  Logger.log('Setting up submission reminder trigger.');

  // Remove existing triggers for 'checkNonRespondents'
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((trigger) => {
    if (trigger.getHandlerFunction() === 'checkNonRespondents') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('Existing reminder trigger removed.');
    }
  });

  // Calculate the time for the reminder
  const triggerDate = new Date(submissionStartTime.getTime() + SUBMISSION_WINDOW_REMINDER_MINUTES * 60 * 1000);

  if (isNaN(triggerDate.getTime())) {
    Logger.log('Invalid Date for trigger.');
    return;
  }

  // Create the reminder trigger
  ScriptApp.newTrigger('checkNonRespondents').timeBased().at(triggerDate).create();
  Logger.log('Reminder trigger created.');
}

// Check for non-respondents by comparing 'Form Responses' and 'Registry' sheets based on the submission window
function checkNonRespondents() {
  Logger.log('Checking for non-respondents.');

  // Retrieve submission window start time
  const submissionWindowStartStr = PropertiesService.getScriptProperties().getProperty('submissionWindowStart');
  if (!submissionWindowStartStr) {
    Logger.log('Submission window start time not found.');
    return;
  }

  // Calculate submission window start and end times
  const submissionWindowStart = new Date(submissionWindowStartStr);
  const submissionWindowEnd = new Date(submissionWindowStart.getTime() + SUBMISSION_WINDOW_MINUTES * 60 * 1000);

  // Open Registry and Form Responses sheets
  const registrySheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(REGISTRY_SHEET_NAME);
  const formResponseSheet = getSubmissionFormResponseSheet();

  Logger.log('Sheets successfully fetched.');

  // Get column indices for required headers
  const registryEmailColIndex = getColumnIndexByName(registrySheet, AMBASSADOR_EMAIL_COLUMN);
  const registryAmbassadorStatusColIndex = getColumnIndexByName(registrySheet, AMBASSADOR_STATUS_COLUMN);
  const responseEmailColIndex = getColumnIndexByName(formResponseSheet, SUBM_FORM_USER_PROVIDED_EMAIL_COLUMN);
  const responseTimestampColIndex = getColumnIndexByName(formResponseSheet, GOOGLE_FORM_TIMESTAMP_COLUMN);

  // Fetch registry data and filter eligible emails
  const registryData = registrySheet.getRange(2, 1, registrySheet.getLastRow() - 1, registrySheet.getLastColumn()).getValues();
  const eligibleEmails = registryData
    .filter((row) => !row[registryAmbassadorStatusColIndex - 1].includes('Expelled'))
    .map((row) => row[registryEmailColIndex - 1]);

  Logger.log(`Eligible emails: ${eligibleEmails}`);

  // Fetch form responses
  const responseData = formResponseSheet.getRange(2, 1, formResponseSheet.getLastRow() - 1, formResponseSheet.getLastColumn()).getValues();

  // Filter valid responses within submission window
  const validResponses = responseData.filter((row) => {
    const timestamp = new Date(row[responseTimestampColIndex - 1]);
    return timestamp >= submissionWindowStart && timestamp <= submissionWindowEnd;
  });

  const respondedEmails = validResponses.map((row) => row[responseEmailColIndex - 1]);
  Logger.log(`Responded emails: ${respondedEmails}`);

  // Identify non-respondents
  const nonRespondents = eligibleEmails.filter((email) => !respondedEmails.includes(email));
  Logger.log(`Non-respondents: ${nonRespondents}`);

  // Send reminders to non-respondents
  if (nonRespondents.length > 0) {
    sendReminderEmails(nonRespondents);
    Logger.log(`Reminders sent to ${nonRespondents.length} non-respondents.`);
  } else {
    Logger.log('No non-respondents found.');
  }
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

  // Dynamically fetch column indices
  const registryEmailColIndex = getColumnIndexByName(registrySheet, AMBASSADOR_EMAIL_COLUMN); // Email column index
  const registryDiscordHandleColIndex = getColumnIndexByName(registrySheet, AMBASSADOR_DISCORD_HANDLE_COLUMN); // Discord Handle column index

  // Validate column indices
  if (registryEmailColIndex === -1 || registryDiscordHandleColIndex === -1) {
    Logger.log('Error: One or more required columns not found in Registry sheet.');
    return;
  }

  nonRespondents.forEach((email) => {
    // Validate email format
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/; // Simple regex for validating email
    if (!emailRegex.test(email)) {
      Logger.log(`Warning: Invalid email "${email}". Skipping.`);
      return; // Skip invalid or empty email
    }

    // Find the row with the given email in the Registry
    const result = registrySheet.createTextFinder(email).findNext();
    if (result) {
      const row = result.getRow(); // Get the row number
      Logger.log(`Non-respondent found at row: ${row}`);

      // Fetch Discord Handle dynamically
      const discordHandle = registrySheet.getRange(row, registryDiscordHandleColIndex).getValue();
      Logger.log(`Discord handle found for ${email}: ${discordHandle}`);

      // Create the reminder email message
      const message = REMINDER_EMAIL_TEMPLATE.replace('{AmbassadorDiscordHandle}', discordHandle);

      // Send the email
      if (SEND_EMAIL) {
        try {
          MailApp.sendEmail(email, '🕚 Reminder to Submit', message); // Send the reminder email
          Logger.log(`Reminder email sent to: ${email} (Discord: ${discordHandle})`);
        } catch (error) {
          Logger.log(`Failed to send reminder email to ${email}. Error: ${error}`);
        }
      } else {
        Logger.log(`Testing mode: Reminder email logged for ${email}`);
      }
    } else {
      Logger.log(`Error: Could not find the ambassador with email ${email}`);
    }
  });
}
