// Module1.gs

// Request Submissions: sends emails, sets up the new mailing, and reminder trigger
// Request Submissions: sends emails, sets up the new mailing, and reminder trigger
function requestSubmissionsModule() {
  Logger.log('Request Submissions Module started.');

  // Update form titles with the current reporting month
  updateFormTitlesWithCurrentReportingMonth();
  Logger.log('Form titles updated with the current reporting month.');

  // Get the universal spreadsheet time zone
  const spreadsheetTimeZone = getProjectTimeZone();
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

  eligibleEmails.forEach((email, index) => {
    const discordHandle = registryData[index][1]; // Get Discord Handle from Registry

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

// Function to set up submission reminder trigger
function setupSubmissionReminderTrigger(submissionStartTime) {
  Logger.log('Setting up submission reminder trigger.');

  // Calculate reminder trigger time
  const triggerDate = new Date(submissionStartTime.getTime() + SUBMISSION_WINDOW_REMINDER_MINUTES * 60 * 1000);

  if (isNaN(triggerDate.getTime())) {
    Logger.log('Invalid Date for trigger.');
    return;
  }

  // Log trigger details
  const spreadsheetTimeZone = getProjectTimeZone(); // Updated function name
  Logger.log(`Trigger date for reminder set to: ${Utilities.formatDate(triggerDate, spreadsheetTimeZone, 'yyyy-MM-dd HH:mm:ss z')}`);

  // Calculate and log submission window
  const submissionWindowEnd = new Date(submissionStartTime.getTime() + SUBMISSION_WINDOW_MINUTES * 60 * 1000);
  Logger.log(`Submission window is from ${Utilities.formatDate(submissionStartTime, spreadsheetTimeZone, 'yyyy-MM-dd HH:mm:ss z')} to ${Utilities.formatDate(submissionWindowEnd, spreadsheetTimeZone, 'yyyy-MM-dd HH:mm:ss z')}`);

  // Create a time-based trigger for checking non-respondents
  ScriptApp.newTrigger('checkNonRespondents').timeBased().at(triggerDate).create();
  Logger.log('Reminder trigger created.');
}

// Check for non-respondents by comparing 'Form Responses' and 'Registry' sheets based on the submission window
function checkNonRespondents() {
  Logger.log('Checking for non-respondents.');

  // Retrieve the submission window start time
  const submissionWindowStartStr = PropertiesService.getScriptProperties().getProperty('submissionWindowStart');
  if (!submissionWindowStartStr) {
    Logger.log('Submission window start time not found, aborting checkNonRespondents.');
    return;
  }
  
  // Convert submission window start to Date and calculate window end
  const submissionWindowStart = new Date(submissionWindowStartStr);
  const submissionWindowEnd = new Date(submissionWindowStart.getTime() + SUBMISSION_WINDOW_MINUTES * 60 * 1000);

  // Get project time zone for consistent logging
  const spreadsheetTimeZone = getProjectTimeZone();
  Logger.log(`Submission window is from ${Utilities.formatDate(submissionWindowStart, spreadsheetTimeZone, 'yyyy-MM-dd HH:mm:ss z')} to ${Utilities.formatDate(submissionWindowEnd, spreadsheetTimeZone, 'yyyy-MM-dd HH:mm:ss z')}`);

  // Open Registry sheet
  const registrySheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(REGISTRY_SHEET_NAME);
  if (!registrySheet) {
    Logger.log('Error: Registry sheet not found.');
    return;
  }
  Logger.log('Opened "Registry" sheet.');

  const formResponseSheet = getSubmissionFormResponseSheet(); // Call from SharedUtilities.gs
  if (!formResponseSheet) {
    Logger.log('Error: Form Response sheet not found.');
    return;
  }
  Logger.log('Form Response sheet found.');

  // Fetch data from Registry and filter eligible ambassadors
  const registryData = registrySheet.getRange(2, 1, registrySheet.getLastRow() - 1, 3).getValues(); // Columns: Email, Discord, Status
  Logger.log(`Registry data fetched: ${registryData.length} rows`);
  
  const eligibleEmails = registryData.filter(row => !row[2].includes('Expelled')).map(row => row[0]);
  Logger.log(`Eligible emails (excluding Expelled): ${eligibleEmails}`);

  // Fetch response data from Form Responses
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

  // Find non-respondents
  const nonRespondents = eligibleEmails.filter((email) => !respondedEmails.includes(email));
  Logger.log(`Non-respondents (eligible only): ${nonRespondents}`);

  // Send reminder emails to non-respondents
  if (nonRespondents.length > 0) {
    sendReminderEmails(nonRespondents); // Call from SharedUtilities.gs
    Logger.log(`Reminder emails sent to ${nonRespondents.length} non-respondents.`);
  } else {
    Logger.log('No non-respondents found. No reminder emails sent.');
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

  nonRespondents.forEach((email, index) => {
    // Validating emails
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/; // Simple regex for validating email
    if (!emailRegex.test(email)) {
      const discordHandle = registrySheet.getRange(index + 2, 2).getValue(); // getting Discord Handle from Registry
      Logger.log(`Warning: Invalid or missing email for Discord Handle "${discordHandle}". Skipping.`);
      return; // skipping incorrect or empty email
    }

    // Finding row with given email in Registry
    const result = registrySheet.createTextFinder(email).findNext();
    if (result) {
      const row = result.getRow(); // Get the row number
      Logger.log(`Non-respondent found at row: ${row}`);
      const discordHandle = registrySheet.getRange(row, 2).getValue(); // Fetch Discord Handle from column B
      Logger.log(`Discord handle found for ${email}: ${discordHandle}`);

      // Create the reminder email message
      const message = REMINDER_EMAIL_TEMPLATE.replace('{AmbassadorDiscordHandle}', discordHandle);

      if (SEND_EMAIL) {
        try {
          MailApp.sendEmail(email, '🕚Reminder to Submit', message); // Send the reminder email
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
