function requestMonthlySubmissions() {
  const ui = SpreadsheetApp.getUi();

  // Fetch the form responses spreadsheet
  Logger.log('Finding the Monthly Submission from Responses spreadsheet...');
  const formResponseSheet = getSubmissionResponsesSheet();

  // Get the index of the timestamp column
  const timestampColumnIndex = getRequiredColumnIndexByName(formResponseSheet, GOOGLE_FORM_TIMESTAMP_COLUMN);
  Logger.log(`Timestamp column index: ${timestampColumnIndex}`);

  // Fetch all the responses
  const formData = formResponseSheet
    .getRange(2, 1, formResponseSheet.getLastRow() - 1, formResponseSheet.getLastColumn())
    .getValues();

  // Find the latest submission
  let latestSubmission = null;
  Logger.log('Finding the latest submission...');
  for (let row of formData) {
    const timestamp = new Date(row[timestampColumnIndex - 1]);
    if (!latestSubmission || timestamp > latestSubmission) {
      latestSubmission = timestamp;
    }
  }

  if (!latestSubmission) {
    Logger.log('Error: No submissions found.');
    ui.alert('Error', 'Unable to find previous submissions. Please set it up manually.', ui.ButtonSet.OK);
    return;
  }

  // Calculate the next month
  Logger.log(`Latest submission date: ${latestSubmission}`);
  Logger.log('Calculating the next month...');
  const nextMonth = new Date(latestSubmission);
  nextMonth.setMonth(nextMonth.getMonth() + 1);

  const month = Utilities.formatDate(nextMonth, getProjectTimeZone(), 'MMMM');
  const year = nextMonth.getFullYear();

  Logger.log(`Next month: ${month}, Year: ${year}`);
  Logger.log('Asking for confirmation to send submission requests...');
  const response = ui.alert(
    'Confirm Submission',
    `Do you want to request submissions for ${month} ${year}?`,
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    Logger.log('User confirmed to send submission requests.');
    processFormData({ month, year });
  } else {
    //case user responded NO
    Logger.log('User wants to specify a different month for the submission request.');
    const form = HtmlService.createHtmlOutputFromFile('requestSubmissionsForm').setWidth(400).setHeight(100);
    ui.showModalDialog(form, 'Request Submissions');
  }
}

function processFormData(formData) {
  try {
    if (!formData.month || !formData.year) {
      throw new Error('Month and year not provided');
    }

    Logger.log(formData.month);
    Logger.log(formData.year);
    requestSubmissionsModule(formData.month, formData.year);
    return true;
  } catch (error) {
    console.error('Error processing form data', error);
    return false;
  }
}

// Request Submissions: sends emails, sets up the new mailing, and reminder trigger
function requestSubmissionsModule(month, year) {
  if (!month || !year) [month, year] = getPreviousMonthYear();

  // Update form titles with the current reporting month
  updateFormTitlesWithCurrentReportingMonth(month, year);
  Logger.log('Form titles updated with the current reporting month.');

  const registrySheet = getRegistrySheet(); // Open the "Registry" sheet
  Logger.log('Opened "Registry" sheet from "Ambassador Registry" spreadsheet.');

  // Fetch data from Registry sheet (Emails and Status)
  const registryEmailColIndex = getRequiredColumnIndexByName(registrySheet, AMBASSADOR_EMAIL_COLUMN);
  const registryAmbassadorStatus = getRequiredColumnIndexByName(registrySheet, AMBASSADOR_STATUS_COLUMN);
  const registryAmbassadorDiscordHandle = getRequiredColumnIndexByName(registrySheet, AMBASSADOR_DISCORD_HANDLE_COLUMN);
  const registryData = registrySheet
    .getRange(2, 1, registrySheet.getLastRow() - 1, registrySheet.getLastColumn())
    .getValues(); // Fetch Emails, Discord Handles, and Status
  Logger.log(`Fetched data from "Registry" sheet: ${JSON.stringify(registryData)}`);

  // Filter out ambassadors with 'Expelled' in their status
  const eligibleEmails = registryData
    .filter((row) => isActiveAmbassador(row, registryEmailColIndex - 1, registryAmbassadorStatus - 1)) // Exclude expelled ambassadors
    .map((row) => [normalizeEmail(row[registryEmailColIndex - 1]), row[registryAmbassadorDiscordHandle - 1]]); // Extract only emails
  Logger.log(`Eligible ambassadors emails: ${JSON.stringify(eligibleEmails)}`);

  // Calculate the exact deadline date based on submission window
  const submissionWindowStart = new Date();
  const submissionDeadline = new Date(submissionWindowStart.getTime() + minutesToMilliseconds(SUBMISSION_WINDOW_MINUTES));
  const submissionDeadlineDate = Utilities.formatDate(submissionDeadline, 'UTC', 'MMMM dd, yyyy HH:mm:ss') + ' UTC';

  eligibleEmails.forEach((row) => {
    const email = row[0]; //
    const discordHandle = row[1]; // Get Discord Handle from Registry

    // Validating email
    if (!isValidEmail(email)) {
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
    sendEmailNotification(email, '☑️Request for Submission', message);
  });

  // Save the submission window start time in Los Angeles time zone format
  setSubmissionWindowStart(submissionWindowStart);
  // Log the request in the "Request Log" sheet
  try {
    logRequest('Submission', month, year, submissionWindowStart, submissionDeadline);
  } catch (error) {
    Logger.log(`Error logging request: ${error.message}`);
  }

  // Set a trigger to check for non-respondents and send reminders
  setupSubmissionReminderTrigger(submissionWindowStart);

  Logger.log('Request Submission completed successfully!');
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
  const submissionWindowEnd = new Date(submissionWindowStart.getTime() + minutesToMilliseconds(SUBMISSION_WINDOW_MINUTES));

  // Open Registry and Form Responses sheets
  const registrySheet = getRegistrySheet();
  const formResponseSheet = getSubmissionFormResponseSheet();

  Logger.log('Sheets successfully fetched.');

  // Get column indices for required headers
  const registryEmailColIndex = getRequiredColumnIndexByName(registrySheet, AMBASSADOR_EMAIL_COLUMN);
  const registryAmbassadorStatusColIndex = getRequiredColumnIndexByName(registrySheet, AMBASSADOR_STATUS_COLUMN);
  const responseEmailColIndex = getRequiredColumnIndexByName(formResponseSheet, SUBM_FORM_USER_PROVIDED_EMAIL_COLUMN);
  const responseTimestampColIndex = getRequiredColumnIndexByName(formResponseSheet, GOOGLE_FORM_TIMESTAMP_COLUMN);

  // Fetch registry data and filter eligible emails
  const registryData = registrySheet
    .getRange(2, 1, registrySheet.getLastRow() - 1, registrySheet.getLastColumn())
    .getValues();
  const eligibleEmails = registryData
    .filter((row) => isActiveAmbassador(row, registryEmailColIndex - 1, registryAmbassadorStatusColIndex - 1)) // Case-insensitive check
    .map((row) => normalizeEmail(row[registryEmailColIndex - 1]));

  Logger.log(`Eligible emails: ${eligibleEmails}`);

  // Fetch form responses
  const responseData = formResponseSheet
    .getRange(2, 1, formResponseSheet.getLastRow() - 1, formResponseSheet.getLastColumn())
    .getValues();

  // Filter valid responses within submission window
  const validResponses = responseData.filter((row) => {
    const timestamp = new Date(row[responseTimestampColIndex - 1]);
    return timestamp >= submissionWindowStart && timestamp <= submissionWindowEnd;
  });

  const respondedEmails = validResponses.map((row) => normalizeEmail(row[responseEmailColIndex - 1]));
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
  const registrySheet = getRegistrySheet(); // Open the "Registry" sheet
  Logger.log('Opened "Registry" sheet.');

  if (!nonRespondents || nonRespondents.length === 0) {
    Logger.log('No non-respondents found.');
    return; // Exit if there are no non-respondents
  }

  // Dynamically fetch column indices
  const registryEmailColIndex = getRequiredColumnIndexByName(registrySheet, AMBASSADOR_EMAIL_COLUMN); // Email column index
  const registryDiscordHandleColIndex = getRequiredColumnIndexByName(registrySheet, AMBASSADOR_DISCORD_HANDLE_COLUMN); // Discord Handle column index

  nonRespondents.forEach((email) => {
    // Validate email format
    if (!isValidEmail(email)) {
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
      sendEmailNotification(email, '🕚 Reminder to Submit', message);
    } else {
      Logger.log(`Error: Could not find the ambassador with email ${email}`);
    }
  });
}

function getPreviousMonthYear() {
  // Get deliverable date of the reporting month at first time (previous month date)
  const deliverableDate = getPreviousMonthDate();
  Logger.log(`Deliverable date: ${deliverableDate}`);

  // Get the universal spreadsheet time zone
  const spreadsheetTimeZone = getProjectTimeZone();
  Logger.log(`Spreadsheet time zone: ${spreadsheetTimeZone}`);

  const month = Utilities.formatDate(deliverableDate, spreadsheetTimeZone, 'MMMM'); // Format the deliverable date to get the month name
  const year = Utilities.formatDate(deliverableDate, spreadsheetTimeZone, 'yyyy'); // Format the deliverable date to get the year
  Logger.log(`Formatted month and year: ${month} ${year}`);
  return [month, year];
}
