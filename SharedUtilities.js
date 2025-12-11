// ===== Time Conversion Utilities =====

/**
 * Converts minutes to milliseconds.
 * @param {number} minutes - The number of minutes to convert
 * @returns {number} The equivalent milliseconds
 */
function minutesToMilliseconds(minutes) {
  return minutes * 60 * 1000;
}

// ===== Sheet Access Utilities =====
// Centralized functions to access commonly used sheets across modules

/**
 * Gets the Registry sheet from the Ambassador Registry spreadsheet.
 * @returns {Sheet} The Registry sheet
 */
function getRegistrySheet() {
  return SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(REGISTRY_SHEET_NAME);
}

/**
 * Gets the Review Log sheet from the Ambassador Registry spreadsheet.
 * @returns {Sheet} The Review Log sheet
 */
function getReviewLogSheet() {
  return SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(REVIEW_LOG_SHEET_NAME);
}

/**
 * Gets the Conflict Resolution Team sheet from the Ambassador Registry spreadsheet.
 * @returns {Sheet} The CRT sheet
 */
function getCRTSheet() {
  return SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(
    CONFLICT_RESOLUTION_TEAM_SHEET_NAME
  );
}

/**
 * Gets the Ambassadors Scores spreadsheet.
 * @returns {Spreadsheet} The Ambassadors Scores spreadsheet
 */
function getScoresSpreadsheet() {
  return SpreadsheetApp.openById(AMBASSADORS_SCORES_SPREADSHEET_ID);
}

/**
 * Gets the Overall Score sheet from the Ambassadors Scores spreadsheet.
 * @returns {Sheet} The Overall Score sheet
 */
function getOverallScoreSheet() {
  return getScoresSpreadsheet().getSheetByName(OVERALL_SCORE_SHEET_NAME);
}

/**
 * Gets the Ambassador Registry spreadsheet.
 * @returns {Spreadsheet} The Ambassador Registry spreadsheet
 */
function getRegistrySpreadsheet() {
  return SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID);
}

/**
 * Gets the Submission Form Responses sheet.
 * @returns {Sheet} The Submission Form Responses sheet
 */
function getSubmissionResponsesSheet() {
  return SpreadsheetApp.openById(AMBASSADORS_SUBMISSIONS_SPREADSHEET_ID).getSheetByName(FORM_RESPONSES_SHEET_NAME);
}

/**
 * Gets the Evaluation Form Responses sheet.
 * @returns {Sheet} The Evaluation Form Responses sheet
 */
function getEvaluationResponsesSheet() {
  return SpreadsheetApp.openById(EVALUATION_RESPONSES_SPREADSHEET_ID).getSheetByName(EVAL_FORM_RESPONSES_SHEET_NAME);
}

// ===== Generic function to send email =====
/**
 * Sends an email notification. Supports TO, CC, and BCC recipients.
 * In testing mode with TESTER_EMAIL set, redirects all emails to tester with original recipient info.
 * @param {string} recipientEmail - The main recipient's email address (can be empty if only CC/BCC is used).
 * @param {string} subject - The subject of the email.
 * @param {string} body - The body of the email (plain text or HTML).
 * @param {string} [bcc] - Optional comma-separated list of BCC recipients.
 * @param {string} [cc] - Optional comma-separated list of CC recipients.
 */
function sendEmailNotification(recipientEmail, subject, body, bcc, cc) {
  try {
    if (!SEND_EMAIL) {
      logEmailNotSent(recipientEmail, subject, bcc, cc);
      return;
    }

    if (TESTING) {
      sendTestEmail(recipientEmail, subject, body, bcc, cc);
    } else {
      sendProductionEmail(recipientEmail, subject, body, bcc, cc);
    }
  } catch (error) {
    Logger.log(`Failed to send email to ${recipientEmail}, CC: ${cc}, BCC: ${bcc}: ${error.message}`);
  }
}

/**
 * Sends email in testing mode - redirected to tester with original recipient info.
 * @param {string} recipientEmail - Original recipient email
 * @param {string} subject - Email subject
 * @param {string} body - Email body
 * @param {string} bcc - BCC recipients
 * @param {string} cc - CC recipients
 */
function sendTestEmail(recipientEmail, subject, body, bcc, cc) {
  const originalRecipient = recipientEmail || '[none]';
  const originalBcc = bcc || '[none]';
  const originalCc = cc || '[none]';
  const testSubject = `[TEST] ${subject}`;
  const testBody = buildTestEmailBody(originalRecipient, originalBcc, originalCc, body);

  const mailOptions = {
    to: TESTER_EMAIL,
    subject: testSubject,
    htmlBody: testBody,
  };

  MailApp.sendEmail(mailOptions);
  Logger.log(
    `Test email redirected to tester: ${TESTER_EMAIL}, Original recipient: ${originalRecipient}, CC: ${originalCc}, BCC: ${originalBcc}, Subject: ${testSubject}`
  );
}

/**
 * Builds the test email body with original recipient information.
 * @param {string} originalRecipient - Original recipient email
 * @param {string} originalBcc - Original BCC recipients
 * @param {string} originalCc - Original CC recipients
 * @param {string} body - Original email body
 * @returns {string} - Formatted test email body
 */
function buildTestEmailBody(originalRecipient, originalBcc, originalCc, body) {
  let testBody = `<p><strong>Testing email would have gone to: ${originalRecipient}</strong></p>
`;

  if (originalCc !== '[none]') {
    testBody += `<p><strong>CC would have gone to: ${originalCc}</strong></p>
`;
  }

  if (originalBcc !== '[none]') {
    testBody += `<p><strong>BCC would have gone to: ${originalBcc}</strong></p>
`;
  }

  testBody += `<hr>
${body}`;
  return testBody;
}

/**
 * Sends email in production mode.
 * @param {string} recipientEmail - Recipient email
 * @param {string} subject - Email subject
 * @param {string} body - Email body
 * @param {string} bcc - BCC recipients
 * @param {string} cc - CC recipients
 */
function sendProductionEmail(recipientEmail, subject, body, bcc, cc) {
  const mailOptions = {
    to: recipientEmail,
    subject: subject,
    htmlBody: body,
  };

  if (cc) {
    mailOptions.cc = cc;
  }

  if (bcc) {
    mailOptions.bcc = bcc;
  }

  MailApp.sendEmail(mailOptions);
  Logger.log(
    `Email sent to: ${recipientEmail || '[none]'}, CC: ${cc || '[none]'}, BCC: ${bcc || '[none]'}, Subject: ${subject}`
  );
}

/**
 * Logs when email is not sent due to SEND_EMAIL being false.
 * @param {string} recipientEmail - Recipient email
 * @param {string} subject - Email subject
 * @param {string} bcc - BCC recipients
 * @param {string} cc - CC recipients
 */
function logEmailNotSent(recipientEmail, subject, bcc, cc) {
  if (!TESTING) {
    Logger.log(
      `WARNING: Production mode with email disabled. Email logged but NOT SENT to: ${recipientEmail}, CC: ${cc}, BCC: ${bcc}, Subject: ${subject}`
    );
  } else {
    Logger.log(`Test mode: Email to be sent to: ${recipientEmail}, CC: ${cc}, BCC: ${bcc}, Subject: ${subject}`);
  }
}

/**
 * Normalizes an email address to a consistent format.
 * - Converts to lowercase.
 * - For Gmail addresses, removes dots from the local part.
 * @param {string} email - The email address to normalize.
 * @returns {string|null} The normalized email address, or null if the input is invalid.
 */
function normalizeEmail(email) {
  if (!email || typeof email !== 'string') {
    return email;
  }

  let trimmedEmail = email.trim().toLowerCase();
  const atIndex = trimmedEmail.lastIndexOf('@');

  if (atIndex === -1) {
    return trimmedEmail; // Return as-is if not a valid email format
  }

  const localPart = trimmedEmail.substring(0, atIndex);
  const domain = trimmedEmail.substring(atIndex);

  if (domain === '@gmail.com') {
    return `${localPart.replace(/\./g, '')}${domain}`;
  }

  return trimmedEmail;
}

/**
 * Normalizes a Discord handle to a consistent format.
 * - Converts to lowercase.
 * - Trims whitespace.
 * @param {string} discordHandle - The Discord handle to normalize.
 * @returns {string} The normalized Discord handle, or empty string if input is invalid.
 */
function normalizeDiscordHandle(discordHandle) {
  if (!discordHandle || typeof discordHandle !== 'string') {
    return '';
  }
  return discordHandle.trim().toLowerCase();
}

//     Submission/Evaluation WINDOW TIME SET/GET

// Save the submission window start time (in PST)
function setSubmissionWindowStart(time) {
  const formattedTime = Utilities.formatDate(time, getProjectTimeZone(), "yyyy-MM-dd'T'HH:mm:ssXXX"); // Format with timezone offset
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
  const submissionWindowEnd = new Date(
    submissionWindowStart.getTime() + minutesToMilliseconds(EVALUATION_WINDOW_MINUTES)
  );
  return { submissionWindowStart, submissionWindowEnd };
}

// Save the evaluation window start time (in PST)
function setEvaluationWindowStart(time) {
  const formattedTime = Utilities.formatDate(time, getProjectTimeZone(), "yyyy-MM-dd'T'HH:mm:ssXXX"); // Format with timezone offset
  PropertiesService.getScriptProperties().setProperty('evaluationWindowStart', formattedTime); // Save the formatted time
  Logger.log(`Evaluation window start time saved: ${formattedTime}`);
}

function getEvaluationWindowTimes() {
  const evaluationWindowStartStr = PropertiesService.getScriptProperties().getProperty('evaluationWindowStart');
  if (!evaluationWindowStartStr) {
    throw new Error('Evaluation window start time not found.');
  }
  const evaluationWindowStart = new Date(evaluationWindowStartStr);
  const evaluationWindowEnd = new Date(
    evaluationWindowStart.getTime() + minutesToMilliseconds(EVALUATION_WINDOW_MINUTES)
  );
  return { evaluationWindowStart, evaluationWindowEnd };
}

/**
 * Gets all evaluation windows (original + supplemental).
 * Supplemental windows are identified by timestamp headers in the Review Log sheet.
 * @returns {Array<Object>} - Array of window objects with {start, end} dates
 */
function getAllEvaluationWindows() {
  const windows = [];

  Logger.log('=== Getting All Evaluation Windows ===');

  // Get the original evaluation window
  try {
    const { evaluationWindowStart, evaluationWindowEnd } = getEvaluationWindowTimes();
    windows.push({ start: evaluationWindowStart, end: evaluationWindowEnd });
    Logger.log(`Original evaluation window: ${evaluationWindowStart} to ${evaluationWindowEnd}`);
  } catch (error) {
    Logger.log(`Error getting original evaluation window: ${error}`);
  }

  // Get supplemental windows from Review Log column headers
  try {
    const reviewLogSheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(
      REVIEW_LOG_SHEET_NAME
    );

    if (!reviewLogSheet) {
      Logger.log('Review Log sheet not found. No supplemental windows to check.');
      return windows;
    }

    const lastColumn = reviewLogSheet.getLastColumn();
    const headerRow = reviewLogSheet.getRange(1, 1, 1, lastColumn).getValues()[0];

    // Track supplemental windows found (deduplicate by timestamp)
    const supplementalTimestamps = new Set();

    // Look for timestamp headers (ISO 8601 format) that indicate supplemental windows
    headerRow.forEach((header) => {
      if (!header) {
        return; // Skip empty headers
      }

      const headerStr = header.toString();

      if (isSupplementalColumnHeader(headerStr)) {
        try {
          const supplementalStart = new Date(headerStr);

          if (isNaN(supplementalStart.getTime())) {
            Logger.log(`ERROR: Failed to parse supplemental timestamp "${headerStr}"`);
            return;
          }

          const timestampKey = supplementalStart.getTime();

          if (supplementalTimestamps.has(timestampKey)) {
            return; // Duplicate, skip silently
          }

          supplementalTimestamps.add(timestampKey);

          const supplementalEnd = new Date(
            supplementalStart.getTime() + minutesToMilliseconds(EVALUATION_WINDOW_MINUTES)
          );
          windows.push({ start: supplementalStart, end: supplementalEnd });

          Logger.log(`Supplemental evaluation window: ${supplementalStart} to ${supplementalEnd}`);
        } catch (error) {
          Logger.log(`ERROR parsing supplemental timestamp "${headerStr}": ${error}`);
        }
      }
    });

    Logger.log(`Total supplemental windows found: ${supplementalTimestamps.size}`);
  } catch (error) {
    Logger.log(`Error reading supplemental windows from Review Log: ${error}`);
    Logger.log(`Stack trace: ${error.stack || 'Not available'}`);
  }

  Logger.log(`=== Total Evaluation Windows: ${windows.length} ===`);
  return windows;
}

/**
 * Extracts the list of valid submission emails from the request log sheet within the submission time window.
 * Filters out emails not present in Registry.
 * Uses dynamic column indexing for the submitter email and timestamp columns.
 * @param {Sheet} submissionSheet - The sheet containing submissions.
 * @returns {Array} - A list of unique valid submission emails within the submission time window.
 */
function getValidSubmissionEmails(submissionSheet, submissionWindowStart = null, submissionWindowEnd = null) {
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

  // Only get from script properties if dates not provided
  if (!submissionWindowStart || !submissionWindowEnd) {
    const windowTimes = getSubmissionWindowTimes();
    submissionWindowStart = windowTimes.submissionWindowStart;
    submissionWindowEnd = windowTimes.submissionWindowEnd;
  }
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
          `Row ${
            index + 2
          }: Submission at ${submissionTimestamp} outside time window ${submissionWindowStart} to ${submissionWindowEnd}.`
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

  // Get all evaluation time windows (original + supplemental)
  const allWindows = getAllEvaluationWindows();

  // Dynamically retrieve column indices using headers
  const evalEmailColumnIndex = getRequiredColumnIndexByName(
    evaluationResponsesSheet,
    EVAL_FORM_USER_PROVIDED_EMAIL_COLUMN
  );
  const evalTimestampColumnIndex = getRequiredColumnIndexByName(evaluationResponsesSheet, GOOGLE_FORM_TIMESTAMP_COLUMN);

  // Extract valid responses within any evaluation time window
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

      // Check if response falls within any evaluation window (original or supplemental)
      const isWithinAnyWindow = allWindows.some(
        (window) => responseTimestamp >= window.start && responseTimestamp <= window.end
      );

      if (!isWithinAnyWindow) {
        Logger.log(`Row ${index + 2}: Response outside all evaluation time windows.`);
        return false;
      }

      Logger.log(`Row ${index + 2}: Valid response found for email: ${evaluatorEmail}`);
      return true;
    })
    .map((row) => row[evalEmailColumnIndex - 1]?.trim().toLowerCase()); // Extract evaluator email

  Logger.log(`Valid evaluators (within time windows): ${validEvaluators.join(', ')}`);
  return validEvaluators;
}

/**
 * Fetches and returns the submitter-evaluator assignments from the Review Log.
 * Includes both original assignments (Reviewer 1-3) and supplemental assignments (ISO 8601 timestamp columns).
 * Dynamically determines column indices based on header names to avoid hardcoded indices.
 * @returns {Object} - A map of submitter emails to a list of evaluator emails.
 */
function getReviewLogAssignments() {
  Logger.log('Fetching all submitter-evaluator assignments from Review Log (including supplemental).');

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
  const submitterColIndex = getRequiredColumnIndexByName(reviewLogSheet, SUBMITTER_HANDLE_COLUMN_IN_REVIEW_LOG);

  // Get all evaluator columns (Reviewer 1, Reviewer 2, Reviewer 3, and any ISO 8601 timestamp-based supplemental columns)
  const evaluatorColIndices = [];
  for (let i = 0; i < headers.length; i++) {
    const header = headers[i];
    // Match "Reviewer N" or ISO 8601 timestamp format
    if (header && (header.toString().startsWith('Reviewer ') || isSupplementalColumnHeader(header))) {
      evaluatorColIndices.push(i + 1); // Convert to 1-indexed
    }
  }

  if (evaluatorColIndices.length === 0) {
    Logger.log('Error: No evaluator columns found in Review Log sheet.');
    return {};
  }

  // Fetch data for the entire sheet
  const reviewData = reviewLogSheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();

  // Structure the data as { submitter: [evaluators] }
  const assignments = {};
  reviewData.forEach((row) => {
    const submitterEmail = row[submitterColIndex - 1];
    const evaluators = evaluatorColIndices.map((colIndex) => row[colIndex - 1]).filter((email) => email); // Collect all evaluators' emails
    if (submitterEmail && evaluators.length > 0) {
      assignments[submitterEmail] = evaluators;
    }
  });

  return assignments;
}

/**
 * Checks if an evaluator is assigned as a supplemental evaluator for a specific submitter.
 * Returns the assignment information including the window start time parsed from the column header.
 * @param {string} submitterEmail - The submitter's email
 * @param {string} evaluatorEmail - The evaluator's email
 * @returns {Object|null} - Assignment object with windowStart, windowEnd, columnHeader, or null if not found
 */
function getSupplementalAssignmentForSubmitterAndEvaluator(submitterEmail, evaluatorEmail) {
  const reviewLogSheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(
    REVIEW_LOG_SHEET_NAME
  );

  if (!reviewLogSheet) {
    Logger.log('Error: Review Log sheet not found.');
    return null;
  }

  const lastRow = reviewLogSheet.getLastRow();
  const lastColumn = reviewLogSheet.getLastColumn();

  if (lastRow < 2) {
    Logger.log('No data found in Review Log sheet.');
    return null;
  }

  // Get headers
  const headers = reviewLogSheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const submitterColIndex = getRequiredColumnIndexByName(reviewLogSheet, SUBMITTER_HANDLE_COLUMN_IN_REVIEW_LOG);

  // Find supplemental columns (ISO 8601 timestamp headers)
  const supplementalColumns = [];
  for (let i = 0; i < headers.length; i++) {
    const header = headers[i];
    if (isSupplementalColumnHeader(header)) {
      supplementalColumns.push({
        colIndex: i + 1,
        timestampStr: header.toString(),
      });
    }
  }

  if (supplementalColumns.length === 0) {
    Logger.log('No supplemental columns found in Review Log.');
    return null;
  }

  // Find the submitter's row
  const reviewData = reviewLogSheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
  let submitterRow = null;

  for (let i = 0; i < reviewData.length; i++) {
    if (normalizeEmail(reviewData[i][submitterColIndex - 1]) === normalizeEmail(submitterEmail)) {
      submitterRow = i;
      break;
    }
  }

  if (submitterRow === null) {
    Logger.log(`Submitter ${submitterEmail} not found in Review Log.`);
    return null;
  }

  // Check if evaluator is in any supplemental column for this submitter
  const rowData = reviewData[submitterRow];

  for (const suppCol of supplementalColumns) {
    const evaluatorInCell = rowData[suppCol.colIndex - 1];
    if (evaluatorInCell && normalizeEmail(evaluatorInCell) === normalizeEmail(evaluatorEmail)) {
      // Found the evaluator in a supplemental column
      // Parse the timestamp from the header to get the window start time
      const windowStart = new Date(suppCol.timestampStr);
      const windowEnd = new Date(windowStart.getTime() + minutesToMilliseconds(EVALUATION_WINDOW_MINUTES));

      Logger.log(
        `Found supplemental assignment: ${evaluatorEmail} for ${submitterEmail} in column ${suppCol.timestampStr}`
      );
      Logger.log(`Supplemental window: ${windowStart} to ${windowEnd}`);

      return {
        windowStart: windowStart,
        windowEnd: windowEnd,
        columnHeader: suppCol.timestampStr,
      };
    }
  }

  Logger.log(`Evaluator ${evaluatorEmail} not found in supplemental columns for ${submitterEmail}.`);
  return null;
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
      .filter((row) =>
        isActiveAmbassador(row, registryAmbassadorEmailColumnIndex - 1, registryAmbassadorStatusColumnIndex - 1)
      ) // Exclude those marked as expelled
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
      `First day of the previous month (Local Time): ${Utilities.formatDate(
        previousMonth,
        timeZone,
        'yyyy-MM-dd HH:mm:ss z'
      )}`
    );

    return previousMonth;
  } catch (error) {
    Logger.log(`Error in getFirstDayOfReportingMonth: ${error.message}`);
    return null;
  }
}

//    FORMS' TITLES
//
// Main function to update the form titles based on the current reporting month
function updateFormTitlesWithCurrentReportingMonth(month, year) {
  // Retrieve the reporting month in "MMMM yyyy" format, e.g., "August 2024"
  const reportingMonth = `${month} ${year}`;

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
    alertAndLog(`Expected Column "${columnName}" not found in sheet "${sheet.getName()}".`);
    throw new Error('Required column not found');
  }
  return index;
}

/**
 * Finds the column index for a given column name in the header row of the sheet.
 * Handles both string and Date object comparisons.
 * String comparisons are case-insensitive and trim whitespace before comparison.
 * @param {Sheet} sheet - The sheet to search.
 * @param {string|Date} columnName - The name of the column to find (string) or Date object for month columns.
 * @returns {number} The column index (1-based), or -1 if not found.
 */
function getColumnIndexByName(sheet, columnName) {
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const index = header.findIndex((headerValue) => {
    // Handle Date object comparison for month columns
    if (columnName instanceof Date && headerValue instanceof Date) {
      // Compare dates by checking if they represent the same month/year
      return (
        columnName.getFullYear() === headerValue.getFullYear() &&
        columnName.getMonth() === headerValue.getMonth() &&
        columnName.getDate() === headerValue.getDate()
      );
    }

    // Handle string comparison for regular columns - case insensitive with trimming
    if (typeof columnName === 'string' && headerValue != null) {
      const headerStr = (headerValue?.trim?.() ?? headerValue.toString().trim()).toLowerCase();
      const searchStr = columnName.trim().toLowerCase();
      return headerStr === searchStr;
    }

    return false;
  });

  if (index == -1) {
    Logger.log(`Expected Column "${columnName}" not found in sheet "${sheet.getName()}" header row: "${header}".`);
    return -1;
  }
  return index + 1; // Convert to 1-based index
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

  const normalizedDiscordHandle = normalizeDiscordHandle(discordHandle);
  const rowIndex = handlesColumn.findIndex((handle) => normalizeDiscordHandle(handle) === normalizedDiscordHandle);

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

/**
 * Funciton to get primary team responsibilities to be used in the evaluation email
 */
function getPrimaryTeamResponsibilities(primaryTeam) {
  const teamKey = primaryTeam.toLowerCase();
  return PrimaryTeamResponsibilities[teamKey] || 'Responsibilities not found for the specified team.';
}

function logRequest(type, month, year, requestDateTime, windowEndDateTime) {
  const spreadsheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID);
  let requestLogSheet = spreadsheet.getSheetByName('Request Log');

  // Create the sheet if it doesn't exist
  if (!requestLogSheet) {
    requestLogSheet = spreadsheet.insertSheet('Request Log');
    requestLogSheet.appendRow(['Type', 'Month', 'Year', 'Request Date Time', 'Window End Date Time']);
  }

  // Append the new request details
  // Store dates in ISO 8601 format with explicit timezone offset
  requestLogSheet.appendRow([
    type,
    month,
    year,
    Utilities.formatDate(requestDateTime, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ssXXX"),
    Utilities.formatDate(windowEndDateTime, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ssXXX"),
  ]);

  Logger.log(`Logged ${type} request for ${month} ${year} in "Request Log" sheet.`);
}

/**
 * Gets a specific request from the Request Log by type, month, and year.
 * @param {string} type - Request type ('Submission' or 'Evaluation')
 * @param {string} month - Month name (e.g., 'October')
 * @param {string} year - Year as string (e.g., '2024')
 * @returns {Object|null} Request object with month, year, requestDateTime, windowEndDateTime or null if not found
 */
function getRequestByTypeMonthAndYear(type, month, year) {
  try {
    const spreadsheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID);
    const requestLogSheet = spreadsheet.getSheetByName('Request Log');
    if (!requestLogSheet) return null;

    const lastRow = requestLogSheet.getLastRow();
    if (lastRow < 2) return null;

    const typeColIndex = getRequiredColumnIndexByName(requestLogSheet, REQUEST_LOG_REQUEST_TYPE_COLUMN);
    const monthColIndex = getRequiredColumnIndexByName(requestLogSheet, REQUEST_LOG_MONTH_COLUMN);
    const yearColIndex = getRequiredColumnIndexByName(requestLogSheet, REQUEST_LOG_YEAR_COLUMN);
    const startTimeColIndex = getRequiredColumnIndexByName(requestLogSheet, REQUEST_LOG_START_TIME_COLUMN);
    const endTimeColIndex = getRequiredColumnIndexByName(requestLogSheet, REQUEST_LOG_END_TIME_COLUMN);

    const data = requestLogSheet.getRange(2, 1, lastRow - 1, requestLogSheet.getLastColumn()).getValues();

    for (const row of data) {
      const rowType = row[typeColIndex - 1];
      const rowMonth = row[monthColIndex - 1];
      const rowYear = row[yearColIndex - 1].toString();

      if (rowType === type && rowMonth === month && rowYear === year) {
        const matchedRequest = {
          month: rowMonth,
          year: rowYear,
          requestDateTime: new Date(row[startTimeColIndex - 1]),
          windowEndDateTime: new Date(row[endTimeColIndex - 1]),
        };
        Logger.log(
          `Found ${type} request for ${month} ${year}: window ${matchedRequest.requestDateTime} to ${matchedRequest.windowEndDateTime}`
        );
        return matchedRequest;
      }
    }

    Logger.log(`No ${type} request found for ${month} ${year} in Request Log.`);
    return null;
  } catch (error) {
    Logger.log(`Error in getRequestByTypeMonthAndYear: ${error.message}`);
    return null;
  }
}

/**
 * Gets the most recent request of a given type (e.g., 'Submission', 'Evaluation') from the Request Log,
 * based on the latest Window End Date Time (i.e., the most recently completed request).
 * @param {string} type - The type of request to search for (e.g., 'Submission', 'Evaluation').
 * @returns {{month: string, year: string, requestDateTime: Date, windowEndDateTime: Date}|null} - The most recent request details or null if not found
 */
function getLatestRequestByType(type) {
  try {
    const spreadsheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID);
    const requestLogSheet = spreadsheet.getSheetByName('Request Log');
    if (!requestLogSheet) return null;

    const lastRow = requestLogSheet.getLastRow();
    if (lastRow < 2) return null;

    const typeColIndex = getRequiredColumnIndexByName(requestLogSheet, REQUEST_LOG_REQUEST_TYPE_COLUMN);
    const monthColIndex = getRequiredColumnIndexByName(requestLogSheet, REQUEST_LOG_MONTH_COLUMN);
    const yearColIndex = getRequiredColumnIndexByName(requestLogSheet, REQUEST_LOG_YEAR_COLUMN);
    const startTimeColIndex = getRequiredColumnIndexByName(requestLogSheet, REQUEST_LOG_START_TIME_COLUMN);
    const endTimeColIndex = getRequiredColumnIndexByName(requestLogSheet, REQUEST_LOG_END_TIME_COLUMN);

    const data = requestLogSheet.getRange(2, 1, lastRow - 1, requestLogSheet.getLastColumn()).getValues();

    let latestRequest = null;
    let latestEndTime = null;

    for (const row of data) {
      const rowType = row[typeColIndex - 1];
      if (rowType !== type) continue;

      const endTimeStr = row[endTimeColIndex - 1];
      const endTime = new Date(endTimeStr);
      if (!latestEndTime || endTime > latestEndTime) {
        latestEndTime = endTime;
        latestRequest = {
          month: row[monthColIndex - 1],
          year: row[yearColIndex - 1].toString(),
          requestDateTime: new Date(row[startTimeColIndex - 1]),
          windowEndDateTime: endTime,
        };
      }
    }

    if (latestRequest) {
      Logger.log(
        `Found latest ${type} request: ${latestRequest.month} ${latestRequest.year} (window ended on ${latestRequest.windowEndDateTime})`
      );
    } else {
      Logger.log(`No ${type} requests found in Request Log.`);
    }

    return latestRequest;
  } catch (error) {
    Logger.log(`Error in getLatestRequestByType: ${error.message}`);
    return null;
  }
}

/**
 * Gets the reporting month information from the latest request of a given type.
 * @param {string} type - The type of request to search for (e.g., 'Submission', 'Evaluation').
 * @returns {{month: string, year: string, monthName: string, firstDayDate: Date}|null} - The month and year to be evaluated
 */
function getReportingMonthFromRequestLog(type) {
  const latestRequest = getLatestRequestByType(type);

  if (!latestRequest) {
    return null;
  }

  // Create date object for the first day of the reporting month
  const monthNames = [
    'January',
    'February',
    'March',
    'April',
    'May',
    'June',
    'July',
    'August',
    'September',
    'October',
    'November',
    'December',
  ];
  const monthIndex = monthNames.indexOf(latestRequest.month);
  if (monthIndex === -1) {
    Logger.log(`Error: Invalid month name: ${latestRequest.month}`);
    return;
  }
  const deliverableMonthDate = new Date(parseInt(latestRequest.year), monthIndex, 1);
  Logger.log(
    `Reporting month date: ${Utilities.formatDate(deliverableMonthDate, getProjectTimeZone(), 'yyyy-MM-dd HH:mm:ss z')}`
  );

  return {
    month: latestRequest.month,
    year: latestRequest.year,
    monthName: `${latestRequest.month} ${latestRequest.year}`,
    firstDayDate: deliverableMonthDate,
  };
}

/**
 * Looks up both email and Discord handle for a given identifier (email or Discord handle).
 * @param {string} identifier - Email or Discord handle.
 * @returns {{email: string, discordHandle: string}|null}
 */
function lookupEmailAndDiscord(identifier) {
  const registrySheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(REGISTRY_SHEET_NAME);
  const emailCol = getRequiredColumnIndexByName(registrySheet, AMBASSADOR_EMAIL_COLUMN) - 1;
  const discordCol = getRequiredColumnIndexByName(registrySheet, AMBASSADOR_DISCORD_HANDLE_COLUMN) - 1;
  const registryData = registrySheet.getDataRange().getValues();
  const normalizedIdentifier = normalizeEmail(identifier) || normalizeDiscordHandle(identifier);
  for (const row of registryData) {
    const email = normalizeEmail(row[emailCol]);
    const discordHandle = normalizeDiscordHandle(row[discordCol]);
    if (email === normalizedIdentifier || discordHandle === normalizedIdentifier) {
      return { email, discordHandle };
    }
  }
  return null;
}

/**
 * Gets the current CRT members' emails and Discord handles from the last row of the CRT sheet.
 * @returns {Array<{email: string, discordHandle: string}>}
 */
function getCurrentCRTMemberEmails() {
  const crtSheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(
    CONFLICT_RESOLUTION_TEAM_SHEET_NAME
  );
  if (!crtSheet) return [];

  const crtData = crtSheet.getDataRange().getValues();
  if (crtData.length < 2) return [];

  // Get header row and find columns containing ambassador identifiers
  const headers = crtData[0];
  const identifierColumns = headers
    .map((header, index) => {
      const headerStr = (header || '').toString().toLowerCase();
      // Look for columns that likely contain ambassador identifiers
      return headerStr.includes('discord') || headerStr.includes('email') || headerStr.includes('ambassador')
        ? index
        : -1;
    })
    .filter((index) => index > 0); // Filter out -1 and first column (date)

  if (identifierColumns.length === 0) {
    Logger.log('No columns found containing ambassador identifiers');
    return [];
  }

  const lastRow = crtData[crtData.length - 1];
  const crtIdentifiers = identifierColumns
    .map((colIndex) => (lastRow[colIndex] || '').toString().trim().toLowerCase())
    .filter(Boolean);

  // For each, look up both email and discord handle
  return crtIdentifiers.map((id) => lookupEmailAndDiscord(id) || { email: id, discordHandle: id });
}

/**
 * Gets the current reporting month name from the request log.
 * @returns {string} - The month name (e.g., "January", "February")
 */
function getCurrentReportingMonthName() {
  try {
    const reportingMonth = getReportingMonthFromRequestLog('Submission');
    if (reportingMonth && reportingMonth.monthName) {
      return reportingMonth.monthName.toUpperCase();
    }

    // Fallback to previous month if no request log found
    const previousMonth = getPreviousMonthDate();
    const monthNames = [
      'JANUARY',
      'FEBRUARY',
      'MARCH',
      'APRIL',
      'MAY',
      'JUNE',
      'JULY',
      'AUGUST',
      'SEPTEMBER',
      'OCTOBER',
      'NOVEMBER',
      'DECEMBER',
    ];
    return monthNames[previousMonth.getMonth()];
  } catch (error) {
    Logger.log(`Error getting current reporting month name: ${error.message}`);
    return 'CURRENT MONTH';
  }
}

/**
 * Calculates a date that is the specified number of business days from today.
 * @param {number} businessDays - Number of business days to add
 * @returns {string} - Formatted date string (e.g., "January 15, 2024")
 */
function getBusinessDaysFromToday(businessDays) {
  try {
    let currentDate = new Date();
    let addedDays = 0;

    while (addedDays < businessDays) {
      currentDate.setDate(currentDate.getDate() + 1);
      // Check if it's a weekday (Monday = 1, Sunday = 0)
      const dayOfWeek = currentDate.getDay();
      if (dayOfWeek !== 0 && dayOfWeek !== 6) {
        // Not Saturday or Sunday
        addedDays++;
      }
    }

    return Utilities.formatDate(currentDate, getProjectTimeZone(), 'MMMM dd, yyyy');
  } catch (error) {
    Logger.log(`Error calculating business days: ${error.message}`);
    return 'within 3 business days';
  }
}

// ===== Predicate Functions for Complex Conditionals =====

/**
 * Checks if a value is a Date object.
 * @param {*} value - Value to check
 * @returns {boolean} True if value is a Date instance
 */
function isDate(value) {
  return value instanceof Date;
}

/**
 * Checks if an ambassador has already been expelled (status contains "Expelled").
 * @param {string} status - Ambassador status string
 * @returns {boolean} True if status indicates expulsion
 */
function isAlreadyExpelled(status) {
  return status && status.includes('Expelled');
}

/**
 * Checks if an ambassador registry row represents an active (non-expelled) ambassador.
 * @param {Array} row - Registry row data
 * @param {number} emailColumnIndex - Email column index (0-based)
 * @param {number} statusColumnIndex - Status column index (0-based)
 * @returns {boolean} True if ambassador is active and not expelled
 */
function isActiveAmbassador(row, emailColumnIndex, statusColumnIndex) {
  return row[emailColumnIndex]?.trim() && !row[statusColumnIndex]?.toLowerCase().includes('expelled');
}

/**
 * Checks if an email address is valid (not null/undefined/empty) and passes basic format validation.
 * @param {string} email - Email to validate
 * @returns {boolean} True if email is valid
 */
function isValidEmail(email) {
  if (!email || !email.trim()) {
    return false;
  }
  // Basic email format validation
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}
