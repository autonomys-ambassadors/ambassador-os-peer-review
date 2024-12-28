function validateEmailsInSubmissionForm(formResponseSheetId, registrySheetId) {
  Logger.log('Starting validation of emails in the Submission Form Responses sheet.');

  try {
    // Open the sheets
    const formResponseSheet = SpreadsheetApp.openById(AMBASSADORS_SUBMISSIONS_SPREADSHEET_ID).getSheetByName(
      FORM_RESPONSES_SHEET_NAME
    );
    const registrySheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(REGISTRY_SHEET_NAME);

    if (!formResponseSheet) {
      Logger.log('Error: Submission Form Responses sheet not found.');
      return false;
    }
    if (!registrySheet) {
      Logger.log('Error: Registry sheet not found.');
      return false;
    }

    // Get submission time window
    const { submissionWindowStart, submissionWindowEnd } = getSubmissionWindowTimes();

    // Get all emails from the Registry sheet
    const registryEmails = registrySheet
      .getRange(2, getColumnIndexByName(registrySheet, AMBASSADOR_EMAIL_COLUMN), registrySheet.getLastRow() - 1, 1)
      .getValues()
      .flat()
      .map((email) => email.trim().toLowerCase());
    Logger.log(`Loaded ${registryEmails.length} emails from the Registry.`);

    // Get headers and column indices dynamically
    const emailColumnIndex = getColumnIndexByName(formResponseSheet, 'Your Email Address');
    const timestampColumnIndex = getColumnIndexByName(formResponseSheet, 'Timestamp');

    if (emailColumnIndex === 0 || timestampColumnIndex === 0) {
      Logger.log('Error: Required columns not found in the Submission Form Responses sheet.');
      return false;
    }

    // Get all rows from the Submission Form Responses sheet
    const allRows = formResponseSheet
      .getRange(2, 1, formResponseSheet.getLastRow() - 1, formResponseSheet.getLastColumn())
      .getValues();
    Logger.log(`Loaded ${allRows.length} rows from the Submission Form Responses sheet.`);

    // Validate emails and log invalid rows
    let withinWindowCount = 0;
    let invalidEmailCount = 0;
    const invalidRows = [];

    allRows.forEach((row, index) => {
      const timestamp = new Date(row[timestampColumnIndex - 1]);
      const email = row[emailColumnIndex - 1]?.trim().toLowerCase();

      const isWithinWindow = timestamp >= submissionWindowStart && timestamp <= submissionWindowEnd;
      const isValidEmail = registryEmails.includes(email);

      if (isWithinWindow) {
        withinWindowCount++;
        if (!isValidEmail) {
          invalidEmailCount++;
          invalidRows.push({ row: index + 2, email, timestamp, fullRow: row });
          Logger.log(`Row ${index + 2}: Invalid email detected - ${email}`);
        }
      }
    });

    Logger.log(`${withinWindowCount} submissions were within the submission time window.`);
    Logger.log(`${invalidEmailCount} emails were invalid within the submission time window.`);

    if (invalidRows.length === 0) {
      Logger.log('All emails in the Submission Form Responses sheet are valid within the submission window.');
      return true;
    } else {
      Logger.log(
        `Found ${invalidRows.length} invalid emails in the Submission Form Responses sheet: ${JSON.stringify(
          invalidRows
        )}`
      );
      invalidRows.forEach((invalidRow) =>
        Logger.log(`Row ${invalidRow.row}: Full data - ${JSON.stringify(invalidRow.fullRow)}`)
      );
      return false;
    }
  } catch (error) {
    Logger.log(`Error in validateEmailsInSubmissionForm: ${error.message}`);
    return false;
  }
}