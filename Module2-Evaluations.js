// MODULE 2

// Basic function for Request Evaluations menu item processing
function requestEvaluationsModule() {
  const evaluationWindowStart = new Date(); // Capture start time for the evaluation window

  // Step 1: Create a month sheet and column in the Overall score
  createMonthSheetAndOverallColumn();

  // Step 2: Generating the review matrix (submitters and evaluators)
  generateReviewMatrix();

  // Step 3: Sending evaluation requests
  sendEvaluationRequests();

  // Set Evaluation Window Start Time
  setEvaluationWindowStart(evaluationWindowStart); // Save the evaluation window start time
  Logger.log(`Evaluation window start time set to: ${evaluationWindowStart}`);

  // Step 4: Filling out the Discord handle evaluators in the month sheet
  populateMonthSheetWithEvaluators();

  // Step 5: Deleting existing triggers before setting new ones
  deleteExistingTriggers(); // Delete all triggers before adding new ones

  // Step 6: Setting triggers
  setupEvaluationResponseTrigger(); // Setting the onFormSubmit trigger to process evaluation responses
  setupEvaluationTriggers(evaluationWindowStart); // Setting triggers for reminders and closures
}

/**
 * Creates a month sheet and corresponding column in the 'Overall score' sheet.
 */
function createMonthSheetAndOverallColumn() {
  try {
    Logger.log('Execution started');

    // Opening "Ambassadors' Scores" spreadsheet and "Overall score" sheet
    const scoresSpreadsheet = SpreadsheetApp.openById(AMBASSADORS_SCORES_SPREADSHEET_ID);
    const overallScoreSheet = scoresSpreadsheet.getSheetByName(OVERALL_SCORE_SHEET_NAME);

    if (!overallScoreSheet) {
      Logger.log(`Sheet "${OVERALL_SCORE_SHEET_NAME}" isn't found in "Ambassadors' Scores" sheet.`);
      return;
    }

    // Get project time zone
    const spreadsheetTimeZone = getProjectTimeZone();
    Logger.log(`Time zone of the table: ${spreadsheetTimeZone}`);

    // Get first day of previous month
    const deliverableMonthDate = getPreviousMonthDate(spreadsheetTimeZone);
    Logger.log(
      `Previous month date: ${Utilities.formatDate(deliverableMonthDate, spreadsheetTimeZone, 'yyyy-MM-dd HH:mm:ss z')}`
    );

    // Form month name, e.g., 'September 2024'
    const deliverableMonthName = Utilities.formatDate(deliverableMonthDate, spreadsheetTimeZone, 'MMMM yyyy');
    Logger.log(`Month name: "${deliverableMonthName}"`);

    // Create or clear existing sheet if there is one
    let monthSheet = scoresSpreadsheet.getSheetByName(deliverableMonthName);
    if (monthSheet) {
      monthSheet.clear(); // Clear existing sheet if there is
      Logger.log(`Cleared existing sheet: "${deliverableMonthName}".`);
    } else {
      // Finding index for inserting new month sheet before existing month sheets
      const sheetIndex = findInsertIndexForMonthSheet(scoresSpreadsheet);
      monthSheet = scoresSpreadsheet.insertSheet(deliverableMonthName, sheetIndex);
      Logger.log(`New sheet created: "${deliverableMonthName}".`);
    }

    // Adding headers
    monthSheet.getRange(1, 1).setValue('Submitter');
    monthSheet.getRange(1, 2).setValue('Score-1');
    monthSheet.getRange(1, 3).setValue("Evaluator's Discord-1");
    monthSheet.getRange(1, 4).setValue('Remarks-1');
    monthSheet.getRange(1, 5).setValue('Score-2');
    monthSheet.getRange(1, 6).setValue("Evaluator's Discord-2");
    monthSheet.getRange(1, 7).setValue('Remarks-2');
    monthSheet.getRange(1, 8).setValue('Score-3');
    monthSheet.getRange(1, 9).setValue("Evaluator's Discord-3");
    monthSheet.getRange(1, 10).setValue('Remarks-3');
    monthSheet.getRange(1, 11).setValue('Final Score');
    Logger.log(`Headers added to sheet: "${deliverableMonthName}".`);

    // Apply background colors to the specified columns
    monthSheet.getRange(2, 2, monthSheet.getMaxRows() - 1, 3).setBackground('#ebeee3');
    monthSheet.getRange(2, 5, monthSheet.getMaxRows() - 1, 3).setBackground('#ffffff');
    monthSheet.getRange(2, 8, monthSheet.getMaxRows() - 1, 3).setBackground('#ebeee3');

    // Sort submitter column alphabetically
    const lastRow = monthSheet.getLastRow();
    if (lastRow > 1) {
      // Ensure there are rows for sorting
      monthSheet.getRange(2, 1, lastRow - 1, 1).sort({ column: 1, ascending: true });
      Logger.log('Sorted the Submitter column alphabetically.');
    }

    // Getting existing columns in "Overall score" sheet
    const lastColumn = overallScoreSheet.getLastColumn();
    const existingColumns = overallScoreSheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    Logger.log(`Existing columns in "Overall score": ${existingColumns.join(', ')}`);

    // Check if month column already exists in Overall Score sheet
    const columnExists = doesColumnExist(existingColumns, deliverableMonthDate, spreadsheetTimeZone);
    if (columnExists) {
      Logger.log(`Column for "${deliverableMonthName}" already exists in "Overall score". Skipping creation.`);
      return;
    }

    // Finding index for inserting the new column after the last existing column
    const insertIndex = findInsertIndexForMonth(existingColumns);
    Logger.log(`New column insertion index: ${insertIndex}`);

    // Insert new column after last index
    overallScoreSheet.insertColumnAfter(insertIndex);
    const newHeaderCell = overallScoreSheet.getRange(1, insertIndex + 1);
    Logger.log(`Insert date in cell: Column ${insertIndex + 1}, Row 1`);

    // Set type of header as Date object (with same time as other columns are)
    const safeDate = new Date(deliverableMonthDate.getTime());
    safeDate.setUTCHours(7, 0, 0, 0); // Set time on 7:00 UTC, to match other columns
    newHeaderCell.setValue(safeDate);

    // Set cells format as 'MMMM yyyy', to display only month and year
    newHeaderCell.setNumberFormat('MMMM yyyy');

    // Clear any unintended background in the new month column
    const columnRange = overallScoreSheet.getRange(2, insertIndex + 1, overallScoreSheet.getLastRow() - 1);
    columnRange.setBackground(null); // Resetting background

    Logger.log(`Column for "${deliverableMonthName}" successfully added to "Overall score".`);
  } catch (error) {
    Logger.log(`Error in createMonthSheetAndOverallColumn: ${error}`);
  }
}

/**
 * Generates the review matrix by assigning evaluators to submitters multiple times
 * and chooses the attempt with the fewest "Has No Evaluator" results.
 * If an attempt finds a perfect solution (0 "Has No Evaluator"), it stops early and uses that result.
 */
function generateReviewMatrix() {
  try {
    Logger.log('Starting generateReviewMatrix with multiple attempts.');

    const MAX_ATTEMPTS = 5;
    let bestAssignments = null;
    let bestScore = Infinity; // The lower the score, the better (score = count of "Has No Evaluator")

    // Try multiple times and choose the best result
    for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
      Logger.log(`Attempt ${attempt} of ${MAX_ATTEMPTS}`);
      const { assignments, countHasNoEvaluator } = attemptSingleAssignment();

      if (!assignments || assignments.length === 0) {
        // If no assignments were made, consider this attempt as a failure
        continue;
      }

      Logger.log(`HasNoEvaluator count: ${countHasNoEvaluator}`);

      // If we found an attempt with zero "Has No Evaluator", use it immediately
      if (countHasNoEvaluator === 0) {
        bestAssignments = assignments;
        bestScore = 0;
        Logger.log(`Found a perfect assignment with no "Has No Evaluator" on attempt ${attempt}.`);
        break;
      }

      // Otherwise, keep track if this is the best so far
      if (countHasNoEvaluator < bestScore) {
        bestAssignments = assignments;
        bestScore = countHasNoEvaluator;
      }
    }

    if (bestAssignments) {
      // Write the best assignments to the Review Log
      writeAssignmentsToReviewLog(bestAssignments);
      Logger.log(`Final assignments chosen with ${bestScore} "Has No Evaluator".`);
    } else {
      Logger.log('No assignments could be made in the given attempts.');
    }
  } catch (error) {
    Logger.log(`Error in generateReviewMatrix: ${error}`);
  }
}

// ensure the 'Email Address' column name in the form is used or replaced with the actual column name from your environment or sheet.
function attemptSingleAssignment() {
  // Logging start of the function
  Logger.log('Starting single attempt of generateReviewMatrix.');

  // Define and compile a regex for validating email format
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

  // Access the Registry Spreadsheet and Registry Sheet
  Logger.log('Accessing registry spreadsheet and sheet.');
  const registrySpreadsheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID);
  const registrySheet = registrySpreadsheet.getSheetByName(REGISTRY_SHEET_NAME);
  if (!registrySheet) {
    Logger.log('Error: Registry sheet not found.');
    return { assignments: [], countHasNoEvaluator: Infinity };
  }

  // Get the Submission Form Responses sheet
  Logger.log('Accessing submission form responses sheet.');
  const formResponseSheet = getSubmissionFormResponseSheet(); // Must be defined in SharedUtilities
  if (!formResponseSheet) {
    Logger.log('Error: Submission form responses sheet not found.');
    return { assignments: [], countHasNoEvaluator: Infinity };
  }

  // Get Review Log Sheet
  Logger.log('Accessing Review Log sheet.');
  const reviewLogSheet = registrySpreadsheet.getSheetByName(REVIEW_LOG_SHEET_NAME);
  if (!reviewLogSheet) {
    Logger.log('Error: Review Log sheet not found.');
    return { assignments: [], countHasNoEvaluator: Infinity };
  }

  // Clear and initialize Review Log for this attempt
  Logger.log('Clearing and initializing Review Log sheet.');
  reviewLogSheet.clearContents();
  reviewLogSheet.getRange(1, 1).setValue('Submitter');
  reviewLogSheet.getRange(1, 2).setValue('Reviewer 1');
  reviewLogSheet.getRange(1, 3).setValue('Reviewer 2');
  reviewLogSheet.getRange(1, 4).setValue('Reviewer 3');

  // Get Submission Window start and end times
  Logger.log('Retrieving submission window start time.');
  const submissionWindowStart = getSubmissionWindowStart(); // must be defined in SharedUtilities
  if (!submissionWindowStart) {
    Logger.log('Submission window start time not found. Exiting this attempt.');
    return { assignments: [], countHasNoEvaluator: Infinity };
  }
  const submissionWindowEnd = new Date(submissionWindowStart);
  submissionWindowEnd.setMinutes(submissionWindowStart.getMinutes() + SUBMISSION_WINDOW_MINUTES);

  Logger.log(`Submission window is from ${submissionWindowStart} to ${submissionWindowEnd}`);

  // Get responses from Submission form
  // TODO Suggestion: pick a convention - -1 when assigning, or when using in getRange?
  const lastRow = formResponseSheet.getLastRow() - 1; // excluding headers
  if (lastRow < 1) {
    Logger.log('No submissions found in Form Responses sheet.');
    return { assignments: [], countHasNoEvaluator: Infinity };
  }

  Logger.log(`Retrieving submission form responses (without header). Rows of data: ${lastRow}`);
  const responseData = formResponseSheet.getRange(2, 1, lastRow, formResponseSheet.getLastColumn()).getValues();

  // Dynamically get column indices for timestamp and email columns in the form responses sheet
  Logger.log(`Looking for column: ${GOOGLE_FORM_TIMESTAMP_COLUMN}`);
  const timestampColumnIndex = getColumnIndexByName(formResponseSheet, GOOGLE_FORM_TIMESTAMP_COLUMN);

  // Replace 'Email Address' below with the exact column name you identified from your logs or environment
  // TODO Confirm: make sure this is the intended email constant
  Logger.log('Looking for column: Email Address');
  const emailColumnIndex = getColumnIndexByName(formResponseSheet, SUBM_FORM_USER_PROVIDED_EMAIL_COLUMN);

  if (timestampColumnIndex === -1 || emailColumnIndex === -1) {
    Logger.log(
      `Error: Required columns not found. Timestamp index: ${timestampColumnIndex}, Email index: ${emailColumnIndex}`
    );
    return { assignments: [], countHasNoEvaluator: Infinity };
  }

  // Filter responses by Submission window
  Logger.log('Filtering responses by submission window.');
  const validResponses = responseData.filter((row) => {
    const timestamp = new Date(row[timestampColumnIndex - 1]);
    return timestamp >= submissionWindowStart && timestamp <= submissionWindowEnd;
  });

  Logger.log(`Found ${validResponses.length} valid submissions within the submission window.`);

  if (validResponses.length === 0) {
    Logger.log('No valid submissions found within the submission window.');
    return { assignments: [], countHasNoEvaluator: Infinity };
  }

  // Extract submitter emails
  Logger.log('Extracting submitter emails from valid responses.');
  const submittersEmails = validResponses.map((row) =>
    row[emailColumnIndex - 1]
      ? String(row[emailColumnIndex - 1])
          .trim()
          .toLowerCase()
      : ''
  );
  Logger.log(`Submitters Emails: ${JSON.stringify(submittersEmails)}`);

  // Retrieve all ambassadors from registry excluding those with 'Expelled' status
  Logger.log('Retrieving ambassadors data from registry.');
  const registryAmbassadorEmailColumnIndex = getColumnIndexByName(registrySheet, AMBASSADOR_EMAIL_COLUMN);
  const registryAmbassadorStatusColumnIndex = getColumnIndexByName(registrySheet, AMBASSADOR_STATUS_COLUMN);

  if (registryAmbassadorEmailColumnIndex === -1 || registryAmbassadorStatusColumnIndex === -1) {
    Logger.log('Error: Required columns (Email or Status) not found in Registry sheet.');
    return { assignments: [], countHasNoEvaluator: Infinity };
  }

  const ambassadorData = registrySheet
    .getRange(2, 1, registrySheet.getLastRow() - 1, registrySheet.getLastColumn())
    .getValues();

  Logger.log('Filtering ambassadors by status and valid email format.');
  const allAmbassadorsEmails = ambassadorData
    .filter((row) => !row[registryAmbassadorStatusColumnIndex - 1].includes('Expelled')) // exclude expelled
    .filter((row) => {
      const email = row[registryAmbassadorEmailColumnIndex - 1];
      if (!emailRegex.test(email)) {
        Logger.log(`Invalid email format encountered: "${email}". Excluding.`);
        return false;
      }
      return true;
    })
    .map((row) =>
      String(row[registryAmbassadorEmailColumnIndex - 1])
        .trim()
        .toLowerCase()
    );

  Logger.log(`Eligible Ambassadors Emails: ${JSON.stringify(allAmbassadorsEmails)}`);

  // Create pool of potential evaluators (each evaluator can appear up to 3 times)
  Logger.log('Creating pool of potential evaluators.');
  const potentialEvaluators = [...allAmbassadorsEmails, ...allAmbassadorsEmails, ...allAmbassadorsEmails];

  // Dictionary to track how many times an evaluator is assigned
  const ambassadorCount = {};

  // Initialize assignments
  const assignments = [];
  let countHasNoEvaluator = 0;

  // Assign evaluators to each submitter
  Logger.log('Assigning evaluators to submitters.');
  submittersEmails.forEach((submitter) => {
    const reviewers = [];
    for (let i = 0; i < 3; i++) {
      const availableEvaluators = potentialEvaluators.filter(
        (evaluator) =>
          evaluator !== submitter && // exclude submitter evaluating themselves
          (ambassadorCount[evaluator] || 0) < 3 && // no more than 3 assignments per evaluator
          !reviewers.includes(evaluator) // no duplicate evaluators for same submitter
      );

      if (availableEvaluators.length === 0) {
        reviewers.push('Has No Evaluator');
        Logger.log(`No available evaluators for ${submitter} in round ${i + 1}`);
        countHasNoEvaluator++;
        continue;
      }

      const randomIndex = Math.floor(Math.random() * availableEvaluators.length);
      const selectedEvaluator = availableEvaluators[randomIndex];

      reviewers.push(selectedEvaluator);
      ambassadorCount[selectedEvaluator] = (ambassadorCount[selectedEvaluator] || 0) + 1;
      Logger.log(`Assigned ${selectedEvaluator} to ${submitter} in round ${i + 1}`);
    }
    assignments.push({ submitter, reviewers });
  });

  // Determine which evaluators were never used
  Logger.log('Determining unassigned evaluators.');
  const assignedEvaluators = Object.keys(ambassadorCount);
  const unassignedEvaluators = allAmbassadorsEmails.filter((email) => !assignedEvaluators.includes(email));

  // Send exemption emails to unassigned evaluators
  Logger.log('Sending exemption emails to unassigned evaluators.');
  sendExemptionEmails(allAmbassadorsEmails, unassignedEvaluators);

  Logger.log(`Attempt result: ${assignments.length} assignments, countHasNoEvaluator = ${countHasNoEvaluator}`);
  return { assignments, countHasNoEvaluator };
}

/**
 * Writes the given assignments to the Review Log sheet.
 * Ensures the correct headers and assignments are written dynamically.
 */
function writeAssignmentsToReviewLog(assignments) {
  try {
    const registrySpreadsheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID);
    const reviewLogSheet = registrySpreadsheet.getSheetByName(REVIEW_LOG_SHEET_NAME);

    if (!reviewLogSheet) {
      Logger.log(`Error: Review Log sheet not found.`);
      return;
    }

    // Clear the sheet and write headers
    Logger.log(`Clearing contents of the Review Log sheet.`);
    reviewLogSheet.clearContents();
    const headers = ['Submitter', 'Reviewer 1', 'Reviewer 2', 'Reviewer 3'];
    reviewLogSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Write assignments dynamically
    Logger.log(`Writing assignments to the Review Log sheet.`);
    const dataToWrite = assignments.map((assignment) => {
      const row = [assignment.submitter, ...assignment.reviewers];
      return row;
    });

    if (dataToWrite.length === 0) {
      Logger.log(`No assignments to write to Review Log.`);
      return;
    }

    reviewLogSheet.getRange(2, 1, dataToWrite.length, headers.length).setValues(dataToWrite);

    Logger.log(`Assignments successfully written to Review Log sheet.`);
  } catch (error) {
    Logger.log(`Error in writeAssignmentsToReviewLog: ${error.message}`);
  }
}

/**
 * Sends exemption emails to evaluators who were not assigned any submitters.
 * @param {Array} allEvaluators - List of all possible evaluators.
 * @param {Array} assignedEvaluators - List of evaluators who have been assigned submitters.
 */
function sendExemptionEmails(allEvaluators, unassignedEvaluators) {
  Logger.log('Starting sendExemptionEmails.');

  Logger.log(`Unassigned evaluators: ${JSON.stringify(unassignedEvaluators)}`);

  unassignedEvaluators.forEach((evaluator) => {
    try {
      const subject = 'Exemption from Evaluation';
      const body = EXEMPTION_FROM_EVALUATION_TEMPLATE;

      if (SEND_EMAIL) {
        MailApp.sendEmail({
          to: evaluator,
          subject: subject,
          htmlBody: body,
        });
        Logger.log(`Exemption email sent to: ${evaluator}`);
      } else {
        Logger.log(`Warning! Sending email disabled: Exemption email must be sent to ${evaluator}`);
      }
    } catch (error) {
      Logger.log(`Failed to send email to: ${evaluator}. Error: ${error}`);
    }
  });

  Logger.log('sendExemptionEmails completed.');
}

/**
 * Sends evaluation requests based on the generated review matrix.
 */
function sendEvaluationRequests() {
  try {
    // Opening Review Log sheet
    const reviewLogSheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(
      REVIEW_LOG_SHEET_NAME
    ); // Correct ID
    Logger.log(`Opened sheet: ${REVIEW_LOG_SHEET_NAME}`);

    // Get project time zone
    const spreadsheetTimeZone = getProjectTimeZone();
    Logger.log(`Spreadsheet TimeZone: ${spreadsheetTimeZone}`);

    const lastRow = reviewLogSheet.getLastRow();
    const lastColumn = reviewLogSheet.getLastColumn();
    Logger.log(`Review Log Sheet - Last Row: ${lastRow}, Last Column: ${lastColumn}`);

    // Checking if there is data in Review Log for processing
    if (lastRow < 2) {
      Logger.log('No data in Review Log sheet. Exiting sendEvaluationRequests.');
      return;
    }

    // Get data from the sheet (starting from the second row, first through fourth columns)
    const reviewData = reviewLogSheet.getRange(2, 1, lastRow - 1, 4).getValues(); // Evaluations matrix
    Logger.log(`Retrieved ${reviewData.length} rows of data for the review.`);

    // Get the name of the previous month for sending requests
    const deliverableMonthDate = getPreviousMonthDate(spreadsheetTimeZone); // Use shared utility
    const deliverableMonthName = Utilities.formatDate(deliverableMonthDate, spreadsheetTimeZone, 'MMMM yyyy');
    Logger.log(`Name of previous month: ${deliverableMonthName}`);

    // Calculate evaluation window deadline date
    const evaluationWindowStart = new Date();
    const evaluationDeadline = new Date(evaluationWindowStart.getTime() + EVALUATION_WINDOW_MINUTES * 60 * 1000); // Adjust to milliseconds
    const evaluationDeadlineDate = Utilities.formatDate(evaluationDeadline, spreadsheetTimeZone, 'MMMM dd, yyyy');

    reviewData.forEach((row, rowIndex) => {
      const submitterEmail = row[0]; // submitter's email
      const reviewersEmails = [row[1], row[2], row[3]].filter((email) => email); // Evaluators' Emails

      Logger.log(
        `String processing ${
          rowIndex + 2
        }: Email Submitter: ${submitterEmail}, Email Evaluators: ${reviewersEmails.join(', ')}`
      );

      // Getting submitter's Discord handle
      const submitterDiscordHandle = getDiscordHandleFromEmail(submitterEmail); // Call from SharedUtilities
      Logger.log(`Discord Submitter: ${submitterDiscordHandle}`);

      // Getting the details of the contribution
      const contributionDetails = getContributionDetailsByEmail(submitterEmail, spreadsheetTimeZone); // Call from SharedUtilities
      Logger.log(`Contribution details: ${contributionDetails}`);

      reviewersEmails.forEach((reviewerEmail) => {
        try {
          const evaluatorDiscordHandle = getDiscordHandleFromEmail(reviewerEmail); // Call from SharedUtilities
          Logger.log(`Discord Evaluator: ${evaluatorDiscordHandle}`);

          // Forming a message for evaluation
          const message = REQUEST_EVALUATION_EMAIL_TEMPLATE.replace('{AmbassadorDiscordHandle}', evaluatorDiscordHandle)
            .replace('{Month}', deliverableMonthName) // Use string name of the month
            .replace('{AmbassadorSubmitter}', submitterDiscordHandle)
            .replace('{SubmissionsList}', contributionDetails)
            .replace('{EvaluationFormURL}', EVALUATION_FORM_URL)
            .replace('{EVALUATION_DEADLINE_DATE}', evaluationDeadlineDate);

          if (SEND_EMAIL) {
            MailApp.sendEmail({
              to: reviewerEmail,
              subject: '⚖️Request for Evaluation',
              htmlBody: message, // Use htmlBody to ensure clickable link
            });
            Logger.log(
              `Evaluation request sent to ${reviewerEmail} (Discord: ${evaluatorDiscordHandle}) for submitter: ${submitterDiscordHandle}`
            );
          } else {
            if (!testing) {
              Logger.log(
                `WARNING: Production mode with email disabled. Evaluation request email logged but NOT SENT for ${reviewerEmail}`
              );
            } else {
              Logger.log(`Test mode: The evaluation request must be sent to ${reviewerEmail}`);
            }
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
  try {
    Logger.log(`Fetching contribution details for email: ${email}`);

    // Use unified project time zone
    const projectTimeZone = getProjectTimeZone();

    const formResponseSheet = SpreadsheetApp.openById(AMBASSADORS_SUBMISSIONS_SPREADSHEET_ID).getSheetByName(
      FORM_RESPONSES_SHEET_NAME
    );

    if (!formResponseSheet) {
      Logger.log(`Error: Sheet "${FORM_RESPONSES_SHEET_NAME}" not found.`);
      return 'No contribution details found for this submitter.';
    }

    // Retrieve submission window start and calculate end times
    const submissionWindowStart = getSubmissionWindowStart(); // Retrieve start time from SharedUtilities
    if (!submissionWindowStart) {
      Logger.log('Error: Submission window start time not found.');
      return 'No contribution details found for this submitter.';
    }
    const submissionWindowEnd = new Date(submissionWindowStart.getTime() + SUBMISSION_WINDOW_MINUTES * 60 * 1000);
    Logger.log(`Submission window: ${submissionWindowStart} to ${submissionWindowEnd}`);

    // Get form responses
    const formResponseTimestampColumnIndex = getColumnIndexByName(formResponseSheet, GOOGLE_FORM_TIMESTAMP_COLUMN);
    const contributionDetailsColumnIndex = getColumnIndexByName(
      formResponseSheet,
      GOOGLE_FORM_CONTRIBUTION_DETAILS_COLUMN
    );
    const contributionLinksColumnIndex = getColumnIndexByName(formResponseSheet, GOOGLE_FORM_CONTRIBUTION_LINKS_COLUMN);

    const formResponseEmailColumnIndex = getColumnIndexByName(formResponseSheet, SUBM_FORM_USER_PROVIDED_EMAIL_COLUMN);
    const formData = formResponseSheet
      .getRange(2, 1, formResponseSheet.getLastRow() - 1, formResponseSheet.getLastColumn())
      .getValues();

    // Find the corresponding response within the submission window
    for (let row of formData) {
      const timestamp = new Date(row[formResponseTimestampColumnIndex - 1]); // Assuming Timestamp is in the 1st column
      const respondentEmail = row[formResponseEmailColumnIndex - 1]?.trim().toLowerCase(); // Assuming Email is in the 2nd column

      if (timestamp >= submissionWindowStart && timestamp <= submissionWindowEnd && respondentEmail === email) {
        const contributionText = row[contributionDetailsColumnIndex - 1]; // Contribution details in the 4th column
        const contributionLinks = row[contributionLinksColumnIndex - 1]; // Links in the 5th column
        Logger.log(`Contribution found for email: ${email}`);
        return `Contribution Details: ${contributionText}\nLinks: ${contributionLinks}`;
      }
    }

    Logger.log(`No contribution details found for email: ${email}`);
    return 'No contribution details found for this submitter.';
  } catch (error) {
    Logger.log(`Error in getContributionDetailsByEmail: ${error.message}`);
    return 'An error occurred while fetching contribution details.';
  }
}

/**
 * Populates the month sheet with Discord handles of evaluators.
 */
function populateMonthSheetWithEvaluators() {
  try {
    Logger.log('Populating month sheet with evaluators.');

    // Use unified project time zone
    const projectTimeZone = getProjectTimeZone();

    // Open the Ambassadors' Scores spreadsheet and get the month sheet
    const scoresSheet = SpreadsheetApp.openById(AMBASSADORS_SCORES_SPREADSHEET_ID);
    const monthSheetName = Utilities.formatDate(getPreviousMonthDate(projectTimeZone), projectTimeZone, 'MMMM yyyy');
    const monthSheet = scoresSheet.getSheetByName(monthSheetName);

    if (!monthSheet) {
      Logger.log(`Month sheet ${monthSheetName} not found.`);
      return;
    }

    // Get submitter-evaluator assignments from the shared function
    const assignments = getReviewLogAssignments();

    Object.keys(assignments).forEach((submitterEmail, index) => {
      const submitterDiscordHandle = getDiscordHandleFromEmail(submitterEmail);

      if (!submitterDiscordHandle) {
        Logger.log(`Discord handle not found for submitter email: ${submitterEmail}`);
        return;
      }

      Logger.log(`Row ${index + 2}: Submitter Discord Handle: ${submitterDiscordHandle}`);

      // Fill submitter handle in the Month Sheet
      monthSheet.getRange(index + 2, 1).setValue(submitterDiscordHandle);

      // Get evaluators' Discord handles and fill them in the Month Sheet
      const evaluatorsEmails = assignments[submitterEmail];
      const evaluatorsDiscordHandles = evaluatorsEmails.map((email) => {
        const handle = getDiscordHandleFromEmail(email);
        return handle || 'Unknown Evaluator';
      });

      // Populate evaluator Discord handles in columns 3, 6, 9
      evaluatorsDiscordHandles.forEach((handle, idx) => {
        const discordColumnIndex = 3 * idx + 3;
        monthSheet.getRange(index + 2, discordColumnIndex).setValue(handle);
      });
    });

    Logger.log(`Evaluators populated in month sheet ${monthSheetName}.`);
  } catch (error) {
    Logger.log(`Error in populateMonthSheetWithEvaluators: ${error}`);
  }
}

/**
 * Function to reprocess all evaluation forms within the evaluation window
 */
function batchProcessEvaluationResponses() {
  try {
    Logger.log('Starting batch processing of evaluation responses.');

    const form = FormApp.openById(EVALUATION_FORM_ID);
    if (!form) {
      Logger.log('Error: Form not found with the given ID.');
      return;
    }

    const { evaluationWindowStart, evaluationWindowEnd } = getEvaluationWindowTimes();
    Logger.log(`Evaluation window: ${evaluationWindowStart} to ${evaluationWindowEnd}`);

    const formResponses = form.getResponses();
    const filteredResponses = formResponses.filter((response) => {
      const timestamp = new Date(response.getTimestamp());
      return timestamp >= evaluationWindowStart && timestamp <= evaluationWindowEnd;
    });

    Logger.log(`Total form responses to process: ${filteredResponses.length}`);

    filteredResponses.forEach((formResponse) => {
      const event = { response: formResponse };
      processEvaluationResponse(event);
    });

    Logger.log('Batch processing of evaluation responses completed.');
  } catch (error) {
    Logger.log(`Error in batchProcessEvaluationResponses: ${error}`);
  }
}

/**
 * Function to process evaluation form responses from Google Forms.
 * It extracts the evaluator's email, the Discord handle of the submitter, and the grade,
 * and then updates the respective columns in the month sheet.
 */
// Function to process evaluation responses and update the month sheet
function processEvaluationResponse(e) {
  try {
    Logger.log('processEvaluationResponse triggered.');

    const spreadsheetTimeZone = getProjectTimeZone(); // Get project time zone
    const scoresSpreadsheet = SpreadsheetApp.openById(AMBASSADORS_SCORES_SPREADSHEET_ID);

    if (!e || !e.response) {
      Logger.log('Error: Event parameter is missing or does not have a response.');
      return;
    }

    const formResponse = e.response;
    Logger.log('Form response received.');

    const formSubmitterEmail = formResponse.getRespondentEmail();
    if (!formSubmitterEmail) {
      Logger.log('Error: Respondent email is missing. Ensure that email collection is enabled in the form.');
      return;
    }
    Logger.log(`Form Submitter's Email from google form: ${formSubmitterEmail}`);
    const responseTime = formResponse.getTimestamp();
    Logger.log(`Evaluation response received at: ${responseTime}`);

    const itemResponses = formResponse.getItemResponses();
    let evaluatorEmail = '';
    let submitterDiscordHandle = '';
    let grade = NaN;
    let remarks = '';
    evaluatorEmail = formSubmitterEmail;

    itemResponses.forEach((itemResponse) => {
      const question = itemResponse.getItem().getTitle();
      const answer = itemResponse.getResponse();
      Logger.log(`Question: ${question}, Answer: ${answer}, Type of answer: ${typeof answer}`);
      // TODO Suggestion: change to use constants, changed this because forms has different value than what is hard coded
      if (question === 'Discord handle of the ambassador you are evaluating? (Not your own D-Handle)') {
        submitterDiscordHandle = String(answer).trim();
      } else if (question === 'Please assign a grade on a scale of 0 to 5 ') {
        const gradeMatch = String(answer).match(/\d+/);
        if (gradeMatch) grade = parseFloat(gradeMatch[0]);
      } else if (question === 'Remarks (optional)') {
        remarks = answer;
      }
    });

    Logger.log(`Evaluator Email Provided: ${evaluatorEmail}`);
    Logger.log(`Submitter Discord Handle (provided): ${submitterDiscordHandle}`);
    Logger.log(`Grade: ${grade}`);
    Logger.log(`Remarks: ${remarks}`);

    if (!evaluatorEmail || !submitterDiscordHandle || isNaN(grade)) {
      Logger.log('Missing required data. Exiting processEvaluationResponse.');
      return;
    }

    // TODO Discuss: why is this filter commented out?
    // confirmed that we are processing late evaluations; putting this back in.
    const { evaluationWindowStart, evaluationWindowEnd } = getEvaluationWindowTimes();
    if (responseTime < evaluationWindowStart || responseTime > evaluationWindowEnd) {
      Logger.log(
        `Evaluation received at ${responseTime} outside the window from ${evaluationWindowStart} to ${evaluationWindowEnd}. Response will be ignored.`
      );
      return;
    }

    // Retrieve assignments from Review Log and find expected submitters
    const assignments = getReviewLogAssignments();
    const expectedSubmitters = [];

    for (const [submitterEmail, evaluators] of Object.entries(assignments)) {
      if (evaluators.includes(evaluatorEmail)) {
        const submitterDiscord = getDiscordHandleFromEmail(submitterEmail);
        if (submitterDiscord) expectedSubmitters.push(submitterDiscord.trim());
      }
    }

    if (expectedSubmitters.length === 0) {
      Logger.log(`No expected submitters found for evaluator: ${evaluatorEmail}`);
      return;
    }
    Logger.log(`Expected submitters for evaluator ${evaluatorEmail}: ${expectedSubmitters.join(', ')}`);

    const correctedDiscordHandle = bruteforceDiscordHandle(submitterDiscordHandle, expectedSubmitters);
    if (!correctedDiscordHandle) {
      Logger.log(`Could not match Discord handle: ${submitterDiscordHandle} for evaluator: ${evaluatorEmail}`);
      return;
    }
    Logger.log(`Corrected Discord Handle: ${correctedDiscordHandle}`);
    submitterDiscordHandle = correctedDiscordHandle;

    const evaluatorDiscordHandle = getDiscordHandleFromEmail(evaluatorEmail);
    if (!evaluatorDiscordHandle) {
      Logger.log(`Discord handle not found for evaluator email: ${evaluatorEmail}`);
      return;
    }
    Logger.log(`Evaluator Discord Handle: ${evaluatorDiscordHandle}`);

    // Get the month before the submissions were collected as teh deliverable Month
    const submissionStart = getSubmissionWindowTimes().submissionWindowStart;
    const deliverableMonthDate = getStartOfPriorMonth(spreadsheetTimeZone, submissionStart); // getting reporting month

    const monthSheetName = Utilities.formatDate(deliverableMonthDate, spreadsheetTimeZone, 'MMMM yyyy');
    const monthSheet = scoresSpreadsheet.getSheetByName(monthSheetName);

    if (!monthSheet) {
      Logger.log(`Month sheet ${monthSheetName} not found.`);
      return;
    }

    // Find row for submitter
    const submitterDiscordColumnIndex = getColumnIndexByName(monthSheet, 'Submitter');
    const submitterRows = monthSheet
      .getRange(2, submitterDiscordColumnIndex, monthSheet.getLastRow() - 1, 1)
      .getValues();
    let row = null;
    for (let i = 0; i < submitterRows.length; i++) {
      if (submitterRows[i][0] && submitterRows[i][0].toLowerCase() === submitterDiscordHandle.toLowerCase()) {
        row = i + 2; // Offset for header row
        break;
      }
    }

    if (!row) {
      Logger.log(`Submitter ${submitterDiscordHandle} not found in month sheet.`);
      return;
    }
    Logger.log(`Submitter ${submitterDiscordHandle} found at row ${row}`);

    // Update evaluator's grade and remarks in the correct column
    let gradeUpdated = false;
    for (let col = 2; col <= 8; col += 3) {
      const cellValue = monthSheet.getRange(row, col + 1).getValue();
      if (cellValue === evaluatorDiscordHandle) {
        monthSheet.getRange(row, col).setValue(grade);
        monthSheet.getRange(row, col + 2).setValue(remarks);
        Logger.log(
          `Updated grade and remarks for submitter ${submitterDiscordHandle} by evaluator ${evaluatorDiscordHandle}. Grade: ${grade}, Remarks: ${remarks}`
        );
        gradeUpdated = true;
        break;
      }
    }

    if (!gradeUpdated) {
      Logger.log(
        `Evaluator ${evaluatorDiscordHandle} not assigned to submitter ${submitterDiscordHandle} in month sheet.`
      );
    }

    // Retrieve grades from Score-1, Score-2, and Score-3 columns (ignoring Remarks columns)
    const gradesRange = [
      monthSheet.getRange(row, 2).getValue(), // Score-1
      monthSheet.getRange(row, 5).getValue(), // Score-2
      monthSheet.getRange(row, 8).getValue(), // Score-3
    ];
    // Counts grades from 0 to 5, excluding empty (NaN)
    const validGrades = gradesRange.filter((value) => typeof value === 'number' && !isNaN(value));

    if (validGrades.length > 0) {
      const finalScore = validGrades.reduce((sum, grade) => sum + grade, 0) / validGrades.length;
      monthSheet.getRange(row, 11).setValue(finalScore);
      Logger.log(`Final score updated for submitter ${submitterDiscordHandle}: ${finalScore}`);
    }
  } catch (error) {
    Logger.log(`Error in processEvaluationResponse: ${error}`);
  }
}

// This function attempts to find the closest match among expected Discord handles (in case of a typo)
/**
 * Attempts to find the closest matching Discord handle among expected handles.
 * First checks for an exact match, then allows a single-character difference if needed.
 * @param {string} providedHandle - The handle entered by the evaluator.
 * @param {Array<string>} expectedHandles - The list of possible handles for this evaluator.
 * @return {string|null} - The best-matching handle or null if no match is found.
 */
function bruteforceDiscordHandle(providedHandle, expectedHandles) {
  providedHandle = providedHandle.toLowerCase();

  // Step 1: Check for an exact match
  for (let handle of expectedHandles) {
    if (providedHandle === handle.toLowerCase()) {
      return handle;
    }
  }

  // Step 2: Check for a single-character difference
  let bestMatch = null;
  let foundSingleCharDifference = false;

  for (let handle of expectedHandles) {
    if (isSingleCharDifference(providedHandle, handle.toLowerCase())) {
      bestMatch = handle;
      foundSingleCharDifference = true;
      break;
    }
  }

  return foundSingleCharDifference ? bestMatch : null;
}

/**
 * Checks if two strings differ by only one character (insertion, deletion, or substitution).
 * @param {string} a - The first string.
 * @param {string} b - The second string.
 * @return {boolean} - True if the strings differ by exactly one character.
 */
function isSingleCharDifference(a, b) {
  if (Math.abs(a.length - b.length) > 1) return false;

  let differences = 0;
  let i = 0,
    j = 0;

  while (i < a.length && j < b.length) {
    if (a[i] !== b[j]) {
      differences++;
      if (differences > 1) return false;

      if (a.length > b.length) i++;
      else if (a.length < b.length) j++;
      else {
        i++;
        j++;
      }
    } else {
      i++;
      j++;
    }
  }

  return differences + (a.length - i) + (b.length - j) === 1;
}

/**
 * Check Review Log sheet and Form Responses sheet to identify evaluators who have not responded to all assigned evaluations.
 */
function sendEvaluationReminderEmails() {
  try {
    Logger.log('Starting to send evaluation reminder emails.');

    // Open the Review Log and Form Responses sheets
    const reviewLogSheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(
      REVIEW_LOG_SHEET_NAME
    );
    const formResponseSheet = SpreadsheetApp.openById(AMBASSADORS_SCORES_SPREADSHEET_ID).getSheetByName(
      EVAL_FORM_RESPONSES_SHEET_NAME
    );

    if (!reviewLogSheet || !formResponseSheet) {
      Logger.log('Review Log sheet or Form Responses sheet not found. Exiting sendEvaluationReminderEmails.');
      return;
    }

    // Step 1: Retrieve expected evaluations per evaluator from Review Log
    const reviewAssignments = getReviewLogAssignments(); // {submitterEmail: [evaluatorEmail1, evaluatorEmail2, ...]}

    // Step 2: Retrieve valid evaluator emails from Form Responses
    const validEvaluators = new Set(getValidEvaluationEmails(formResponseSheet)); // A Set of valid evaluator emails

    // Step 3: Track evaluators with incomplete evaluations for reminder emails
    const nonRespondents = new Set(); // Use Set to ensure unique entries

    // Iterate over each evaluator in the review assignments
    for (const [submitterEmail, evaluators] of Object.entries(reviewAssignments)) {
      evaluators.forEach((evaluatorEmail) => {
        // Count assigned evaluations per evaluator only once
        const assignedEvaluations = reviewAssignments[submitterEmail].includes(evaluatorEmail) ? 1 : 0;

        // Check if this specific evaluation has been completed
        const completedEvaluations = validEvaluators.has(evaluatorEmail) ? 1 : 0;

        // Add to non-respondents if completed evaluations are less than assigned
        if (completedEvaluations < assignedEvaluations) {
          nonRespondents.add(evaluatorEmail);
          Logger.log(`Evaluator ${evaluatorEmail} has not completed all evaluations.`);
        }
      });
    }

    // Step 4: Send a single reminder to each non-responding evaluator
    if (nonRespondents.size > 0) {
      Logger.log(`Non-respondents: ${[...nonRespondents].join(', ')}`);
      sendReminderEmailsToUniqueEvaluators([...nonRespondents]); // Pass unique evaluators list
    } else {
      Logger.log('All evaluations are completed within the time window. No reminders to send.');
    }
  } catch (error) {
    Logger.log(`Error in sendEvaluationReminderEmails: ${error}`);
  }
}

/**
 * Send a single reminder email to each unique evaluator with incomplete evaluations.
 */
function sendReminderEmailsToUniqueEvaluators(nonRespondents) {
  try {
    Logger.log('Sending reminder emails.');

    // Fetch eligible ambassador emails excluding those with "Expelled" status
    const eligibleEmails = getEligibleAmbassadorsEmails(); // Fetch eligible emails from SharedUtilities

    const registrySheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(
      REGISTRY_SHEET_NAME
    );
    if (!registrySheet) {
      Logger.log('Registry sheet not found.');
      return;
    }

    // Get column indices dynamically
    const registryEmailColIndex = getColumnIndexByName(registrySheet, AMBASSADOR_EMAIL_COLUMN);
    const registryDiscordColIndex = getColumnIndexByName(registrySheet, AMBASSADOR_DISCORD_HANDLE_COLUMN);

    if (registryEmailColIndex === -1 || registryDiscordColIndex === -1) {
      Logger.log('Error: Required columns (Email or Discord Handle) not found in Registry sheet.');
      return;
    }

    // Define regex for email validation
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

    nonRespondents.forEach((evaluatorEmail) => {
      // Skip ambassadors who are not eligible (marked as 'Expelled' or not found)
      if (!eligibleEmails.includes(evaluatorEmail)) {
        Logger.log(`Skipping evaluator ${evaluatorEmail} (marked as 'Expelled' or not eligible).`);
        return;
      }

      // Validate email format
      if (!emailRegex.test(evaluatorEmail)) {
        Logger.log(`Invalid email format for evaluator: "${evaluatorEmail}". Skipping.`);
        return;
      }

      const result = registrySheet.createTextFinder(evaluatorEmail).findNext(); // Find evaluator's row by email
      if (result) {
        const row = result.getRow();
        const discordHandle = registrySheet.getRange(row, registryDiscordColIndex).getValue();
        const email = registrySheet.getRange(row, registryEmailColIndex).getValue(); // Get email dynamically
        const message = REMINDER_EMAIL_TEMPLATE.replace('{AmbassadorDiscordHandle}', discordHandle);

        if (SEND_EMAIL) {
          MailApp.sendEmail(email, '🕚Reminder to Submit Evaluation', message);
          Logger.log(`Reminder email sent to: ${email} (Discord: ${discordHandle})`);
        } else {
          Logger.log(`Warning! Sending email disabled: Reminder email logged for ${email}`);
        }
      } else {
        Logger.log(`Error: Could not find the ambassador with email ${evaluatorEmail}`);
      }
    });
  } catch (error) {
    Logger.log(`Error in sendReminderEmailsToUniqueEvaluators: ${error}`);
  }
}

// Sets up the evaluation response trigger based on the form's submission
function setupEvaluationResponseTrigger() {
  try {
    Logger.log('Setting up evaluation response trigger.');

    const form = FormApp.openById(EVALUATION_FORM_ID);
    if (!form) {
      Logger.log('Error: Form not found with the given ID.');
      return;
    }

    // Delete any existing processEvaluationResponse triggers to avoid duplicates
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach((trigger) => {
      if (trigger.getHandlerFunction() === 'processEvaluationResponse') {
        ScriptApp.deleteTrigger(trigger);
        Logger.log(`Deleted existing trigger: ${trigger.getHandlerFunction()}`);
      }
    });

    // Set a new trigger for Evaluation form submissions
    ScriptApp.newTrigger('processEvaluationResponse').forForm(form).onFormSubmit().create();

    Logger.log('Evaluation response trigger set up successfully.');
  } catch (error) {
    Logger.log(`Error in setupEvaluationResponseTrigger: ${error}`);
  }
}

// Sets up all triggers needed for evaluation process and logs evaluation start time
function setupEvaluationTriggers(evaluationWindowStart) {
  try {
    const timeZone = getProjectTimeZone(); // Get project time zone

    // Save evaluation start time
    const evalStartTime = Utilities.formatDate(evaluationWindowStart, timeZone, 'yyyy-MM-dd HH:mm:ss z');
    PropertiesService.getScriptProperties().setProperty('evaluationWindowStart', evalStartTime);
    Logger.log(`Evaluation start time set to: ${evalStartTime}`);

    // Calculate evaluation end time
    const evaluationWindowEnd = new Date(evaluationWindowStart.getTime() + EVALUATION_WINDOW_MINUTES * 60 * 1000);
    Logger.log(`Evaluation window is from ${evalStartTime} to ${evaluationWindowEnd}`);

    // Set up evaluation reminder trigger
    setupEvaluationReminderTrigger(evaluationWindowStart);
  } catch (error) {
    Logger.log(`Error in setupEvaluationTriggers: ${error}`);
  }
}

// Sets up the evaluation reminder trigger and logs trigger times
function setupEvaluationReminderTrigger(evaluationWindowStart) {
  try {
    Logger.log('Setting up evaluation reminder trigger.');

    // Calculate reminder time by adding EVALUATION_WINDOW_REMINDER_MINUTES to evaluationWindowStart
    const reminderTime = new Date(evaluationWindowStart.getTime() + EVALUATION_WINDOW_REMINDER_MINUTES * 60 * 1000);

    // Create a trigger for sending evaluation reminder emails
    ScriptApp.newTrigger('sendEvaluationReminderEmails').timeBased().at(reminderTime).create();

    Logger.log(`Reminder trigger for evaluation set for: ${reminderTime}`);

    // Get the project timezone for consistent formatting
    const timeZone = getProjectTimeZone(); // Use shared utility

    // Format current time (trigger setup time) in project timezone
    const setupTime = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd HH:mm:ss z');
    Logger.log(`Trigger setup time: ${setupTime}`);

    // Format reminder trigger time in project timezone
    const formattedReminderTime = Utilities.formatDate(reminderTime, timeZone, 'yyyy-MM-dd HH:mm:ss z');
    Logger.log(`Trigger fire time: ${formattedReminderTime}`);
  } catch (error) {
    Logger.log(`Error in setupEvaluationReminderTrigger: ${error}`);
  }
}
