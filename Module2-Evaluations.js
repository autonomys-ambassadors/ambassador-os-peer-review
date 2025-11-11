// MODULE 2
// Basic function for Request Evaluations menu item processing
function requestEvaluationsModule() {
  // Step 0:determine which month will be evaluated
  const reportingMonth = getReportingMonthFromRequestLog('Submission');
  if (!reportingMonth) {
    alertAndLog('Error: Could not determine reporting month from Request Log.');
    return;
  }

  // Show user which month will be evaluated
  const confirmationMessage = `This will request evaluations for ${reportingMonth.monthName} based on the latest submission request. Continue?`;
  const userConfirmed = promptAndLog('Confirm Evaluation Month', confirmationMessage, ButtonSet.YES_NO);

  if (userConfirmed !== ButtonResponse.YES) {
    Logger.log('User cancelled evaluation request.');
    return;
  }

  const evaluationWindowStart = new Date(); // Capture start time for the evaluation window

  // Step 1: Create a month sheet and column in the Overall score
  createMonthSheetAndOverallColumn(reportingMonth);

  // Step 2: Generating the review matrix (submitters and evaluators)
  generateReviewMatrix();

  // Step 2.5: Update evaluation form questions
  updateEvaluationFormQuestions();

  // Step 3: Sending evaluation requests
  sendEvaluationRequests(reportingMonth);

  // Set Evaluation Window Start Time
  setEvaluationWindowStart(evaluationWindowStart); // Save the evaluation window start time
  Logger.log(`Evaluation window start time set to: ${evaluationWindowStart}`);

  // Step 4: Filling out the Discord handle evaluators in the month sheet
  populateMonthSheetWithEvaluators(reportingMonth);

  // Step 5: Deleting existing triggers before setting new ones
  deleteExistingTriggers(); // Delete all triggers before adding new ones

  // Step 6: Setting triggers
  setupEvaluationResponseTrigger(); // Setting the onFormSubmit trigger to process evaluation responses
  setupEvaluationTriggers(evaluationWindowStart); // Setting triggers for reminders and closures
}

/**
 * Creates a month sheet and corresponding column in the 'Overall score' sheet.
 * @param {Object} reportingMonth - The reporting month object containing month and year
 */
function createMonthSheetAndOverallColumn(reportingMonth) {
  try {
    Logger.log('Execution started');

    // Opening "Ambassadors' Scores" spreadsheet and "Overall score" sheet
    const scoresSpreadsheet = getScoresSpreadsheet();
    const overallScoreSheet = scoresSpreadsheet.getSheetByName(OVERALL_SCORE_SHEET_NAME);

    if (!overallScoreSheet) {
      Logger.log(`Sheet "${OVERALL_SCORE_SHEET_NAME}" isn't found in "Ambassadors' Scores" sheet.`);
      return;
    }

    Logger.log(`Reporting month date: ${reportingMonth.monthName}`);

    // Create or clear existing sheet if there is one
    let monthSheet = scoresSpreadsheet.getSheetByName(reportingMonth.monthName);
    if (monthSheet) {
      monthSheet.clear(); // Clear existing sheet if there is
      Logger.log(`Cleared existing sheet: "${reportingMonth.monthName}".`);
    } else {
      // Finding index for inserting new month sheet before existing month sheets
      const sheetIndex = findInsertIndexForMonthSheet(scoresSpreadsheet);
      monthSheet = scoresSpreadsheet.insertSheet(reportingMonth.monthName, sheetIndex);
      Logger.log(`New sheet created: "${reportingMonth.monthName}".`);
    }

    // Adding headers
    monthSheet.getRange(1, 1).setValue('Submitter');
    monthSheet.getRange(1, 2).setValue('Score-1');
    monthSheet.getRange(1, 3).setValue('Amb-1');
    monthSheet.getRange(1, 4).setValue('Remarks-1');
    monthSheet.getRange(1, 5).setValue('Score-2');
    monthSheet.getRange(1, 6).setValue('Amb-2');
    monthSheet.getRange(1, 7).setValue('Remarks-2');
    monthSheet.getRange(1, 8).setValue('Score-3');
    monthSheet.getRange(1, 9).setValue('Amb-3');
    monthSheet.getRange(1, 10).setValue('Remarks-3');
    monthSheet.getRange(1, 11).setValue('Final Score');
    Logger.log(`Headers added to sheet: "${reportingMonth.monthName}".`);

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
    const scoreColumn = getColumnIndexByName(overallScoreSheet, reportingMonth.firstDayDate);
    if (scoreColumn > 0) {
      Logger.log(`Column for "${reportingMonth.monthName}" already exists in "Overall score". Skipping creation.`);
      return;
    }

    // Finding index for inserting the new column after the last existing column
    const insertIndex = findInsertIndexForMonth(existingColumns);
    Logger.log(`New column insertion index: ${insertIndex}`);

    // Insert new column after last index
    overallScoreSheet.insertColumnAfter(insertIndex);
    const newHeaderCell = overallScoreSheet.getRange(1, insertIndex + 1);
    Logger.log(`Insert date in cell: Column ${insertIndex + 1}, Row 1`);

    newHeaderCell.setValue(reportingMonth.firstDayDate);

    // Set cells format as 'MMMM yyyy', to display only month and year
    newHeaderCell.setNumberFormat('MMMM yyyy');

    // Clear any unintended background in the new month column
    const columnRange = overallScoreSheet.getRange(2, insertIndex + 1, overallScoreSheet.getLastRow() - 1);
    columnRange.setBackground(null); // Resetting background

    Logger.log(`Column for "${reportingMonth.monthName}" successfully added to "Overall score".`);
  } catch (error) {
    Logger.log(`Error in createMonthSheetAndOverallColumn: ${error}`);
  }
}

function updateEvaluationFormQuestions(primaryTeam) {
  const form = FormApp.openById(EVALUATION_FORM_ID);
  const items = form.getItems();
  items.forEach((item) => {
    if (item.getTitle().includes('Please assign a grade')) {
      item.setHelpText(
        `Please consider the ambassador's contributions in relation to their primary team when making your assessment.`
      );
    }
  });
}

/**
 * Generates the review matrix by assigning evaluators to submitters multiple times
 * and chooses the attempt with the fewest "Has No Evaluator" results.
 * If an attempt finds a perfect solution (0 "Has No Evaluator"), it stops early and uses that result.
 * Filters out invalid submitters (e.g., those with emails not in Registry or empty rows).
 */
function generateReviewMatrix() {
  try {
    Logger.log('Starting generateReviewMatrix with multiple attempts.');

    // Access the Registry and Submission Form Sheets
    const registrySpreadsheet = getRegistrySpreadsheet();
    const registrySheet = registrySpreadsheet.getSheetByName(REGISTRY_SHEET_NAME);
    const formResponseSheet = getSubmissionFormResponseSheet();
    const reviewLogSheet = registrySpreadsheet.getSheetByName(REVIEW_LOG_SHEET_NAME);

    if (!registrySheet || !formResponseSheet || !reviewLogSheet) {
      Logger.log('Error: Required sheets not found.');
      return;
    }

    // Get valid submitters from the Submission Form sheet
    const validSubmitters = getValidSubmissionEmails(formResponseSheet);
    Logger.log(`Valid submitters: ${JSON.stringify(validSubmitters)}`);

    // Get eligible evaluators from the Registry
    const allEvaluators = getEligibleAmbassadorsEmails();
    Logger.log(`Eligible evaluators: ${JSON.stringify(allEvaluators)}`);

    const { assignments, countHasNoEvaluator } = attemptSingleAssignment(validSubmitters, allEvaluators);

    if (!assignments || assignments.length === 0) {
      alertAndLog(`Failed to generate assignments.`);
      throw new Error('Failed to generate assignements. Notify developers.');
    }

    Logger.log(`Assignment attempt resulted in ${countHasNoEvaluator} "Has No Evaluator".`);

    // Write the best assignments to the Review Log
    if (countHasNoEvaluator === 0) {
      writeAssignmentsToReviewLog(assignments);
      Logger.log(`Final assignments chosen with ${countHasNoEvaluator} "Has No Evaluator".`);
    } else {
      alertAndLog('Failed to generate valid assignments after one attempt.');
      throw new Error('Failed to generate valid assignments after one attempt. Notify Developers.');
    }
  } catch (error) {
    alertAndLog(`Error in generateReviewMatrix: ${error.message}`);
    throw new Error('Failed to generate review matrix. Notify developers.');
  }
}

/**
 * Attempts to assign evaluators to submitters in a single attempt.
 * Filters out invalid or empty data during the process.
 * @param {Array} validSubmitters - Array of valid submitters.
 * @param {Array} allEvaluators - Array of eligible evaluators.
 * @returns {Object} - Contains assignments and count of "Has No Evaluator".
 */
function attemptSingleAssignment(validSubmitters, allEvaluators) {
  Logger.log('Starting single attempt to assign evaluators.');

  const assignments = [];
  const evaluatorPool = [...allEvaluators, ...allEvaluators, ...allEvaluators]; // Triplicate evaluator pool for flexibility
  const evaluatorCounts = {}; // Track evaluator usage
  let countHasNoEvaluator = 0;

  validSubmitters.forEach((submitter) => {
    const reviewers = [];
    for (let i = 0; i < 3; i++) {
      const availableEvaluators = evaluatorPool.filter(
        (evaluator) =>
          evaluator !== submitter && // Exclude self-evaluation
          (evaluatorCounts[evaluator] || 0) < 3 && // Limit evaluators to 3 assignments
          !reviewers.includes(evaluator) // Avoid duplicate evaluators for the same submitter
      );

      if (availableEvaluators.length === 0) {
        reviewers.push('Has No Evaluator');
        countHasNoEvaluator++;
        alertAndLog(`No available evaluators for submitter ${submitter} in round ${i + 1}.`);
        throw new Error('No available evaluators found. Notify developers.');
      }

      const randomIndex = Math.floor(Math.random() * availableEvaluators.length);
      const selectedEvaluator = availableEvaluators[randomIndex];

      reviewers.push(selectedEvaluator);
      evaluatorCounts[selectedEvaluator] = (evaluatorCounts[selectedEvaluator] || 0) + 1;
      Logger.log(`Assigned evaluator ${selectedEvaluator} to submitter ${submitter} in round ${i + 1}.`);
    }
    assignments.push({ submitter, reviewers });
  });

  Logger.log(`Single attempt completed. "Has No Evaluator" count: ${countHasNoEvaluator}`);
  return { assignments, countHasNoEvaluator };
}

/**
 * Writes the given assignments to the Review Log sheet.
 * Ensures the correct headers and assignments are written dynamically.
 */
function writeAssignmentsToReviewLog(assignments) {
  try {
    const registrySpreadsheet = getRegistrySpreadsheet();
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

      sendEmailNotification(evaluator, subject, body);
    } catch (error) {
      Logger.log(`Failed to send email to: ${evaluator}. Error: ${error}`);
    }
  });

  Logger.log('sendExemptionEmails completed.');
}

/**
 * Sends evaluation requests based on the generated review matrix.
 * @param {Object} reportingMonth - The reporting month object containing month and year
 */
function sendEvaluationRequests(reportingMonth) {
  try {
    // Opening Review Log sheet
    const reviewLogSheet = getReviewLogSheet(); // Correct ID
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

    // Use the reporting month passed as parameter
    if (!reportingMonth) {
      Logger.log('Error: Could not determine reporting month from Request Log.');
      throw new Error('Could not determine reporting month from Request Log.');
    }
    Logger.log(`Reporting month: ${reportingMonth.monthName}`);

    // Calculate evaluation window deadline date
    const evaluationWindowStart = new Date();
    const evaluationDeadline = new Date(evaluationWindowStart.getTime() + minutesToMilliseconds(EVALUATION_WINDOW_MINUTES));
    const evaluationDeadlineDate = Utilities.formatDate(evaluationDeadline, 'UTC', 'MMMM dd, yyyy HH:mm:ss') + ' UTC';
    try {
      logRequest('Evaluation', reportingMonth.month, reportingMonth.year, evaluationWindowStart, evaluationDeadline);
    } catch (error) {
      Logger.log(error);
    }

    reviewData.forEach((row, rowIndex) => {
      const submitterEmail = row[0]; // submitter's email
      const reviewersEmails = [row[1], row[2], row[3]].filter((email) => email); // Evaluators' Emails

      Logger.log(
        `String processing ${
          rowIndex + 2
        }: Email Submitter: ${submitterEmail}, Email Evaluators: ${reviewersEmails.join(', ')}`
      );

      // Getting submitter's Discord handle
      const submitterDiscordHandle = lookupEmailAndDiscord(submitterEmail)?.discordHandle; // Call from SharedUtilities
      Logger.log(`Discord Submitter: ${submitterDiscordHandle}`);

      // Getting the details of the contribution
      const contributionDetails = getContributionDetailsByEmail(submitterEmail); // Call from SharedUtilities
      Logger.log(`Contribution details: ${contributionDetails}`);

      const primaryTeam = getAmbassadorPrimaryTeam(submitterEmail);

      reviewersEmails.forEach((reviewerEmail) => {
        try {
          const evaluatorDiscordHandle = lookupEmailAndDiscord(reviewerEmail)?.discordHandle; // Call from SharedUtilities
          Logger.log(`Discord Evaluator: ${evaluatorDiscordHandle}`);

          // Forming a message for evaluation
          const message = REQUEST_EVALUATION_EMAIL_TEMPLATE.replace('{AmbassadorDiscordHandle}', evaluatorDiscordHandle)
            .replace('{Month}', reportingMonth.monthName) // Use string name of the month
            .replace('{AmbassadorSubmitter}', submitterDiscordHandle)
            .replace('{SubmissionsList}', contributionDetails)
            .replace('{EvaluationFormURL}', EVALUATION_FORM_URL)
            .replace('{EVALUATION_DEADLINE_DATE}', evaluationDeadlineDate)
            .replace('{PrimaryTeam}', primaryTeam)
            .replace('{PrimaryTeamResponsibilities}', getPrimaryTeamResponsibilities(primaryTeam));

          sendEmailNotification(reviewerEmail, '⚖️Request for Evaluation', message);
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
function getContributionDetailsByEmail(email, submissionWindowStart = null, submissionWindowEnd = null) {
  try {
    Logger.log(`Fetching contribution details for email: ${email}`);

    // Use unified project time zone
    const projectTimeZone = getProjectTimeZone();

    const formResponseSheet = getSubmissionResponsesSheet();

    if (!formResponseSheet) {
      Logger.log(`Error: Sheet "${FORM_RESPONSES_SHEET_NAME}" not found.`);
      return 'No contribution details found for this submitter.';
    }

    // Retrieve submission window start and calculate end times
    // If not provided, get from current script properties (for current evaluation cycle)
    if (!submissionWindowStart) {
      submissionWindowStart = getSubmissionWindowStart(); // Retrieve start time from SharedUtilities
      if (!submissionWindowStart) {
        Logger.log('Error: Submission window start time not found.');
        return 'No contribution details found for this submitter.';
      }
    }
    if (!submissionWindowEnd) {
      submissionWindowEnd = new Date(submissionWindowStart.getTime() + minutesToMilliseconds(SUBMISSION_WINDOW_MINUTES));
    }
    Logger.log(`Submission window: ${submissionWindowStart} to ${submissionWindowEnd}`);

    // Get form responses
    const formResponseTimestampColumnIndex = getRequiredColumnIndexByName(
      formResponseSheet,
      GOOGLE_FORM_TIMESTAMP_COLUMN
    );
    const contributionDetailsColumnIndex = getRequiredColumnIndexByName(
      formResponseSheet,
      GOOGLE_FORM_CONTRIBUTION_DETAILS_COLUMN
    );
    const contributionLinksColumnIndex = getRequiredColumnIndexByName(
      formResponseSheet,
      GOOGLE_FORM_CONTRIBUTION_LINKS_COLUMN
    );

    const formResponseEmailColumnIndex = getRequiredColumnIndexByName(
      formResponseSheet,
      SUBM_FORM_USER_PROVIDED_EMAIL_COLUMN
    );
    const formData = formResponseSheet
      .getRange(2, 1, formResponseSheet.getLastRow() - 1, formResponseSheet.getLastColumn())
      .getValues();

    // Find the latest response within the submission window
    let latestSubmissionRow = null;
    let latestTimestamp = null;

    for (let row of formData) {
      const timestamp = new Date(row[formResponseTimestampColumnIndex - 1]); // Assuming Timestamp is in the 1st column
      const respondentEmail = normalizeEmail(row[formResponseEmailColumnIndex - 1]); // Assuming Email is in the 2nd column

      if (
        timestamp >= submissionWindowStart &&
        timestamp <= submissionWindowEnd &&
        respondentEmail === normalizeEmail(email)
      ) {
        if (!latestTimestamp || timestamp > latestTimestamp) {
          latestTimestamp = timestamp;
          latestSubmissionRow = row;
        }
      }
    }

    if (latestSubmissionRow) {
      const contributionText = latestSubmissionRow[contributionDetailsColumnIndex - 1]; // Contribution details in the 4th column
      const contributionLinks = latestSubmissionRow[contributionLinksColumnIndex - 1]; // Links in the 5th column
      Logger.log(`Latest contribution found for email: ${email}`);
      return `Contribution Details: ${contributionText}<br><br>Links: ${contributionLinks}`;
    }

    Logger.log(`No contribution details found for email: ${email}`);
    return 'No contribution details found for this submitter.';
  } catch (error) {
    Logger.log(`Error in getContributionDetailsByEmail: ${error.message}`);
    return 'An error occurred while fetching contribution details.';
  }
}

/**
 * Fetches the primary team of the ambassador from the registry.
 * @param {string} email - Ambassador's email
 * @returns {string} - Name of the primary team or empty string if not found
 */
function getAmbassadorPrimaryTeam(email) {
  try {
    Logger.log(`Looking for Primary Team for the email ${email}`);

    // Open the Registry spreadsheet
    const registrySpreadsheet = getRegistrySpreadsheet();
    const registrySheet = registrySpreadsheet.getSheetByName(REGISTRY_SHEET_NAME);

    if (!registrySheet) {
      Logger.log('Error: Registry spreadsheet not found');
      return '';
    }

    // Get all data from the Registry spreadsheet
    const registryData = registrySheet.getDataRange().getValues();
    const headerRow = registryData[0];

    // Find indices of the columns
    const emailColIndex = headerRow.indexOf(AMBASSADOR_EMAIL_COLUMN);
    const primaryTeamColIndex = headerRow.indexOf(AMBASSADOR_PRIMARY_TEAM_COLUMN);

    if (emailColIndex === -1 || primaryTeamColIndex === -1) {
      Logger.log('Error: Required columns not found in the Registry');
      return '';
    }

    // Search for the ambassador's email and return their primary team
    for (let i = 1; i < registryData.length; i++) {
      if (normalizeEmail(registryData[i][emailColIndex]) === normalizeEmail(email)) {
        const primaryTeam = registryData[i][primaryTeamColIndex] || '';
        Logger.log(`Primary team found for ${email}: ${primaryTeam}`);
        return primaryTeam;
      }
    }

    Logger.log(`No primary team found for the email: ${email}`);
    return '';
  } catch (error) {
    Logger.log(`Error while fetching primary team: ${error}`);
    return '';
  }
}

/**
 * Populates the month sheet with Discord handles of evaluators.
 * @param {Object} reportingMonth - The reporting month object containing month and year
 */
function populateMonthSheetWithEvaluators(reportingMonth) {
  try {
    Logger.log('Populating month sheet with evaluators.');

    // Use the reporting month passed as parameter
    if (!reportingMonth) {
      Logger.log('Error: Could not determine reporting month from Request Log.');
      return;
    }

    // Open the Ambassadors' Scores spreadsheet and get the month sheet
    const scoresSheet = getScoresSpreadsheet();
    const monthSheet = scoresSheet.getSheetByName(reportingMonth.monthName);

    if (!monthSheet) {
      Logger.log(`Month sheet ${reportingMonth.monthName} not found.`);
      return;
    }

    // Get submitter-evaluator assignments from the shared function
    const assignments = getReviewLogAssignments();

    Object.keys(assignments).forEach((submitterEmail, index) => {
      const submitterDiscordHandle = lookupEmailAndDiscord(submitterEmail)?.discordHandle;

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
        const handle = lookupEmailAndDiscord(email)?.discordHandle;
        return handle || 'Unknown Evaluator';
      });

      // Populate evaluator Discord handles in columns 3, 6, 9
      evaluatorsDiscordHandles.forEach((handle, idx) => {
        const discordColumnIndex = 3 * idx + 3;
        monthSheet.getRange(index + 2, discordColumnIndex).setValue(handle);
      });
    });

    Logger.log(`Evaluators populated in month sheet ${reportingMonth.monthName}.`);
  } catch (error) {
    Logger.log(`Error in populateMonthSheetWithEvaluators: ${error}`);
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
    const scoresSpreadsheet = getScoresSpreadsheet();

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
      if (question === EVAL_FORM_USER_PROVIDED_EMAIL_COLUMN) {
        evaluatorEmail = normalizeEmail(String(answer));
      } else if (question === GOOGLE_FORM_EVALUATION_HANDLE_COLUMN) {
        submitterDiscordHandle = String(answer).trim();
      } else if (question === GOOGLE_FORM_EVALUATION_GRADE_COLUMN) {
        const gradeMatch = String(answer).match(/\d+/);
        if (gradeMatch) grade = parseFloat(gradeMatch[0]);
      } else if (question === GOOGLE_FORM_EVALUATION_REMARKS_COLUMN) {
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

    const { evaluationWindowStart, evaluationWindowEnd } = getEvaluationWindowTimes();
    let isSupplementalEvaluation = false;
    let supplementalWindowForSubmitter = null;

    // Check if response is within original evaluation window
    if (responseTime < evaluationWindowStart || responseTime > evaluationWindowEnd) {
      // Check if this might be a supplemental evaluation
      Logger.log(
        `Evaluation received at ${responseTime} outside the original window from ${evaluationWindowStart} to ${evaluationWindowEnd}. Checking for supplemental evaluation windows.`
      );
      isSupplementalEvaluation = true;
    }

    // Retrieve assignments from Review Log and find expected submitters
    const allAssignments = getReviewLogAssignments(); // Get all assignments including supplemental
    const expectedSubmitters = [];

    // Check both original and supplemental assignments
    for (const [submitterEmail, evaluators] of Object.entries(allAssignments)) {
      if (evaluators.map(normalizeEmail).includes(normalizeEmail(evaluatorEmail))) {
        const submitterDiscord = lookupEmailAndDiscord(submitterEmail)?.discordHandle;
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

    const evaluatorDiscordHandle = lookupEmailAndDiscord(evaluatorEmail)?.discordHandle;
    if (!evaluatorDiscordHandle) {
      Logger.log(`Discord handle not found for evaluator email: ${evaluatorEmail}`);
      return;
    }
    Logger.log(`Evaluator Discord Handle: ${evaluatorDiscordHandle}`);

    // If this is a supplemental evaluation, validate it
    if (isSupplementalEvaluation) {
      const submitterEmail = lookupEmailAndDiscord(submitterDiscordHandle)?.email;
      if (!submitterEmail) {
        Logger.log(`Email not found for Discord handle: ${submitterDiscordHandle}`);
        return;
      }

      // Check if evaluator is assigned as a supplemental evaluator for this submitter
      // This also retrieves the window start/end times from the Review Log column header
      const supplementalAssignment = getSupplementalAssignmentForSubmitterAndEvaluator(submitterEmail, evaluatorEmail);
      if (!supplementalAssignment) {
        Logger.log(
          `Evaluator ${evaluatorEmail} is not assigned as a supplemental evaluator for ${submitterDiscordHandle}. Response ignored.`
        );
        return;
      }

      // Check if response is within the supplemental window (calculated from column header timestamp)
      if (responseTime < supplementalAssignment.windowStart || responseTime > supplementalAssignment.windowEnd) {
        Logger.log(
          `Supplemental evaluation from ${evaluatorEmail} for ${submitterDiscordHandle} is outside the supplemental window (${supplementalAssignment.windowStart} to ${supplementalAssignment.windowEnd}). Response ignored.`
        );
        return;
      }

      supplementalWindowForSubmitter = {
        start: supplementalAssignment.windowStart,
        end: supplementalAssignment.windowEnd,
      };

      Logger.log(
        `Supplemental evaluation validated for ${evaluatorEmail} evaluating ${submitterDiscordHandle} within supplemental window (${supplementalAssignment.windowStart} to ${supplementalAssignment.windowEnd}).`
      );
    }

    // Get the reporting month name from Request Log
    const reportingMonth = getReportingMonthFromRequestLog('Evaluation');
    if (!reportingMonth) {
      Logger.log('Error: Could not determine reporting month from Request Log.');
      return;
    }
    const monthSheet = scoresSpreadsheet.getSheetByName(reportingMonth.monthName);

    if (!monthSheet) {
      Logger.log(`Month sheet ${reportingMonth.monthName} not found.`);
      return;
    }

    // Find row for submitter
    const submitterDiscordColumnIndex = getRequiredColumnIndexByName(monthSheet, 'Submitter');
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

    // Get column indices dynamically
    const amb1Col = getRequiredColumnIndexByName(monthSheet, 'Amb-1');
    const score1Col = getRequiredColumnIndexByName(monthSheet, 'Score-1');
    const remarks1Col = getRequiredColumnIndexByName(monthSheet, 'Remarks-1');
    const amb2Col = getRequiredColumnIndexByName(monthSheet, 'Amb-2');
    const score2Col = getRequiredColumnIndexByName(monthSheet, 'Score-2');
    const remarks2Col = getRequiredColumnIndexByName(monthSheet, 'Remarks-2');
    const amb3Col = getRequiredColumnIndexByName(monthSheet, 'Amb-3');
    const score3Col = getRequiredColumnIndexByName(monthSheet, 'Score-3');
    const remarks3Col = getRequiredColumnIndexByName(monthSheet, 'Remarks-3');

    // Map evaluator columns to their respective score and remarks columns
    const evaluatorColumns = [
      { ambCol: amb1Col, scoreCol: score1Col, remarksCol: remarks1Col },
      { ambCol: amb2Col, scoreCol: score2Col, remarksCol: remarks2Col },
      { ambCol: amb3Col, scoreCol: score3Col, remarksCol: remarks3Col },
    ];

    if (isSupplementalEvaluation) {
      // For supplemental evaluations, find which Amb-[n] column has this evaluator pre-assigned
      // and validate they can write to that slot
      for (const { ambCol, scoreCol, remarksCol } of evaluatorColumns) {
        const ambValue = monthSheet.getRange(row, ambCol).getValue();
        const scoreValue = monthSheet.getRange(row, scoreCol).getValue();

        // Check if this evaluator is assigned to this slot
        if (ambValue === evaluatorDiscordHandle) {
          // Validate that we're not overwriting someone else's completed score
          if (typeof scoreValue === 'number' && !isNaN(scoreValue)) {
            Logger.log(
              `Warning: Supplemental evaluator ${evaluatorDiscordHandle} is trying to overwrite an existing score for ${submitterDiscordHandle}. This should not happen. Skipping.`
            );
            break;
          }

          // This is the correct slot for this evaluator - write the evaluation
          monthSheet.getRange(row, scoreCol).setValue(grade);
          monthSheet.getRange(row, remarksCol).setValue(remarks);

          // Apply COLOR_LATE_EVALUATION to mark this as a late evaluation
          monthSheet.getRange(row, ambCol).setBackground(COLOR_LATE_EVALUATION);
          monthSheet.getRange(row, scoreCol).setBackground(COLOR_LATE_EVALUATION);
          monthSheet.getRange(row, remarksCol).setBackground(COLOR_LATE_EVALUATION);

          Logger.log(
            `Supplemental evaluation processed: ${evaluatorDiscordHandle} evaluated ${submitterDiscordHandle}. Grade: ${grade}`
          );
          gradeUpdated = true;
          break;
        }
      }

      if (!gradeUpdated) {
        Logger.log(
          `Supplemental evaluator ${evaluatorDiscordHandle} is not assigned to evaluate ${submitterDiscordHandle} in any Amb slot. Response ignored.`
        );
      }
    } else {
      // For regular evaluations, match evaluator to their assigned column
      for (const { ambCol, scoreCol, remarksCol } of evaluatorColumns) {
        const cellValue = monthSheet.getRange(row, ambCol).getValue();
        if (cellValue === evaluatorDiscordHandle) {
          monthSheet.getRange(row, scoreCol).setValue(grade);
          monthSheet.getRange(row, remarksCol).setValue(remarks);
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

/**
 * Requests supplemental evaluations for submitters who did not receive 3 reviews.
 * Assigns 3 new evaluators randomly from the ambassador pool, excluding original evaluators
 * and non-responders from the first evaluation round.
 */
function requestSupplementalEvaluations() {
  try {
    Logger.log('Starting requestSupplementalEvaluations.');

    // Get the reporting month from Request Log
    const reportingMonth = getReportingMonthFromRequestLog('Evaluation');
    if (!reportingMonth) {
      alertAndLog('Error: Could not determine reporting month from Request Log.');
      return;
    }

    // Confirm with user
    const confirmationMessage = `This will request supplemental evaluations for ${reportingMonth.monthName}. Continue?`;
    const userConfirmed = promptAndLog('Confirm Supplemental Evaluation Request', confirmationMessage, ButtonSet.YES_NO);

    if (userConfirmed !== ButtonResponse.YES) {
      Logger.log('User cancelled supplemental evaluation request.');
      return;
    }

    const supplementalWindowStart = new Date(); // Capture start time for the supplemental evaluation window

    // Get submitters needing supplemental evaluations (those with < 3 reviews)
    const submittersNeedingReviews = getSubmittersNeedingSupplementalEvaluations(reportingMonth);

    if (submittersNeedingReviews.length === 0) {
      alertAndLog('All submitters have received 3 evaluations. No supplemental evaluations needed.');
      return;
    }

    Logger.log(`Found ${submittersNeedingReviews.length} submitters needing supplemental evaluations.`);

    // Get list of non-responders from original evaluation window
    const nonResponders = getNonResponders();
    Logger.log(`Non-responders from original window: ${nonResponders.join(', ')}`);

    // Get current assignments from Review Log
    const currentAssignments = getReviewLogAssignments();

    // Assign supplemental evaluators
    const supplementalAssignments = assignSupplementalEvaluators(
      submittersNeedingReviews,
      currentAssignments,
      nonResponders
    );

    if (supplementalAssignments.length === 0) {
      alertAndLog('Failed to generate supplemental assignments.');
      return;
    }

    // Add supplemental reviewer columns to Review Log with timestamp as header
    addSupplementalColumnsToReviewLog(supplementalAssignments, supplementalWindowStart);

    // Assign supplemental evaluators to Monthly Score sheet in their designated slots
    assignSupplementalEvaluatorsToMonthlySheet(supplementalAssignments, reportingMonth);

    // Send evaluation requests to supplemental evaluators
    sendSupplementalEvaluationRequests(supplementalAssignments, reportingMonth, supplementalWindowStart);

    alertAndLog(`Supplemental evaluation requests sent successfully for ${supplementalAssignments.length} submitter(s).`);
    Logger.log('requestSupplementalEvaluations completed.');
  } catch (error) {
    alertAndLog(`Error in requestSupplementalEvaluations: ${error.message}`);
    Logger.log(`Error in requestSupplementalEvaluations: ${error}`);
  }
}

/**
 * Gets the list of submitters who have received fewer than 3 evaluations.
 * @param {Object} reportingMonth - The reporting month object
 * @returns {Array<Object>} - Array of objects with submitterEmail, submitterDiscord, currentReviewCount, emptySlots, and emptySlotIndices
 */
function getSubmittersNeedingSupplementalEvaluations(reportingMonth) {
  const scoresSpreadsheet = getScoresSpreadsheet();
  const monthSheet = scoresSpreadsheet.getSheetByName(reportingMonth.monthName);

  if (!monthSheet) {
    Logger.log(`Month sheet ${reportingMonth.monthName} not found.`);
    return [];
  }

  const submittersNeedingReviews = [];
  const lastRow = monthSheet.getLastRow();

  if (lastRow < 2) {
    Logger.log('No data in month sheet.');
    return [];
  }

  // Get column indices
  const submitterCol = getRequiredColumnIndexByName(monthSheet, 'Submitter');
  const score1Col = getRequiredColumnIndexByName(monthSheet, 'Score-1');
  const score2Col = getRequiredColumnIndexByName(monthSheet, 'Score-2');
  const score3Col = getRequiredColumnIndexByName(monthSheet, 'Score-3');

  // Check each submitter row
  for (let row = 2; row <= lastRow; row++) {
    const submitterDiscord = monthSheet.getRange(row, submitterCol).getValue();
    if (!submitterDiscord) continue;

    // Check which scores are filled
    const score1 = monthSheet.getRange(row, score1Col).getValue();
    const score2 = monthSheet.getRange(row, score2Col).getValue();
    const score3 = monthSheet.getRange(row, score3Col).getValue();

    const scores = [score1, score2, score3];
    const emptySlotIndices = [];

    // Identify which specific slots are empty (1, 2, or 3)
    scores.forEach((score, index) => {
      if (!(typeof score === 'number' && !isNaN(score))) {
        emptySlotIndices.push(index + 1); // 1-based indexing (1, 2, 3)
      }
    });

    const reviewCount = 3 - emptySlotIndices.length;

    if (emptySlotIndices.length > 0) {
      // Look up email from discord handle
      const lookupResult = lookupEmailAndDiscord(submitterDiscord);
      if (lookupResult) {
        submittersNeedingReviews.push({
          submitterEmail: lookupResult.email,
          submitterDiscord: submitterDiscord,
          currentReviewCount: reviewCount,
          emptySlots: emptySlotIndices.length,
          emptySlotIndices: emptySlotIndices, // Array like [1, 3] if slots 1 and 3 are empty
        });
        Logger.log(`Submitter ${submitterDiscord} has ${reviewCount} reviews, needs ${emptySlotIndices.length} supplemental evaluation(s) in slot(s): ${emptySlotIndices.join(', ')}`);
      }
    }
  }

  return submittersNeedingReviews;
}

/**
 * Gets the list of evaluators who did not respond during the original evaluation window.
 * Only checks the original 3 reviewer columns (Reviewer 1, 2, 3), ignoring supplemental assignments.
 * Validates that evaluators actually evaluated someone they were assigned to.
 * @returns {Array<string>} - Array of evaluator emails who did not respond
 */
function getNonResponders() {
  try {
    const { evaluationWindowStart, evaluationWindowEnd } = getEvaluationWindowTimes();

    // Get all evaluators who were assigned in the ORIGINAL 3 reviewer columns only
    const registrySpreadsheet = getRegistrySpreadsheet();
    const reviewLogSheet = registrySpreadsheet.getSheetByName(REVIEW_LOG_SHEET_NAME);

    if (!reviewLogSheet) {
      Logger.log('Error: Review Log sheet not found.');
      return [];
    }

    // Find the column indices
    const submitterCol = getRequiredColumnIndexByName(reviewLogSheet, 'Submitter');
    const reviewer1Col = getRequiredColumnIndexByName(reviewLogSheet, 'Reviewer 1');
    const reviewer2Col = getRequiredColumnIndexByName(reviewLogSheet, 'Reviewer 2');
    const reviewer3Col = getRequiredColumnIndexByName(reviewLogSheet, 'Reviewer 3');

    const lastRow = reviewLogSheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('No assignments in Review Log.');
      return [];
    }

    // Build a map of evaluator → assigned submitters
    // evaluatorAssignments[evaluatorEmail] = [submitterEmail1, submitterEmail2, ...]
    const evaluatorAssignments = {};
    const allAssignedEvaluators = new Set();

    // Get all data from Review Log
    const reviewLogData = reviewLogSheet.getRange(2, 1, lastRow - 1, reviewLogSheet.getLastColumn()).getValues();

    reviewLogData.forEach((row) => {
      const submitterEmail = row[submitterCol - 1];
      const reviewers = [row[reviewer1Col - 1], row[reviewer2Col - 1], row[reviewer3Col - 1]];

      reviewers.forEach((reviewerEmail) => {
        if (reviewerEmail) {
          const normalizedReviewer = normalizeEmail(reviewerEmail);
          allAssignedEvaluators.add(normalizedReviewer);

          if (!evaluatorAssignments[normalizedReviewer]) {
            evaluatorAssignments[normalizedReviewer] = [];
          }
          evaluatorAssignments[normalizedReviewer].push(normalizeEmail(submitterEmail));
        }
      });
    });

    Logger.log(`Total original assigned evaluators: ${allAssignedEvaluators.size}`);

    // Get evaluators who responded AND evaluated someone they were assigned to
    const evaluationResponsesSheet = getEvaluationResponsesSheet();
    const lastResponseRow = evaluationResponsesSheet.getLastRow();

    if (lastResponseRow < 2) {
      // No responses at all, all assigned evaluators are non-responders
      return Array.from(allAssignedEvaluators);
    }

    const timestampCol = getRequiredColumnIndexByName(evaluationResponsesSheet, GOOGLE_FORM_TIMESTAMP_COLUMN);
    const emailCol = getRequiredColumnIndexByName(evaluationResponsesSheet, EVAL_FORM_USER_PROVIDED_EMAIL_COLUMN);
    const handleCol = getRequiredColumnIndexByName(evaluationResponsesSheet, GOOGLE_FORM_EVALUATION_HANDLE_COLUMN);

    const responses = evaluationResponsesSheet.getRange(2, 1, lastResponseRow - 1, evaluationResponsesSheet.getLastColumn()).getValues();

    const responders = new Set();
    responses.forEach((row) => {
      const timestamp = new Date(row[timestampCol - 1]);
      const evaluatorEmail = normalizeEmail(row[emailCol - 1]);
      const submitterDiscordHandle = String(row[handleCol - 1]).trim();

      // Check if response was within the original evaluation window
      if (timestamp >= evaluationWindowStart && timestamp <= evaluationWindowEnd && evaluatorEmail && submitterDiscordHandle) {
        // Look up the submitter's email from their Discord handle
        const submitterLookup = lookupEmailAndDiscord(submitterDiscordHandle);
        if (submitterLookup) {
          const submitterEmail = normalizeEmail(submitterLookup.email);

          // Check if this evaluator was assigned to evaluate this submitter
          const assignedSubmitters = evaluatorAssignments[evaluatorEmail] || [];
          if (assignedSubmitters.includes(submitterEmail)) {
            // Valid response - they evaluated someone they were assigned to
            responders.add(evaluatorEmail);
            Logger.log(`Valid response: ${evaluatorEmail} evaluated ${submitterDiscordHandle} (assigned)`);
          } else {
            Logger.log(`Invalid response: ${evaluatorEmail} evaluated ${submitterDiscordHandle} (NOT assigned - ignored)`);
          }
        }
      }
    });

    // Non-responders are those assigned but didn't respond (or responded to wrong person)
    const nonResponders = Array.from(allAssignedEvaluators).filter((evaluator) => !responders.has(evaluator));

    Logger.log(`Non-responders: ${nonResponders.join(', ')}`);
    return nonResponders;
  } catch (error) {
    Logger.log(`Error in getNonResponders: ${error}`);
    return [];
  }
}

/**
 * Assigns supplemental evaluators to submitters, excluding original evaluators and non-responders.
 * Only assigns the exact number of evaluators needed based on empty slots.
 * @param {Array<Object>} submittersNeedingReviews - Submitters needing supplemental evaluations
 * @param {Object} currentAssignments - Current assignments from Review Log
 * @param {Array<string>} nonResponders - List of non-responder emails
 * @returns {Array<Object>} - Array of supplemental assignments with emptySlotIndices
 */
function assignSupplementalEvaluators(submittersNeedingReviews, currentAssignments, nonResponders) {
  const supplementalAssignments = [];
  const allEvaluators = getEligibleAmbassadorsEmails();
  const normalizedNonResponders = nonResponders.map(normalizeEmail);

  submittersNeedingReviews.forEach((submitter) => {
    const { submitterEmail, submitterDiscord, emptySlots, emptySlotIndices } = submitter;

    // Get original evaluators for this submitter
    const originalEvaluators = (currentAssignments[submitterEmail] || []).map(normalizeEmail);

    // Filter available evaluators: exclude self, original evaluators, and non-responders
    const availableEvaluators = allEvaluators.filter(
      (evaluator) =>
        normalizeEmail(evaluator) !== normalizeEmail(submitterEmail) &&
        !originalEvaluators.includes(normalizeEmail(evaluator)) &&
        !normalizedNonResponders.includes(normalizeEmail(evaluator))
    );

    if (availableEvaluators.length < emptySlots) {
      Logger.log(
        `Not enough available evaluators for ${submitterDiscord}. Available: ${availableEvaluators.length}, need ${emptySlots}.`
      );
      alertAndLog(`Warning: Only ${availableEvaluators.length} evaluators available for ${submitterDiscord} (needs ${emptySlots})`);
    }

    // Randomly select only the number of evaluators needed (based on empty slots)
    const selectedEvaluators = [];
    const shuffled = availableEvaluators.sort(() => 0.5 - Math.random());

    for (let i = 0; i < Math.min(emptySlots, shuffled.length); i++) {
      selectedEvaluators.push(shuffled[i]);
    }

    if (selectedEvaluators.length > 0) {
      supplementalAssignments.push({
        submitterEmail,
        submitterDiscord,
        supplementalEvaluators: selectedEvaluators,
        emptySlotIndices: emptySlotIndices, // Pass through for Monthly Sheet assignment
      });
      Logger.log(`Assigned ${selectedEvaluators.length} supplemental evaluator(s) to ${submitterDiscord} for slot(s): ${emptySlotIndices.join(', ')}`);
    }
  });

  return supplementalAssignments;
}

/**
 * Adds supplemental reviewer columns to the Review Log with ISO 8601 timestamp as the header.
 * The timestamp in the header is used to calculate the evaluation window on-demand.
 * Adds only the number of columns needed based on the maximum number of evaluators assigned to any submitter.
 * @param {Array<Object>} supplementalAssignments - Supplemental assignments to add
 * @param {Date} supplementalWindowStart - The start time of the supplemental window
 */
function addSupplementalColumnsToReviewLog(supplementalAssignments, supplementalWindowStart) {
  const registrySpreadsheet = getRegistrySpreadsheet();
  const reviewLogSheet = registrySpreadsheet.getSheetByName(REVIEW_LOG_SHEET_NAME);

  if (!reviewLogSheet) {
    Logger.log('Error: Review Log sheet not found.');
    return;
  }

  const lastColumn = reviewLogSheet.getLastColumn();
  const lastRow = reviewLogSheet.getLastRow();

  // Format timestamp with timezone offset (ISO 8601 format): "2024-11-01T14:30:00-08:00"
  const timestampHeader = Utilities.formatDate(
    supplementalWindowStart,
    getProjectTimeZone(),
    "yyyy-MM-dd'T'HH:mm:ssXXX"
  );

  // Calculate the maximum number of supplemental evaluators assigned to any single submitter
  const maxEvaluatorsNeeded = Math.max(0, ...supplementalAssignments.map(a => a.supplementalEvaluators.length));
  Logger.log(`Adding ${maxEvaluatorsNeeded} supplemental column(s) to Review Log.`);

  // Add columns for this supplemental round with the same timestamp
  const newColumnStart = lastColumn + 1;

  for (let i = 0; i < maxEvaluatorsNeeded; i++) {
    const colIndex = newColumnStart + i;
    reviewLogSheet.getRange(1, colIndex).setValue(timestampHeader);
  }

  // Populate the new columns with supplemental evaluator assignments
  supplementalAssignments.forEach((assignment) => {
    const { submitterEmail, supplementalEvaluators } = assignment;

    // Find the row for this submitter
    const submitterCol = getRequiredColumnIndexByName(reviewLogSheet, 'Submitter');
    const submitterData = reviewLogSheet.getRange(2, submitterCol, lastRow - 1, 1).getValues();

    let submitterRow = null;
    for (let i = 0; i < submitterData.length; i++) {
      if (normalizeEmail(submitterData[i][0]) === normalizeEmail(submitterEmail)) {
        submitterRow = i + 2; // Offset for header row
        break;
      }
    }

    if (submitterRow) {
      // Write supplemental evaluators to the new columns
      supplementalEvaluators.forEach((evaluator, index) => {
        reviewLogSheet.getRange(submitterRow, newColumnStart + index).setValue(evaluator);
      });
      Logger.log(`Added ${supplementalEvaluators.length} supplemental evaluator(s) to Review Log for ${submitterEmail}`);
    }
  });

  Logger.log('Supplemental columns added to Review Log.');
}

/**
 * Assigns supplemental evaluators to their designated slots in the Monthly Score sheet.
 * Maps evaluators to specific Amb-[n] columns based on which slots are empty.
 * Preserves existing scores from original evaluators who responded.
 * @param {Array<Object>} supplementalAssignments - Supplemental assignments with emptySlotIndices
 * @param {Object} reportingMonth - The reporting month object
 */
function assignSupplementalEvaluatorsToMonthlySheet(supplementalAssignments, reportingMonth) {
  try {
    Logger.log('Assigning supplemental evaluators to Monthly Score sheet.');

    const scoresSpreadsheet = getScoresSpreadsheet();
    const monthSheet = scoresSpreadsheet.getSheetByName(reportingMonth.monthName);

    if (!monthSheet) {
      Logger.log(`Month sheet ${reportingMonth.monthName} not found.`);
      return;
    }

    // Get column indices
    const submitterCol = getRequiredColumnIndexByName(monthSheet, 'Submitter');
    const amb1Col = getRequiredColumnIndexByName(monthSheet, 'Amb-1');
    const amb2Col = getRequiredColumnIndexByName(monthSheet, 'Amb-2');
    const amb3Col = getRequiredColumnIndexByName(monthSheet, 'Amb-3');

    const ambColumns = [amb1Col, amb2Col, amb3Col];

    supplementalAssignments.forEach((assignment) => {
      const { submitterEmail, submitterDiscord, supplementalEvaluators, emptySlotIndices } = assignment;

      // Find the row for this submitter
      const lastRow = monthSheet.getLastRow();
      const submitterData = monthSheet.getRange(2, submitterCol, lastRow - 1, 1).getValues();

      let submitterRow = null;
      for (let i = 0; i < submitterData.length; i++) {
        if (submitterData[i][0] && submitterData[i][0].toLowerCase() === submitterDiscord.toLowerCase()) {
          submitterRow = i + 2; // Offset for header row
          break;
        }
      }

      if (!submitterRow) {
        Logger.log(`Submitter ${submitterDiscord} not found in month sheet.`);
        return;
      }

      // Map supplemental evaluators to their specific empty slots
      supplementalEvaluators.forEach((evaluatorEmail, index) => {
        if (index >= emptySlotIndices.length) {
          Logger.log(`Warning: More evaluators than empty slots for ${submitterDiscord}`);
          return;
        }

        const slotNumber = emptySlotIndices[index]; // 1, 2, or 3
        const ambColIndex = ambColumns[slotNumber - 1]; // Convert to 0-based index

        // Get evaluator's Discord handle
        const evaluatorDiscordHandle = lookupEmailAndDiscord(evaluatorEmail)?.discordHandle;
        if (!evaluatorDiscordHandle) {
          Logger.log(`Discord handle not found for evaluator: ${evaluatorEmail}`);
          return;
        }

        // Write evaluator Discord handle to the Amb-[n] column
        monthSheet.getRange(submitterRow, ambColIndex).setValue(evaluatorDiscordHandle);
        Logger.log(`Assigned ${evaluatorDiscordHandle} to ${submitterDiscord} in Amb-${slotNumber} (column ${ambColIndex})`);
      });
    });

    Logger.log('Supplemental evaluators assigned to Monthly Score sheet.');
  } catch (error) {
    Logger.log(`Error in assignSupplementalEvaluatorsToMonthlySheet: ${error}`);
  }
}

/**
 * Sends evaluation requests to supplemental evaluators.
 * @param {Array<Object>} supplementalAssignments - Supplemental assignments
 * @param {Object} reportingMonth - The reporting month object
 * @param {Date} supplementalWindowStart - The start time of the supplemental window
 */
function sendSupplementalEvaluationRequests(supplementalAssignments, reportingMonth, supplementalWindowStart) {
  const supplementalDeadline = new Date(supplementalWindowStart.getTime() + minutesToMilliseconds(EVALUATION_WINDOW_MINUTES));
  const supplementalDeadlineDate =
    Utilities.formatDate(supplementalDeadline, 'UTC', 'MMMM dd, yyyy HH:mm:ss') + ' UTC';

  supplementalAssignments.forEach((assignment) => {
    const { submitterEmail, submitterDiscord, supplementalEvaluators } = assignment;

    // Get contribution details
    const contributionDetails = getContributionDetailsByEmail(submitterEmail);
    const primaryTeam = getAmbassadorPrimaryTeam(submitterEmail);

    supplementalEvaluators.forEach((evaluatorEmail) => {
      try {
        const evaluatorDiscordHandle = lookupEmailAndDiscord(evaluatorEmail)?.discordHandle;
        if (!evaluatorDiscordHandle) {
          Logger.log(`Discord handle not found for evaluator: ${evaluatorEmail}`);
          return;
        }

        // Form message for supplemental evaluation request
        const message = REQUEST_EVALUATION_EMAIL_TEMPLATE.replace('{AmbassadorDiscordHandle}', evaluatorDiscordHandle)
          .replace('{Month}', reportingMonth.monthName)
          .replace('{AmbassadorSubmitter}', submitterDiscord)
          .replace('{SubmissionsList}', contributionDetails)
          .replace('{EvaluationFormURL}', EVALUATION_FORM_URL)
          .replace('{EVALUATION_DEADLINE_DATE}', supplementalDeadlineDate)
          .replace('{PrimaryTeam}', primaryTeam)
          .replace('{PrimaryTeamResponsibilities}', getPrimaryTeamResponsibilities(primaryTeam));

        sendEmailNotification(evaluatorEmail, '⚖️Supplemental Evaluation Request', message);
        Logger.log(`Sent supplemental evaluation request to ${evaluatorEmail} for submitter ${submitterDiscord}`);
      } catch (error) {
        Logger.log(`Error sending supplemental evaluation request to ${evaluatorEmail}: ${error}`);
      }
    });
  });

  Logger.log('Supplemental evaluation requests sent.');
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
    const reviewLogSheet = getReviewLogSheet();
    const formResponseSheet = getEvaluationResponsesSheet();

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
        const assignedEvaluations = reviewAssignments[submitterEmail]
          .map(normalizeEmail)
          .includes(normalizeEmail(evaluatorEmail))
          ? 1
          : 0;

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

    const registrySheet = getRegistrySheet();
    if (!registrySheet) {
      Logger.log('Registry sheet not found.');
      return;
    }

    // Get column indices dynamically
    const registryEmailColIndex = getRequiredColumnIndexByName(registrySheet, AMBASSADOR_EMAIL_COLUMN);
    const registryDiscordColIndex = getRequiredColumnIndexByName(registrySheet, AMBASSADOR_DISCORD_HANDLE_COLUMN);


    nonRespondents.forEach((evaluatorEmail) => {
      // Skip ambassadors who are not eligible (marked as 'Expelled' or not found)
      if (!eligibleEmails.map(normalizeEmail).includes(normalizeEmail(evaluatorEmail))) {
        Logger.log(`Skipping evaluator ${evaluatorEmail} (marked as 'Expelled' or not eligible).`);
        return;
      }

      // Validate email format
      if (!isValidEmail(evaluatorEmail)) {
        Logger.log(`Invalid email format for evaluator: "${evaluatorEmail}". Skipping.`);
        return;
      }

      const result = registrySheet.createTextFinder(evaluatorEmail).findNext(); // Find evaluator's row by email
      if (result) {
        const row = result.getRow();
        const discordHandle = registrySheet.getRange(row, registryDiscordColIndex).getValue();
        const email = registrySheet.getRange(row, registryEmailColIndex).getValue(); // Get email dynamically
        const message = REMINDER_EMAIL_TEMPLATE.replace('{AmbassadorDiscordHandle}', discordHandle);

        sendEmailNotification(email, '🕚Reminder to Submit Evaluation', message);
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
    const evaluationWindowEnd = new Date(evaluationWindowStart.getTime() + minutesToMilliseconds(EVALUATION_WINDOW_MINUTES));
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

/**
 * Function to reprocess all evaluation forms within all evaluation windows (original + supplemental)
 */
function batchProcessEvaluationResponses() {
  try {
    Logger.log('Starting batch processing of evaluation responses.');

    const form = FormApp.openById(EVALUATION_FORM_ID);
    if (!form) {
      Logger.log('Error: Form not found with the given ID.');
      return;
    }

    // Get all evaluation windows (original + supplemental)
    const allWindows = getAllEvaluationWindows();
    Logger.log(`Processing responses from ${allWindows.length} evaluation window(s).`);

    // Log all windows for debugging
    allWindows.forEach((window, index) => {
      Logger.log(`Window ${index + 1}: ${window.start} to ${window.end}`);
    });

    const formResponses = form.getResponses();
    Logger.log(`Total form responses retrieved: ${formResponses.length}`);

    // Filter responses by evaluation windows
    const filteredResponses = [];
    let rejectedCount = 0;

    formResponses.forEach((response, index) => {
      try {
        const timestamp = response.getTimestamp();
        const timestampDate = new Date(timestamp);

        // Check if response falls within any evaluation window
        const isInWindow = allWindows.some((window) => timestampDate >= window.start && timestampDate <= window.end);

        if (isInWindow) {
          filteredResponses.push(response);
        } else {
          rejectedCount++;
        }
      } catch (error) {
        Logger.log(`Error filtering response ${index + 1}: ${error}`);
        rejectedCount++;
      }
    });

    Logger.log(`Total responses INCLUDED: ${filteredResponses.length}`);
    Logger.log(`Total responses REJECTED: ${rejectedCount}`);

    const properties = PropertiesService.getScriptProperties();
    const lastProcessedIndex = parseInt(properties.getProperty('lastProcessedIndex') || '0', 10);
    const batchSize = 50; // Adjust batch size as needed

    Logger.log(`Processing batch: Starting from index ${lastProcessedIndex}, batch size ${batchSize}`);

    let processedCount = 0;
    let errorCount = 0;

    for (let i = lastProcessedIndex; i < filteredResponses.length && i < lastProcessedIndex + batchSize; i++) {
      try {
        const formResponse = filteredResponses[i];
        const event = { response: formResponse };
        processEvaluationResponse(event);
        processedCount++;
      } catch (error) {
        errorCount++;
        Logger.log(`Error processing response at index ${i}: ${error}`);
      }
    }

    Logger.log(`Batch processing summary: ${processedCount} responses processed successfully, ${errorCount} errors`);

    const newLastProcessedIndex = lastProcessedIndex + batchSize;
    if (newLastProcessedIndex < filteredResponses.length) {
      properties.setProperty('lastProcessedIndex', newLastProcessedIndex);
      ScriptApp.newTrigger('batchProcessEvaluationResponses').timeBased().after(1).create();
      Logger.log(`Processed batch. Next batch will start from index: ${newLastProcessedIndex}`);
      Logger.log(`Progress: ${newLastProcessedIndex} / ${filteredResponses.length} responses processed`);
    } else {
      properties.deleteProperty('lastProcessedIndex');
      deleteBatchProcessingTriggers();
      Logger.log('=== Batch Processing Completed ===');
      Logger.log(`Total responses in all windows: ${filteredResponses.length}`);
      Logger.log(`Total responses retrieved from form: ${formResponses.length}`);
      Logger.log(`Total responses rejected (outside windows): ${rejectedCount}`);
      Logger.log('Batch processing of evaluation responses completed.');
    }
  } catch (error) {
    Logger.log(`Error in batchProcessEvaluationResponses: ${error}`);
  }
}

function deleteBatchProcessingTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((trigger) => {
    if (trigger.getHandlerFunction() === 'batchProcessEvaluationResponses') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('Deleted batch processing trigger.');
    }
  });
}
