function batchProcessJune2025Responses() {
  try {
    Logger.log('Starting batch re-processing of June 2025 evaluations.');

    const form = FormApp.openById(EVALUATION_FORM_ID);
    if (!form) {
      Logger.log('Error: Form not found with the given ID.');
      return;
    }

    const evaluationWindowStart = new Date('2025-08-04 18:27:39 PDT');
    const evaluationWindowEnd = new Date('2025-08-11 18:27:39 PDT');
    Logger.log(`Evaluation window: ${evaluationWindowStart} to ${evaluationWindowEnd}`);

    const formResponses = form.getResponses();
    const filteredResponses = formResponses.filter((response) => {
      const timestamp = new Date(response.getTimestamp());
      return timestamp >= evaluationWindowStart && timestamp <= evaluationWindowEnd;
    });

    Logger.log(`Total form responses to process: ${filteredResponses.length}`);

    const properties = PropertiesService.getScriptProperties();
    const lastProcessedIndex = parseInt(properties.getProperty('lastProcessedIndex') || '0', 10);
    const batchSize = 50; // Adjust batch size as needed

    for (let i = lastProcessedIndex; i < filteredResponses.length && i < lastProcessedIndex + batchSize; i++) {
      const formResponse = filteredResponses[i];
      const event = { response: formResponse };
      manualProcessEvaluationResponse(event, evaluationWindowStart, evaluationWindowEnd);
    }

    const newLastProcessedIndex = lastProcessedIndex + batchSize;
    if (newLastProcessedIndex < filteredResponses.length) {
      properties.setProperty('lastProcessedIndex', newLastProcessedIndex);
      ScriptApp.newTrigger('batchProcessJune2025Responses').timeBased().after(1).create();
      Logger.log(`Processed batch. Next batch will start from index: ${newLastProcessedIndex}`);
    } else {
      properties.deleteProperty('lastProcessedIndex');
      deleteBatchProcessingTriggers();
      Logger.log('Batch processing of evaluation responses completed.');
    }
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
function manualProcessEvaluationResponse(e, evaluationWindowStart, evaluationWindowEnd) {
  try {
    Logger.log('manualProcessEvaluationResponse triggered.');

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

    if (responseTime < evaluationWindowStart || responseTime > evaluationWindowEnd) {
      Logger.log(
        `Evaluation received at ${responseTime} outside the window from ${evaluationWindowStart} to ${evaluationWindowEnd}. Response will be ignored.`
      );
      return;
    }

    // Retrieve assignments from Review Log and find expected submitters
    const assignments = getJune2025ReviewLogAssignments();
    const expectedSubmitters = [];

    for (const [submitterEmail, evaluators] of Object.entries(assignments)) {
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

    // Get the reporting month name from Request Log
    const reportingMonth = {
      month: 'June',
      year: '2025',
      monthName: `June 2025`,
      firstDayDate: new Date('June 01, 2025'),
    };

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
 * Fetches and returns the submitter-evaluator assignments from the Review Log.
 * Dynamically determines column indices based on header names to avoid hardcoded indices.
 * @returns {Object} - A map of submitter emails to a list of evaluator emails.
 */
function getJune2025ReviewLogAssignments() {
  Logger.log('Fetching submitter-evaluator assignments from Review Log.');

  const reviewLogSheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(
    'June 2025 Review Log'
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
  return assignments;
}
