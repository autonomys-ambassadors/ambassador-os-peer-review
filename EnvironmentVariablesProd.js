function setProductionVariables() {
  // controls wether email will be sent or just logged for troubleshooting - should always be true for production.
  SEND_EMAIL = true;

  // Real sheets:
  //
  // Provide the Id of the google sheet for the registry and scoreing sheets:
  AMBASSADOR_REGISTRY_SPREADSHEET_ID = ''; //"Ambassador Registry"
  AMBASSADORS_SCORES_SPREADSHEET_ID = ''; // "Ambassadors' Scores"

  // Provide the Id and submission URL for the submission and evaluation google forms:
  SUBMISSION_FORM_ID = ''; // ID for Submission form
  EVALUATION_FORM_ID = ''; // ID for Evaluation form
  SUBMISSION_FORM_URL = ''; // Submission Form URL for mailing
  EVALUATION_FORM_URL = ''; // Evaluation Form URL for mailing
  FORM_RESPONSES_SHEET_NAME = 'Form Responses'; // Explicit name for 'Form Responses' sheet
  EVAL_FORM_RESPONSES_SHEET_NAME = 'Form Responses 2'; // Evaluation Form responses sheet

  // Triggers and Delays
  // These values will set the due date and reminder schedule for Submissions and Evaluations.
  // The Submission or Evaluation will be due ofter the relevant WINDOW_MINUTES,
  // and each ambassador will receive a reminder after the relevant WINDOW_REMINDER_MINUTES.
  // specifies as days * hours * minutes
  SUBMISSION_WINDOW_MINUTES = 7 * 24 * 60;
  SUBMISSION_WINDOW_REMINDER_MINUTES = 5 * 24 * 60; // how many minutes after Submission Requests sent to remind
  EVALUATION_WINDOW_MINUTES = 7 * 24 * 60;
  EVALUATION_WINDOW_REMINDER_MINUTES = 5 * 24 * 60; // how many minutes after Evaluation Requests sent to remind

  // Sheet names
  REGISTRY_SHEET_NAME = 'Registry';
  FORM_RESPONSES_SHEET_NAME = 'Form Responses';
  REVIEW_LOG_SHEET_NAME = 'Review Log';
  CONFLICT_RESOLUTION_TEAM_SHEET_NAME = 'Conflict Resolution Team';
  OVERALL_SCORE_SHEET_NAME = 'Overall score'; // Overall score sheet in Ambassadors' Scores
  EVAL_FORM_RESPONSES_SHEET_NAME = 'Form Responses 2'; // Evaluation Form responses sheet

  // Columns
  AMBASSADOR_EMAIL_COLUMN = 'Ambassador Email Address';
  AMBASSADOR_DISCORD_HANDLE_COLUMN = 'Ambassador Discord Handle';

  // Sponsor Email (for notifications when ambassadors are expelled)
  SPONSOR_EMAIL = 'community@autonomys.xyz'; // Sponsor's email

  // Reinitialize color variables to ensure consistency in color-based logic.
  COLOR_MISSED_SUBMISSION = '#f5eee6';
  COLOR_MISSED_EVALUATION = '#e6d6c1';
  COLOR_EXPELLED = '#FF0000';
  COLOR_MISSED_SUBM_AND_EVAL = '#ceae83';
  COLOR_OLD_MISSED_SUBMISSION = '#f5eee6';
}