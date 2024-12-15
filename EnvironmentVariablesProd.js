function setProductionVariables() {
  // Always send emails in production

  SEND_EMAIL = true;
  // Real sheets:
  //
  // Provide the Id of the google sheet for the registry and scoreing sheets:
  AMBASSADOR_REGISTRY_SPREADSHEET_ID = '1YtE-b7088aV3zi0eyFaGMyA7Nvo3bf9dnl0xzH3BTdA'; //"Ambassador Registry", also where the app is run from
  AMBASSADORS_SCORES_SPREADSHEET_ID = '1cjhrqgc84HdS59eQJPsiNIPKbusHtp2j7dN55u-mKdc'; // "Ambassadors Scores"
  AMBASSADORS_SUBMISSIONS_SPREADSHEET_ID = '1EQRSjcvODXQpHzK2g4XNd6imCTsx2m7vTwe-HeNNjjM'; // "Ambassador Submission Responses"
  EVALUATION_RESPONSES_SPREADSHEET_ID = '12S_qu-Uiq0BupN_Z6lVHJ76YaVM5NG0JQY_IC-gFWmg'; // "Ambassador Evaluations' Responses"

  // Provide the Id and submission URL for the submission and evaluation google forms:
  SUBMISSION_FORM_ID = '1mBTic1KtJRaXB93YDRTFMRta6gLIcAQglHh2LWwN8XE'; // ID for Submission form
  EVALUATION_FORM_ID = '1WKQ1acvwVVXJOtYRZgiX-4YXOUgnwwlquL3c5l494ew'; // ID for Evaluation form
  SUBMISSION_FORM_URL = 'https://forms.gle/jU6u22fycgQjQ3z68'; // Submission Form URL for mailing
  EVALUATION_FORM_URL = 'https://forms.gle/MfRt9G8WdvhgVRca6'; // Evaluation Form URL for mailing
  FORM_RESPONSES_SHEET_NAME = 'Form Responses 1'; // Explicit name for 'Form Responses' sheet
  EVAL_FORM_RESPONSES_SHEET_NAME = 'Form Responses 1'; // Evaluation Form responses sheet

  // Triggers and Delays
  // These values will set the due date and reminder schedule for Submissions and Evaluations.
  // The Submission or Evaluation will be due after the relevant WINDOW_MINUTES,
  // and each ambassador will receive a reminder after the relevant WINDOW_REMINDER_MINUTES.
  // specifies as days * hours * minutes
  SUBMISSION_WINDOW_MINUTES = 7 * 24 * 60;
  SUBMISSION_WINDOW_REMINDER_MINUTES = 5 * 24 * 60; // time (expressed in minutes) after Submission Requests sent to remind
  EVALUATION_WINDOW_MINUTES = 7 * 24 * 60;
  EVALUATION_WINDOW_REMINDER_MINUTES = 5 * 24 * 60; // time (expressed in minutes) after Evaluation Requests sent to remind

  // Sheet names
  REGISTRY_SHEET_NAME = 'Registry';
  REVIEW_LOG_SHEET_NAME = 'Review Log';
  CONFLICT_RESOLUTION_TEAM_SHEET_NAME = 'Conflict Resolution Team';
  OVERALL_SCORE_SHEET_NAME = 'Overall score'; // Overall score sheet in Ambassadors' Scores

  // Columns
  AMBASSADOR_EMAIL_COLUMN = 'Ambassador Email Address';
  AMBASSADOR_DISCORD_HANDLE_COLUMN = 'Ambassador Discord Handle';
  AMBASSADOR_STATUS_COLUMN = 'Ambassador Status';

  // Sponsor Email (for notifications when ambassadors are expelled)
  SPONSOR_EMAIL = 'community@autonomys.xyz'; // Sponsor's email

  // Penalty Points threshold - if > or = this number for the past 6 months, ambassador will be expelled
  MAX_PENALTY_POINTS_TO_EXPEL = 3;

  /** Reinitialize color variables to ensure consistency in color-based logic.
   * The color hex string must be in lowercase!
   */
  COLOR_MISSED_SUBMISSION = '#ead1dc';
  COLOR_MISSED_EVALUATION = '#d9d9d9';
  COLOR_EXPELLED = '#d36a6a';
  COLOR_MISSED_SUBM_AND_EVAL = '#ea9999';
  COLOR_OLD_MISSED_SUBMISSION = '#ead1dc';
}
