/**
 * Test configuration for Wilyam
 * Contains all testing variables and settings specific to Wilyam's test environment
 */
function setWilyamVariables() {
  TESTING = true;
  SEND_EMAIL = true;
  TESTER_EMAIL = 'economicsilver@starmail.net'; // Wilyam's email for testing redirects

  // Specify your testing sheets/forms/etc. here:
  // Spreadsheets:
  //
  AMBASSADOR_REGISTRY_SPREADSHEET_ID = '1AIMD61YKfk-JyP6Aia3UWke9pW15bPYTvWc1C46ofkU'; //"Ambassador Registry"
  AMBASSADORS_SCORES_SPREADSHEET_ID = '1RJzCo1FgGWkx0UCYhaY3SIOR0sl7uD6vCUFw55BX0iQ'; // "Ambassadors' Scores"
  AMBASSADORS_SUBMISSIONS_SPREADSHEET_ID = '1AIMD61YKfk-JyP6Aia3UWke9pW15bPYTvWc1C46ofkU'; // "Ambassador Submission Responses"
  EVALUATION_RESPONSES_SPREADSHEET_ID = '1RJzCo1FgGWkx0UCYhaY3SIOR0sl7uD6vCUFw55BX0iQ'; // "Ambassador Evaluations' Responses"
  FORM_RESPONSES_SHEET_NAME = 'Form Responses 14'; // Explicit name for 'Form Responses' sheet
  EVAL_FORM_RESPONSES_SHEET_NAME = 'Form Responses 2'; // Evaluation Form responses sheet

  // Google Forms
  //
  SUBMISSION_FORM_ID = '1SV5rJbzPv6BQgDZkC_xgrauWgoKPcEmtk3aKY6f4ZC8'; // ID for Submission form
  EVALUATION_FORM_ID = '15UXnrpOOoZPO7XCP2TV7mwezewHY6UIsYAU_W_aoMwo'; // ID for Evaluation form
  SUBMISSION_FORM_URL = 'https://forms.gle/beZrwuP9Zs1HvUY49'; // Submission Form URL for mailing
  EVALUATION_FORM_URL = 'https://forms.gle/kndReXQqXT6JyKX68'; // Evaluation Form URL for mailing

  // Sponsor Email (for notifications when ambassadors are expelled)
  //
  SPONSOR_EMAIL = 'economicsilver@starmail.net'; // Sponsor's email

  // Sheet names
  REGISTRY_SHEET_NAME = 'Registry';
  REVIEW_LOG_SHEET_NAME = 'Review Log';
  CONFLICT_RESOLUTION_TEAM_SHEET_NAME = 'Conflict Resolution Team';
  OVERALL_SCORE_SHEET_NAME = 'Overall score'; // Overall score sheet in Ambassadors' Scores

  // Columns
  AMBASSADOR_ID_COLUMN = 'Ambassador Id';
  AMBASSADOR_EMAIL_COLUMN = 'Ambassador Email Address';
  AMBASSADOR_DISCORD_HANDLE_COLUMN = 'Ambassador Discord Handle';
  AMBASSADOR_STATUS_COLUMN = 'Ambassador Status';
  AMBASSADOR_PRIMARY_TEAM_COLUMN = 'Primary Team';
  GOOGLE_FORM_TIMESTAMP_COLUMN = 'Timestamp';
  SUBM_FORM_USER_PROVIDED_EMAIL_COLUMN = 'Your Email Address';
  EVAL_FORM_USER_PROVIDED_EMAIL_COLUMN = 'Your Email Address';
  GOOGLE_FORM_REAL_EMAIL_COLUMN = 'Email Address';
  GOOGLE_FORM_CONTRIBUTION_DETAILS_COLUMN = `Dear Ambassador,
Please add text to your contributions during the month`;
  GOOGLE_FORM_CONTRIBUTION_LINKS_COLUMN = `Dear Ambassador,
Please add links your contributions during the month `;
  GOOGLE_FORM_EVALUATION_HANDLE_COLUMN = 'Discord handle of the ambassador you are evaluating?';
  GOOGLE_FORM_EVALUATION_GRADE_COLUMN = 'Please assign a grade on a scale of 0 to 5';
  GOOGLE_FORM_EVALUATION_REMARKS_COLUMN = 'Remarks';
  SCORE_PENALTY_POINTS_COLUMN = 'Penalty Points';
  SCORE_AVERAGE_SCORE_COLUMN = 'Average Score';
  SCORE_MAX_6M_PP_COLUMN = 'Max 6-Month PP';
  SUBMITTER_HANDLE_COLUMN_IN_MONTHLY_SCORE = 'Submitter';
  GRADE_FINAL_SCORE_COLUMN = 'Final Score';
  GRADE_EVAL_1_SCORE_COLUMN = 'Score-1';
  GRADE_EVAL_1_REMARKS_COLUMN = 'Remarks-1';
  GRADE_EVAL_2_SCORE_COLUMN = 'Score-2';
  GRADE_EVAL_2_REMARKS_COLUMN = 'Remarks-2';
  GRADE_EVAL_3_SCORE_COLUMN = 'Score-3';
  GRADE_EVAL_3_REMARKS_COLUMN = 'Remarks-3';
  CRT_SELECTION_DATE_COLUMN = 'Selection Date';
  SCORE_INADEQUATE_CONTRIBUTION_COLUMN = 'Inadequate Contribution Count';
  SCORE_CRT_REFERRAL_HISTORY_COLUMN = 'CRT Referral History';

  // Request Log columns
  REQUEST_LOG_REQUEST_TYPE_COLUMN = 'Type';
  REQUEST_LOG_MONTH_COLUMN = 'Month';
  REQUEST_LOG_YEAR_COLUMN = 'Year';
  REQUEST_LOG_START_TIME_COLUMN = 'Request Date Time';
  REQUEST_LOG_END_TIME_COLUMN = 'Window End Date Time';

  // Triggers and Delays for testing - use much shorter windows for accelerated testing schedules
  SUBMISSION_WINDOW_MINUTES = 15;
  SUBMISSION_WINDOW_REMINDER_MINUTES = 5; // how many minutes after Submission Requests sent to remind
  EVALUATION_WINDOW_MINUTES = 10;
  EVALUATION_WINDOW_REMINDER_MINUTES = 5; // how many minutes after Evaluation Requests sent to remind

  // Penalty Points threshold - if > or = this number for the past 6 months, ambassador will be expelled
  MAX_PENALTY_POINTS_TO_EXPEL = 3;
  MAX_INADEQUATE_CONTRIBUTION_COUNT_TO_REFER = 2;
  INADEQUATE_CONTRIBUTION_SCORE_THRESHOLD = 3.0;

  /** Reinitialize color variables to ensure consistency in color-based logic.
   * The color hex string must be in lowercase!
   */
  COLOR_MISSED_SUBMISSION = '#ead1dc';
  COLOR_MISSED_EVALUATION = '#d9d9d9';
  COLOR_EXPELLED = '#d36a6a';
  COLOR_MISSED_SUBM_AND_EVAL = '#ea9999';


  //COLOR_OLD_MISSED_SUBMISSION = '#ead1dc';

  // Wilyam Test vars
  //
  //
}
