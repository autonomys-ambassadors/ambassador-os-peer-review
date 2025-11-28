/**
 * Configuration template
 * Copy this file to Config-[YourName].js and customize for your environment
 *
 * Instructions:
 * 1. Replace [YourName] with your actual name in the filename
 * 2. Update the function name to match: setYourNameVariables()
 * 3. Replace placeholder values with your actual environment values
 * 4. Update SharedUtilities.js CONFIG_NAME to include your configuration name
 */

function setYourNameVariables() {
  TESTING = true;
  SEND_EMAIL = true; // Set to false if you don't want emails sent during testing
  TESTER_EMAIL = 'your-email@example.com'; // Your email for testing redirects

  // Spreadsheets - Replace with your test spreadsheet IDs
  AMBASSADOR_REGISTRY_SPREADSHEET_ID = 'your-registry-spreadsheet-id';
  AMBASSADORS_SCORES_SPREADSHEET_ID = 'your-scores-spreadsheet-id';
  AMBASSADORS_SUBMISSIONS_SPREADSHEET_ID = 'your-submissions-spreadsheet-id';
  EVALUATION_RESPONSES_SPREADSHEET_ID = 'your-evaluation-responses-spreadsheet-id';
  ANONYMOUS_SCORES_SPREADSHEET_ID = 'your-anonymous-scores-spreadsheet-id';
  FORM_RESPONSES_SHEET_NAME = 'Form Responses 1'; // Adjust sheet name as needed
  EVAL_FORM_RESPONSES_SHEET_NAME = 'Form Responses 1'; // Adjust sheet name as needed

  // Google Forms - Replace with your test form IDs and URLs
  SUBMISSION_FORM_ID = 'your-submission-form-id';
  EVALUATION_FORM_ID = 'your-evaluation-form-id';
  SUBMISSION_FORM_URL = 'https://forms.gle/your-submission-form-url';
  EVALUATION_FORM_URL = 'https://forms.gle/your-evaluation-form-url';

  // Sponsor Email
  SPONSOR_EMAIL = 'sponsor@example.com';

  // Sheet names - Adjust if your test sheets have different names
  REGISTRY_SHEET_NAME = 'Registry';
  REVIEW_LOG_SHEET_NAME = 'Review Log';
  CONFLICT_RESOLUTION_TEAM_SHEET_NAME = 'Conflict Resolution Team';
  OVERALL_SCORE_SHEET_NAME = 'Overall Score';

  // Column names - Adjust to match your test sheet column headers
  AMBASSADOR_ID_COLUMN = 'Ambassador Id';
  AMBASSADOR_EMAIL_COLUMN = 'Ambassador Email Address';
  AMBASSADOR_DISCORD_HANDLE_COLUMN = 'Ambassador Discord Handle';
  AMBASSADOR_STATUS_COLUMN = 'Ambassador Status';
  AMBASSADOR_PRIMARY_TEAM_COLUMN = 'Primary Team';
  GOOGLE_FORM_TIMESTAMP_COLUMN = 'Timestamp';
  SUBM_FORM_USER_PROVIDED_EMAIL_COLUMN = 'Email Address';
  EVAL_FORM_USER_PROVIDED_EMAIL_COLUMN = 'Email Address';
  GOOGLE_FORM_REAL_EMAIL_COLUMN = 'Email Address';
  GOOGLE_FORM_CONTRIBUTION_DETAILS_COLUMN = `Dear Ambassador,
Please add text to your contributions during the month`;
  GOOGLE_FORM_CONTRIBUTION_LINKS_COLUMN = `Dear Ambassador,
Please add links to your contributions during the month`;
  GOOGLE_FORM_EVALUATION_HANDLE_COLUMN = 'Discord handle of the ambassador you are evaluating?';
  GOOGLE_FORM_EVALUATION_GRADE_COLUMN = 'Please assign a grade on a scale of 0 to 5';
  GOOGLE_FORM_EVALUATION_REMARKS_COLUMN = 'Remarks';
  SCORE_PENALTY_POINTS_COLUMN = 'Penalty Points';
  SCORE_AVERAGE_SCORE_COLUMN = 'Average Score';
  SCORE_MAX_6M_PP_COLUMN = 'Max 6-Month PP';
  SUBMITTER_HANDLE_COLUMN_IN_MONTHLY_SCORE = 'Submitter';
  SUBMITTER_HANDLE_COLUMN_IN_REVIEW_LOG = 'Submitter';
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

  // Notion Configuration
  NOTION_DATABASE_ID = 'your-notion-database-id'; // Replace with your Notion database ID
  NOTION_NUMBER_COLUMN = 'Number (Unique ID)';
  NOTION_EMAIL_COLUMN = 'Email';
  NOTION_DISCORD_COLUMN = 'Discord Handle';
  NOTION_STATUS_COLUMN = 'Status';
  NOTION_PRIMARY_TEAM_COLUMN = 'Team (Guild)';
  NOTION_SECONDARY_TEAM_COLUMN = 'Secondary Team (Guild)';
  NOTION_START_DATE_COLUMN = 'Start Date';

  // New Registry Sheet Columns
  REGISTRY_NOTION_ID_COLUMN = 'Notion Id';
  REGISTRY_SECONDARY_TEAM_COLUMN = 'Secondary Team';
  REGISTRY_START_DATE_COLUMN = 'Start Date';

  // Testing timing - Use shorter windows for accelerated testing
  SUBMISSION_WINDOW_MINUTES = 15; // How long submission window stays open
  SUBMISSION_WINDOW_REMINDER_MINUTES = 5; // When to send reminders after submission requests
  EVALUATION_WINDOW_MINUTES = 15; // How long evaluation window stays open
  EVALUATION_WINDOW_REMINDER_MINUTES = 5; // When to send reminders after evaluation requests

  // Thresholds for testing
  MAX_PENALTY_POINTS_TO_EXPEL = 3; // Penalty points threshold for expulsion
  MAX_INADEQUATE_CONTRIBUTION_COUNT_TO_REFER = 2; // Inadequate contributions before CRT referral
  INADEQUATE_CONTRIBUTION_SCORE_THRESHOLD = 3.0; // Score threshold for inadequate contribution

  // Color codes for cell backgrounds (must be lowercase hex)
  COLOR_MISSED_SUBMISSION = '#ead1dc';
  COLOR_MISSED_EVALUATION = '#d9d9d9';
  COLOR_EXPELLED = '#d36a6a';
  COLOR_MISSED_SUBM_AND_EVAL = '#ea9999';
  COLOR_LATE_EVALUATION = '#fff2cc';
}
