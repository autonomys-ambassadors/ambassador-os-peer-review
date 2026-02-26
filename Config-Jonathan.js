/**
 * Test configuration for Jonathan
 * Contains all testing variables and settings specific to Jonathan's test environment
 */
function setJonathanVariables() {
  TESTING = true;
  SEND_EMAIL = true;
  TESTER_EMAIL = 'ambassadoros+tester@jkw.fm'; // Jonathan's email for testing redirects
  SPONSOR_EMAIL = 'ambassadoros+tester+sponsor@jkw.fm'; // Sponsor's email

  /* SPECIFICALLY FOR TESTING WITH PRODUCTION DATA */
  AMBASSADOR_REGISTRY_SPREADSHEET_ID = '1vBBZaNPlcvWwoaLOGtFlirntvRMgtSdfViDZvJ-wCJ4'; //"Ambassador Registry", also where the app is run from
  AMBASSADORS_SCORES_SPREADSHEET_ID = '1vAY6GWO9bU2xyiO_-a-qt5AjGYpm86zyi1CqEWJuGXM'; // "Ambassadors Scores"
  AMBASSADORS_SUBMISSIONS_SPREADSHEET_ID = '1lVUaCGCCbfD3l9e8MEfQVBKaljm7A5aKX7RJAsUrWfA'; // "Ambassador Submission Responses"
  EVALUATION_RESPONSES_SPREADSHEET_ID = '1lVUaCGCCbfD3l9e8MEfQVBKaljm7A5aKX7RJAsUrWfA'; // "Test Evaluation Form Responses"
  ANONYMOUS_SCORES_SPREADSHEET_ID = '1znTe1E0gXpxHfK3GcL1ToIReo6nwfG_cwaqk7AdKzHQ'; // Your anonymous scores spreadsheet ID
  SUBMISSION_FORM_ID = '13oDRgD2qjryfhv992ZS99zCTOHPXBxsqKAXijupHbfE'; // ID for Submission form
  EVALUATION_FORM_ID = '1EPrKCrg7NXfEje3Ps3aBZ3S_qs9oy3dZl6SbnF3Ek4U'; // ID for test Adequate/Inadequate Evaluation form

  SUBMISSION_FORM_URL = 'https://forms.gle/SkeQnXAT95XtDGa19'; // Submission Form URL for mailing
  EVALUATION_FORM_URL = 'https://forms.gle/WEk1ca1sfeMqycDc8'; // Test Adequate-Inadequate Evaluation Form URL
  FORM_RESPONSES_SHEET_NAME = 'Form Responses 3'; // Explicit name for 'Form Responses' sheet
  EVAL_FORM_RESPONSES_SHEET_NAME = 'Form Responses 4'; // Evaluation Form responses sheet

  // Sheet names
  REGISTRY_SHEET_NAME = 'Registry';
  REVIEW_LOG_SHEET_NAME = 'Review Log';
  CONFLICT_RESOLUTION_TEAM_SHEET_NAME = 'Conflict Resolution Team';
  OVERALL_SCORE_SHEET_NAME = 'Overall Score'; // Pre-scored test data with Adequate/Inadequate/No Value
  CRT_LOG_SHEET_NAME = 'CRT Log'; // CRT Log sheet for tracking conflict resolution referrals

  // Columns
  AMBASSADOR_ID_COLUMN = 'Ambassador Id';
  AMBASSADOR_EMAIL_COLUMN = 'Ambassador Email Address';
  AMBASSADOR_DISCORD_HANDLE_COLUMN = 'Ambassador Discord Handle';
  AMBASSADOR_STATUS_COLUMN = 'Ambassador Status';
  AMBASSADOR_PRIMARY_TEAM_COLUMN = 'Primary Team';
  GOOGLE_FORM_TIMESTAMP_COLUMN = 'Timestamp';
  SUBM_FORM_USER_PROVIDED_EMAIL_COLUMN = 'Email';
  EVAL_FORM_USER_PROVIDED_EMAIL_COLUMN = 'Email';
  GOOGLE_FORM_REAL_EMAIL_COLUMN = 'Email Address';
  GOOGLE_FORM_CONTRIBUTION_DETAILS_COLUMN = `Dear Ambassador,
Please add text to your contributions during the month`;
  GOOGLE_FORM_CONTRIBUTION_LINKS_COLUMN = `Dear Ambassador,
Please add links to your contributions during the month`;
  GOOGLE_FORM_EVALUATION_HANDLE_COLUMN = 'Discord handle of the ambassador you are evaluating?'; //values must match google form questions
  GOOGLE_FORM_EVALUATION_GRADE_COLUMN = 'Please score the ambassadors contributions for the relevant month.';
  GOOGLE_FORM_EVALUATION_REMARKS_COLUMN = 'Remarks (required)';
  SCORE_PENALTY_POINTS_COLUMN = 'Penalty Points Last 6 Months';
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

  // CRT Log columns
  CRT_LOG_EMAIL_COLUMN = 'Email Address';
  CRT_LOG_DISCORD_HANDLE_COLUMN = 'Discord Handle';
  CRT_LOG_REFERRAL_DATE_COLUMN = 'Referral Date';
  CRT_LOG_REASON_COLUMN = 'Reason';
  CRT_LOG_RESOLUTION_DATE_COLUMN = 'Resolution Date';
  CRT_LOG_RESOLUTION_COLUMN = 'Resolution';

  // Request Log columns
  REQUEST_LOG_REQUEST_TYPE_COLUMN = 'Type';
  REQUEST_LOG_MONTH_COLUMN = 'Month';
  REQUEST_LOG_YEAR_COLUMN = 'Year';
  REQUEST_LOG_START_TIME_COLUMN = 'Request Date Time';
  REQUEST_LOG_END_TIME_COLUMN = 'Window End Date Time';

  // Notion Configuration
  NOTION_DATABASE_ID = '1f289b830b398029b2c7f516fbb19e5c'; // Notion database ID
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

  // Triggers and Delays for testing - use much shorter windows for accelerated testing schedules
  //SUBMISSION_WINDOW_MINUTES = 20;
  //SUBMISSION_WINDOW_REMINDER_MINUTES = 15; // how many minutes after Submission Requests sent to remind
  //EVALUATION_WINDOW_MINUTES = 20;
  //EVALUATION_WINDOW_REMINDER_MINUTES = 15; // how many minutes after Evaluation Requests sent to remind
  SUBMISSION_WINDOW_MINUTES = 7 * 24 * 60;
  SUBMISSION_WINDOW_REMINDER_MINUTES = 5 * 24 * 60; // time (expressed in minutes) after Submission Requests sent to remind
  EVALUATION_WINDOW_MINUTES = 7 * 24 * 60;
  EVALUATION_WINDOW_REMINDER_MINUTES = 5 * 24 * 60; // time (expressed in minutes) after Evaluation Requests sent to remind

  // Penalty Points threshold - if > or = this number for the past 6 months, ambassador will be expelled
  MAX_PENALTY_POINTS_TO_EXPEL = 3;
  MAX_INADEQUATE_CONTRIBUTION_COUNT_TO_REFER = 2;
  INADEQUATE_CONTRIBUTION_SCORE_THRESHOLD = 2.0;

  // Score Labels and Values
  SCORE_ADEQUATE_LABEL = 'Adequate';
  SCORE_INADEQUATE_LABEL = 'Inadequate';
  SCORE_ADEQUATE_VALUE = 3.0;
  SCORE_INADEQUATE_VALUE = 1.0;

  /** Reinitialize color variables to ensure consistency in color-based logic.
   * The color hex string must be in lowercase!
   */
  COLOR_MISSED_SUBMISSION = '#ead1dc';
  COLOR_MISSED_EVALUATION = '#d9d9d9';
  COLOR_EXPELLED = '#d36a6a';
  COLOR_MISSED_SUBM_AND_EVAL = '#ea9999';
  COLOR_LATE_EVALUATION = '#fff2cc';
}
