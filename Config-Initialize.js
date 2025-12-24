/**
 * Configuration Variable Declarations
 *
 * This file declares (but does not initialize) all configuration variables used across the application.
 * The actual values are set by the appropriate configuration function based on CONFIG_NAME in SharedUtilities.js
 *
 * Since Google Apps Script loads all .js files together, we need to declare variables in one place
 * to avoid conflicts, then let the configuration functions assign the values.
 */
var TESTING; // Boolean to indicate if the script is in testing mode

// Configuration control
var SEND_EMAIL; // Will control whether emails are sent - must be true for production; may be true or false for testing depending on testing needs.

// Spreadsheet IDs
var AMBASSADOR_REGISTRY_SPREADSHEET_ID; //"Ambassador Registry"
var AMBASSADORS_SCORES_SPREADSHEET_ID; // "Ambassadors' Scores"
var AMBASSADORS_SUBMISSIONS_SPREADSHEET_ID; // "Ambassador Submission Responses"
var EVALUATION_RESPONSES_SPREADSHEET_ID; // "Evaluation Responses"
var ANONYMOUS_SCORES_SPREADSHEET_ID; // "Anonymous Peer Review Scores" - for publishing monthly anonymous scores

// Google Forms
var SUBMISSION_FORM_ID; // ID for Submission form
var EVALUATION_FORM_ID; // ID for Evaluation form
var SUBMISSION_FORM_URL; // Submission Form URL for mailing
var EVALUATION_FORM_URL; // Evaluation Form URL for mailing

// Sheet names
var REGISTRY_SHEET_NAME;
var FORM_RESPONSES_SHEET_NAME;
var REVIEW_LOG_SHEET_NAME;
var CONFLICT_RESOLUTION_TEAM_SHEET_NAME;
var OVERALL_SCORE_SHEET_NAME; // Overall score sheet in Ambassadors' Scores
var EVAL_FORM_RESPONSES_SHEET_NAME; // Evaluation Form responses sheet
var CRT_LOG_SHEET_NAME; // CRT Log sheet for tracking conflict resolution referrals

// Ambassador Registry Columns
var AMBASSADOR_ID_COLUMN;
var AMBASSADOR_EMAIL_COLUMN;
var AMBASSADOR_DISCORD_HANDLE_COLUMN;
var AMBASSADOR_STATUS_COLUMN;
var AMBASSADOR_PRIMARY_TEAM_COLUMN;

// New Registry Sheet Columns
var REGISTRY_NOTION_ID_COLUMN;
var REGISTRY_SECONDARY_TEAM_COLUMN;
var REGISTRY_START_DATE_COLUMN;

// Google Form Columns
var GOOGLE_FORM_TIMESTAMP_COLUMN;
var GOOGLE_FORM_CONTRIBUTION_DETAILS_COLUMN;
var GOOGLE_FORM_CONTRIBUTION_LINKS_COLUMN;
var SUBM_FORM_USER_PROVIDED_EMAIL_COLUMN;
var EVAL_FORM_USER_PROVIDED_EMAIL_COLUMN;
var GOOGLE_FORM_REAL_EMAIL_COLUMN;
var GOOGLE_FORM_EVALUATION_HANDLE_COLUMN;
var GOOGLE_FORM_EVALUATION_GRADE_COLUMN;
var GOOGLE_FORM_EVALUATION_REMARKS_COLUMN;

// Score Columns
var SCORE_PENALTY_POINTS_COLUMN;
var SCORE_AVERAGE_SCORE_COLUMN;
var SCORE_MAX_6M_PP_COLUMN;
var SUBMITTER_HANDLE_COLUMN_IN_MONTHLY_SCORE;
var SUBMITTER_HANDLE_COLUMN_IN_REVIEW_LOG;
var GRADE_FINAL_SCORE_COLUMN;
var GRADE_EVAL_1_SCORE_COLUMN;
var GRADE_EVAL_1_REMARKS_COLUMN;
var GRADE_EVAL_2_SCORE_COLUMN;
var GRADE_EVAL_2_REMARKS_COLUMN;
var GRADE_EVAL_3_SCORE_COLUMN;
var GRADE_EVAL_3_REMARKS_COLUMN;
var CRT_SELECTION_DATE_COLUMN;
var SCORE_INADEQUATE_CONTRIBUTION_COLUMN;
var SCORE_CRT_REFERRAL_HISTORY_COLUMN;

// CRT Log columns
var CRT_LOG_EMAIL_COLUMN;
var CRT_LOG_DISCORD_HANDLE_COLUMN;
var CRT_LOG_REFERRAL_DATE_COLUMN;
var CRT_LOG_RESOLUTION_DATE_COLUMN;
var CRT_LOG_RESOLUTION_COLUMN;

// Request Log columns
var REQUEST_LOG_REQUEST_TYPE_COLUMN;
var REQUEST_LOG_MONTH_COLUMN;
var REQUEST_LOG_YEAR_COLUMN;
var REQUEST_LOG_START_TIME_COLUMN;
var REQUEST_LOG_END_TIME_COLUMN;

// Notion Configuration
var NOTION_DATABASE_ID;

// Notion API Field Names
var NOTION_NUMBER_COLUMN;
var NOTION_EMAIL_COLUMN;
var NOTION_DISCORD_COLUMN;
var NOTION_STATUS_COLUMN;
var NOTION_PRIMARY_TEAM_COLUMN;
var NOTION_SECONDARY_TEAM_COLUMN;
var NOTION_START_DATE_COLUMN;

// Email Configuration
var SPONSOR_EMAIL; // Sponsor's email
var TESTER_EMAIL; // Tester's email for redirecting test emails (only used in test configurations) SHOULD STAY UNDEFINED IN PRODUCTION.

// Timing Configuration
var SUBMISSION_WINDOW_MINUTES;
var SUBMISSION_WINDOW_REMINDER_MINUTES; // how many minutes after Submission Requests sent to remind
var EVALUATION_WINDOW_MINUTES;
var EVALUATION_WINDOW_REMINDER_MINUTES; // how many minutes after Evaluation Requests sent to remind

// Penalty and Threshold Configuration
var MAX_PENALTY_POINTS_TO_EXPEL; // Penalty Points threshold - if >= this number for the past 6 months, ambassador will be expelled
var MAX_INADEQUATE_CONTRIBUTION_COUNT_TO_REFER;
var INADEQUATE_CONTRIBUTION_SCORE_THRESHOLD;

// Color Configuration (The color hex string must be in lowercase!)
var COLOR_MISSED_SUBMISSION;
var COLOR_MISSED_EVALUATION;
var COLOR_EXPELLED;
var COLOR_MISSED_SUBM_AND_EVAL;
var COLOR_LATE_EVALUATION;
