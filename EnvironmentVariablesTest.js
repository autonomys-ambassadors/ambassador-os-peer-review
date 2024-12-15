function setTestVariables() {
  // controls wether email will be sent or just logged for troubleshooting - should always be true for production.
  SEND_EMAIL = true;
  const TESTER = 'Wilyam'; // 'Wilyam' or 'Jonathan'
  if ((TESTER = 'Wilyam')) {
    // Specify your testing sheets/forms/etc. here:
    // Spreadsheets:
    //

    // Wilyam test sheets
    AMBASSADOR_REGISTRY_SPREADSHEET_ID = '1AIMD61YKfk-JyP6Aia3UWke9pW15bPYTvWc1C46ofkU'; //"Ambassador Registry"
    AMBASSADORS_SCORES_SPREADSHEET_ID = '1RJzCo1FgGWkx0UCYhaY3SIOR0sl7uD6vCUFw55BX0iQ'; // "Ambassadors' Scores"
    AMBASSADORS_SUBMISSIONS_SPREADSHEET_ID = '1AIMD61YKfk-JyP6Aia3UWke9pW15bPYTvWc1C46ofkU'; // "Ambassador Submission Responses"
    EVALUATION_RESPONSES_SPREADSHEET_ID = '1RJzCo1FgGWkx0UCYhaY3SIOR0sl7uD6vCUFw55BX0iQ'; // "Ambassador Evaluations' Responses"
    FORM_RESPONSES_SHEET_NAME = 'Form Responses 14'; // Explicit name for 'Form Responses' sheet
    EVAL_FORM_RESPONSES_SHEET_NAME = 'Form Responses 2'; // Evaluation Form responses sheet
    // Google Forms
    //
    // Wilyam test forms:
    SUBMISSION_FORM_ID = '1SV5rJbzPv6BQgDZkC_xgrauWgoKPcEmtk3aKY6f4ZC8'; // ID for Submission form
    EVALUATION_FORM_ID = '15UXnrpOOoZPO7XCP2TV7mwezewHY6UIsYAU_W_aoMwo'; // ID for Evaluation form
    SUBMISSION_FORM_URL = 'https://forms.gle/beZrwuP9Zs1HvUY49'; // Submission Form URL for mailing
    EVALUATION_FORM_URL = 'https://forms.gle/kndReXQqXT6JyKX68'; // Evaluation Form URL for mailing

    // Sponsor Email (for notifications when ambassadors are expelled)
    //
    // Wilyam test sponsor
    SPONSOR_EMAIL = 'economicsilver@starmail.net'; // Sponsor's email
  } else if ((TESTER = 'Jonathan')) {
    // Specify your testing sheets/forms/etc. here:
    // Spreadsheets:
    //
    //
    // Jonathan test sheets
    AMBASSADOR_REGISTRY_SPREADSHEET_ID = '14aHm1EiK48RoclGYydI7OzcFcExYUr9tdQ6PlhjCbAg'; //"Ambassador Registry"
    AMBASSADORS_SCORES_SPREADSHEET_ID = '1lVUaCGCCbfD3l9e8MEfQVBKaljm7A5aKX7RJAsUrWfA'; // "Ambassadors' Scores"
    AMBASSADORS_SUBMISSIONS_SPREADSHEET_ID = '1lVUaCGCCbfD3l9e8MEfQVBKaljm7A5aKX7RJAsUrWfA'; // "Ambassador Submission Responses"
    EVALUATION_RESPONSES_SPREADSHEET_ID = '1lVUaCGCCbfD3l9e8MEfQVBKaljm7A5aKX7RJAsUrWfA'; // "Ambassador Evaluations' Responses"

    // Google Forms
    //
    //
    // Jonathan test forms:
    SUBMISSION_FORM_ID = '13oDRgD2qjryfhv992ZS99zCTOHPXBxsqKAXijupHbfE'; // ID for Submission form
    EVALUATION_FORM_ID = '1EPrKCrg7NXfEje3Ps3aBZ3S_qs9oy3dZl6SbnF3Ek4U'; // ID for Evaluation form
    SUBMISSION_FORM_URL = 'https://forms.gle/FeZzfoD6pw1fJi9D8'; // Submission Form URL for mailing
    EVALUATION_FORM_URL = 'https://forms.gle/rFiw4gxcYE25gPNs5'; // Evaluation Form URL for mailing
    FORM_RESPONSES_SHEET_NAME = 'Form Responses 3'; // Explicit name for 'Form Responses' sheet
    EVAL_FORM_RESPONSES_SHEET_NAME = 'Form Responses 4'; // Evaluation Form responses sheet

    // Sponsor Email (for notifications when ambassadors are expelled)
    //
    //
    // Jonathan test sponsor
    SPONSOR_EMAIL = 'jonathan@jkw.fm'; // Sponsor's email
  } else {
    throw new Error('Invalid tester - configure your sheets and scenarios');
  }

  // Triggers and Delays for testing - use much shorter windows for accelerated testing schedules
  SUBMISSION_WINDOW_MINUTES = 10;
  SUBMISSION_WINDOW_REMINDER_MINUTES = 5; // how many minutes after Submission Requests sent to remind
  EVALUATION_WINDOW_MINUTES = 10;
  EVALUATION_WINDOW_REMINDER_MINUTES = 5; // how many minutes after Evaluation Requests sent to remind

  // Sheet names
  REGISTRY_SHEET_NAME = 'Registry';
  REVIEW_LOG_SHEET_NAME = 'Review Log';
  CONFLICT_RESOLUTION_TEAM_SHEET_NAME = 'Conflict Resolution Team';
  OVERALL_SCORE_SHEET_NAME = 'Overall score'; // Overall score sheet in Ambassadors' Scores

  // Columns
  AMBASSADOR_EMAIL_COLUMN = 'Ambassador Email Address';
  AMBASSADOR_DISCORD_HANDLE_COLUMN = 'Ambassador Discord Handle';
  AMBASSADOR_STATUS_COLUMN = 'Ambassador Status';

  // Penalty Points threshold - if > or = this number for the past 6 months, ambassador will be expelled
  MAX_PENALTY_POINTS_TO_EXPEL = 2;

  /** Reinitialize color variables to ensure consistency in color-based logic.
   * The color hex string must be in lowercase!
   */
  COLOR_MISSED_SUBMISSION = '#ead1dc';
  COLOR_MISSED_EVALUATION = '#d9d9d9';
  COLOR_EXPELLED = '#d36a6a';
  COLOR_MISSED_SUBM_AND_EVAL = '#ea9999';
  COLOR_OLD_MISSED_SUBMISSION = '#ead1dc';
}
