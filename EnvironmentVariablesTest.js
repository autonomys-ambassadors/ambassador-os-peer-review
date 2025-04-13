function setTestVariables() {
  // controls wether email will be sent or just logged for troubleshooting - should always be true for production.
  const TESTER = 'Abdala'; // 'Wilyam' or 'Jonathan'

  if (TESTER === 'Wilyam') {
    SEND_EMAIL = true;

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
    GRADE_SUBMITTER_COLUMN = 'Submitter';
    GRADE_FINAL_SCORE_COLUMN = 'Final Score';
    CRT_SELECTION_DATE_COLUMNT = 'Selection Date';

    // Triggers and Delays for testing - use much shorter windows for accelerated testing schedules
    SUBMISSION_WINDOW_MINUTES = 15;
    SUBMISSION_WINDOW_REMINDER_MINUTES = 5; // how many minutes after Submission Requests sent to remind
    EVALUATION_WINDOW_MINUTES = 10;
    EVALUATION_WINDOW_REMINDER_MINUTES = 5; // how many minutes after Evaluation Requests sent to remind

    // Penalty Points threshold - if > or = this number for the past 6 months, ambassador will be expelled
    MAX_PENALTY_POINTS_TO_EXPEL = 3;

    /** Reinitialize color variables to ensure consistency in color-based logic.
     * The color hex string must be in lowercase!
     */
    COLOR_MISSED_SUBMISSION = '#ead1dc';
    COLOR_MISSED_EVALUATION = '#d9d9d9';
    COLOR_EXPELLED = '#d36a6a';
    COLOR_MISSED_SUBM_AND_EVAL = '#ea9999';
    //COLOR_OLD_MISSED_SUBMISSION = '#ead1dc';

    // Jonathan Test vars
    //
    //
  } else if (TESTER === 'Jonathan') {
    SEND_EMAIL = true;

    // Specify your testing sheets/forms/etc. here:
    // Spreadsheets:
    AMBASSADOR_REGISTRY_SPREADSHEET_ID = '14aHm1EiK48RoclGYydI7OzcFcExYUr9tdQ6PlhjCbAg'; //"Ambassador Registry"
    AMBASSADORS_SCORES_SPREADSHEET_ID = '1lVUaCGCCbfD3l9e8MEfQVBKaljm7A5aKX7RJAsUrWfA'; // "Ambassadors' Scores"
    AMBASSADORS_SUBMISSIONS_SPREADSHEET_ID = '1lVUaCGCCbfD3l9e8MEfQVBKaljm7A5aKX7RJAsUrWfA'; // "Ambassador Submission Responses"
    EVALUATION_RESPONSES_SPREADSHEET_ID = '1lVUaCGCCbfD3l9e8MEfQVBKaljm7A5aKX7RJAsUrWfA'; // "Ambassador Evaluations' Responses"

    // Google Forms
    SUBMISSION_FORM_ID = '13oDRgD2qjryfhv992ZS99zCTOHPXBxsqKAXijupHbfE'; // ID for Submission form
    EVALUATION_FORM_ID = '1EPrKCrg7NXfEje3Ps3aBZ3S_qs9oy3dZl6SbnF3Ek4U'; // ID for Evaluation form
    SUBMISSION_FORM_URL = 'https://forms.gle/FeZzfoD6pw1fJi9D8'; // Submission Form URL for mailing
    EVALUATION_FORM_URL = 'https://forms.gle/rFiw4gxcYE25gPNs5'; // Evaluation Form URL for mailing
    FORM_RESPONSES_SHEET_NAME = 'Form Responses 3'; // Explicit name for 'Form Responses' sheet
    EVAL_FORM_RESPONSES_SHEET_NAME = 'Form Responses 4'; // Evaluation Form responses sheet

    // Sponsor Email (for notifications when ambassadors are expelled)
    SPONSOR_EMAIL = 'auto-sponsor@jkw.fm'; // Sponsor's email

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
    SUBM_FORM_USER_PROVIDED_EMAIL_COLUMN = 'Email';
    EVAL_FORM_USER_PROVIDED_EMAIL_COLUMN = 'Email';
    GOOGLE_FORM_REAL_EMAIL_COLUMN = 'Email Address';
    GOOGLE_FORM_CONTRIBUTION_DETAILS_COLUMN = `Dear Ambassador,
Please add text to your contributions during the month`;
    GOOGLE_FORM_CONTRIBUTION_LINKS_COLUMN = `Dear Ambassador,
Please add links your contributions during the month`;
    GOOGLE_FORM_EVALUATION_HANDLE_COLUMN = 'Discord handle of the ambassador you are evaluating?';
    GOOGLE_FORM_EVALUATION_GRADE_COLUMN = 'Please assign a grade on a scale of 0 to 5';
    GOOGLE_FORM_EVALUATION_REMARKS_COLUMN = 'Remarks (optional)';
    SCORE_PENALTY_POINTS_COLUMN = 'Penalty Points';
    SCORE_AVERAGE_SCORE_COLUMN = 'Average Score';
    SCORE_MAX_6M_PP_COLUMN = 'Max 6-Month PP';
    GRADE_SUBMITTER_COLUMN = 'Submitter';
    GRADE_FINAL_SCORE_COLUMN = 'Final Score';
    CRT_SELECTION_DATE_COLUMNT = 'Selection Date';

    // Triggers and Delays for testing - use much shorter windows for accelerated testing schedules
    SUBMISSION_WINDOW_MINUTES = 20;
    SUBMISSION_WINDOW_REMINDER_MINUTES = 15; // how many minutes after Submission Requests sent to remind
    EVALUATION_WINDOW_MINUTES = 20;
    EVALUATION_WINDOW_REMINDER_MINUTES = 15; // how many minutes after Evaluation Requests sent to remind

    // Penalty Points threshold - if > or = this number for the past 6 months, ambassador will be expelled
    MAX_PENALTY_POINTS_TO_EXPEL = 3;

    /** Reinitialize color variables to ensure consistency in color-based logic.
     * The color hex string must be in lowercase!
     */
    COLOR_MISSED_SUBMISSION = '#ead1dc';
    COLOR_MISSED_EVALUATION = '#d9d9d9';
    COLOR_EXPELLED = '#d36a6a';
    COLOR_MISSED_SUBM_AND_EVAL = '#ea9999';
    //COLOR_OLD_MISSED_SUBMISSION = '#ead1dc';
  } else if (TESTER === 'Abdala') {
    SEND_EMAIL = true;

    AMBASSADOR_REGISTRY_SPREADSHEET_ID = '1eWIAwZFA9EuHXZHDZ5yfydhJhldCFSpLl0-XaPGdQ8Q'; //"Ambassador Registry"
    AMBASSADORS_SCORES_SPREADSHEET_ID = '19uTQyb2bSSRaywx-rA1C2OOml3yPWGejAUkUtWUMlxo'; // "Ambassadors' Scores"
    AMBASSADORS_SUBMISSIONS_SPREADSHEET_ID = '1eWIAwZFA9EuHXZHDZ5yfydhJhldCFSpLl0-XaPGdQ8Q'; // "Ambassador Submission Responses"
    EVALUATION_RESPONSES_SPREADSHEET_ID = '1eWIAwZFA9EuHXZHDZ5yfydhJhldCFSpLl0-XaPGdQ8Q'; // "Ambassador Evaluations' Responses"
    FORM_RESPONSES_SHEET_NAME = 'Form Responses 15'; // Explicit name for 'Form Responses' sheet
    EVAL_FORM_RESPONSES_SHEET_NAME = 'Form Responses 16'; // Evaluation Form responses sheet

    // Google Forms
    //
    SUBMISSION_FORM_ID = '1uLUhlFQON_BKswX9nJtUMP--ItRlu8RjEITZ2k5BI_I'; // ID for Submission form
    EVALUATION_FORM_ID = '1Xk_boQI-4exXSUJ9kpMvyy1ViTEJ63hTFqXFYCSgCkw'; // ID for Evaluation form
    SUBMISSION_FORM_URL = 'https://forms.gle/Ro93fYN2TK78yHHD9'; // Submission Form URL for mailing
    EVALUATION_FORM_URL = 'https://forms.gle/HzQ1ZmkJYqDy8wWf7'; // Evaluation Form URL for mailing

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
    GRADE_SUBMITTER_COLUMN = 'Submitter';
    GRADE_FINAL_SCORE_COLUMN = 'Final Score';
    CRT_SELECTION_DATE_COLUMNT = 'Selection Date';

    // Triggers and Delays for testing - use much shorter windows for accelerated testing schedules
    SUBMISSION_WINDOW_MINUTES = 15;
    SUBMISSION_WINDOW_REMINDER_MINUTES = 5; // how many minutes after Submission Requests sent to remind
    EVALUATION_WINDOW_MINUTES = 10;
    EVALUATION_WINDOW_REMINDER_MINUTES = 5; // how many minutes after Evaluation Requests sent to remind

    // Penalty Points threshold - if > or = this number for the past 6 months, ambassador will be expelled
    MAX_PENALTY_POINTS_TO_EXPEL = 3;

    /** Reinitialize color variables to ensure consistency in color-based logic.
     * The color hex string must be in lowercase!
     */
    COLOR_MISSED_SUBMISSION = '#ead1dc';
    COLOR_MISSED_EVALUATION = '#d9d9d9';
    COLOR_EXPELLED = '#d36a6a';
    COLOR_MISSED_SUBM_AND_EVAL = '#ea9999';
    //COLOR_OLD_MISSED_SUBMISSION = '#ead1dc';

  }   else {
    throw new Error('Invalid tester - configure your sheets and scenarios');
  }
  

}
