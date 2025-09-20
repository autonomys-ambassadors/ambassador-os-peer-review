// (( Configuration System ))
// Configuration selection - set this to the name of your configuration:
// 'Production' for live environment, or any tester name like 'Jonathan', 'Wilyam', etc.
const CONFIG_NAME = 'Jonathan'; // Available: 'Production', 'Jonathan', 'Wilyam' - add more in Config-[Name].js files

// Note: All configuration variables are declared in Config-Initialize.js
// Their values are set by the configuration functions in Config-[Name].js files

// Unified configuration loader - calls the appropriate configuration function based on CONFIG_NAME
switch (CONFIG_NAME) {
  case 'Production':
    setNewProductionVariables();
    break;
  case 'Jonathan':
    setJonathanVariables();
    break;
  case 'Wilyam':
    setWilyamVariables();
    break;
  default:
    throw new Error(
      `Unknown configuration: "${CONFIG_NAME}". Available configurations: 'Production', 'Jonathan', 'Wilyam'. To add a new configuration, create a Config-[Name].js file with a set[Name]Variables() function.`
    );
}

// Log which configuration is active
Logger.log(`Configuration loaded: ${CONFIG_NAME}`);
if (typeof TESTER_EMAIL !== 'undefined') {
  Logger.log(`Tester email: ${TESTER_EMAIL}`);
}

//    On Open Menu
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  Logger.log('Initializing menu on spreadsheet open.');
  ui.createMenu('Ambassador Program')
    .addItem('Request Submissions', 'requestMonthlySubmissions') // Request Submissions
    .addItem('Request Evaluations', 'requestEvaluationsModule') // Request Evaluations
    .addItem('Compliance Audit', 'runComplianceAudit') // Process Scores and Penalties
    .addItem('Notify Upcoming Peer Review', 'notifyUpcomingPeerReview') // Peer Review notifications
    .addItem('Select CRT members', 'selectCRTMembers') // CRT
    .addItem('🔧️Batch process scores', 'batchProcessEvaluationResponses') //Re-runs score responses
    .addItem('🔧️Create/Sync Columns', 'syncRegistryColumnsToOverallScore') // creates Ambassador Status column in Overall score sheet; Syncs Ambassadors' Discord Handles and Ambassador Status columns between Registry and Overall score.
    .addItem('🔧️Check Emails in Submission Form responses', 'validateEmailsInSubmissionForm') // Checks completance of emails in 'Your Email Address' field of Submission Form. Recommended to run before Evaluation Requests to avoid errors caused by users' typo.
    .addItem('🔧️Delete Existing Triggers', 'deleteExistingTriggers') // Optional item
    .addItem('🔧️Force Authorization', 'forceAuthorization') // Authorization trigger
    .addToUi();
  Logger.log('Menu initialized.');
}
