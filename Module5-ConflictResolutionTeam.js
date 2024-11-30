/**
 * Retrieves CRT members from the last 2 months in the Conflict Resolution Team sheet.
 * This function ensures that ambassadors who served on the CRT within the past 2 months
 * are excluded from the new selection, as CRT members rotate every 2 months to allow fair participation.
 *
 * @param {Sheet} crtSheet - The Conflict Resolution Team sheet.
 * @returns {Array} - List of recent CRT members.
 */
function selectCRTMembers() {
  Logger.log('Starting CRT member selection process.');

  // Access Registry and CRT sheets
  const registrySheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(REGISTRY_SHEET_NAME);
  const crtSheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName('Conflict Resolution Team');

  if (!registrySheet) {
    Logger.log('Error: Registry sheet not found.');
    return;
  }

  if (!crtSheet) {
    Logger.log('Error: CRT sheet not found.');
    return;
  }

  // Fetch all data from Registry
  const registryData = registrySheet.getRange(2, 1, registrySheet.getLastRow() - 1, registrySheet.getLastColumn()).getValues();

  // Fetch recent CRT members
  const recentCRTMembers = getRecentCRTMembers(crtSheet); // Helper function to get the last 2 months of CRT members
  Logger.log(`Recent CRT members: ${JSON.stringify(recentCRTMembers)}`);

  // Filter eligible ambassadors
  const eligibleAmbassadors = registryData
    .filter(row => !row[2]?.includes('Expelled')) // Exclude expelled ambassadors
    .map(row => row[0]?.trim().toLowerCase()) // Extract valid emails
    .filter(email => email && !recentCRTMembers.includes(email)); // Exclude empty emails and recent CRT members

  Logger.log(`Eligible ambassadors emails: ${JSON.stringify(eligibleAmbassadors)}`);

  if (eligibleAmbassadors.length < 5) {
    alertAndLog('Failed to select CRT: not enough eligible ambassadors.');
    }
    return;
  }

  // Select 5 random ambassadors
  const selectedAmbassadors = getRandomSelection(eligibleAmbassadors, 5);

  Logger.log(`Selected CRT Members: ${selectedAmbassadors.join(', ')}`);

  if (!testing) {
    // Log selected ambassadors and date in CRT sheet
    const selectionDate = new Date();
    crtSheet.appendRow([selectionDate, ...selectedAmbassadors]);

    // Notify selected ambassadors via email
    selectedAmbassadors.forEach((ambassador) => {
      sendCRTNotification(ambassador, CRT_SELECTING_NOTIFICATION_TEMPLATE); // Helper function for sending emails
      Logger.log(`Notification sent to CRT member: ${ambassador}`);
    });

    Logger.log('CRT members notified.');
  } else {
    Logger.log('Test mode: no changes made to the spreadsheet or emails sent.');
  }
}


/**
 * Retrieves CRT members from the last 2 months in the Conflict Resolution Team sheet.
 * @param {Sheet} crtSheet - The Conflict Resolution Team sheet.
 * @returns {Array} - List of recent CRT members.
 */
function getRecentCRTMembers(crtSheet) {
  Logger.log('Fetching recent CRT members.');

  const today = new Date();
  const twoMonthsAgo = new Date(today.setMonth(today.getMonth() - 2));
  const data = crtSheet.getDataRange().getValues();

  const recentMembers = [];
  data.forEach((row) => {
    const date = row[0]; // Assuming the date is in the first column
    if (date instanceof Date && date >= twoMonthsAgo) {
      recentMembers.push(...row.slice(1)); // Add CRT members from the row
    }
  });

  Logger.log(`Recent CRT members: ${JSON.stringify(recentMembers)}`);
  return recentMembers;
}

/**
 * Selects a random subset of ambassadors.
 * @param {Array} array - Array of ambassadors.
 * @param {number} num - Number of ambassadors to select.
 * @returns {Array} - Randomly selected ambassadors.
 */
function getRandomSelection(array, num) {
  const selected = [];
  while (selected.length < num && array.length > 0) {
    const randomIndex = Math.floor(Math.random() * array.length);
    selected.push(array.splice(randomIndex, 1)[0]); // Remove and select random element
  }
  return selected;
}


/**
 * Sends a notification email to a selected CRT member.
 * @param {string} email - Email of the CRT member.
 * @param {string} template - Email template.
 */
function sendCRTNotification(email, template) {
  if (!email) {
    Logger.log('Skipping notification: no email provided.');
    return;
  }

  MailApp.sendEmail(email, 'CRT Selection Notification', template);
  Logger.log(`Notification sent to CRT member: ${email}`);
}
