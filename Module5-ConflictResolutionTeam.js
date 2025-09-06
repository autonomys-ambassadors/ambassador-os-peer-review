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
  const registrySheet = getRegistrySheet();
  const crtSheet = getCRTSheet();

  if (!registrySheet) {
    alertAndLog('Error: Registry sheet not found.');
    throw new Error('Registry sheet not found.');
  }

  if (!crtSheet) {
    alertAndLog('Error: CRT sheet not found.');
    throw new Error('CRT sheet not found.');
  }

  // Fetch all data from Registry
  const registryData = registrySheet
    .getRange(2, 1, registrySheet.getLastRow() - 1, registrySheet.getLastColumn())
    .getValues();

  // Fetch recent CRT members
  const recentCRTMembers = getRecentCRTMembers(crtSheet); // Helper function to get the last 2 months of CRT members
  Logger.log(`Recent CRT members: ${JSON.stringify(recentCRTMembers)}`);

  const statusColumnIndex = getRequiredColumnIndexByName(registrySheet, AMBASSADOR_STATUS_COLUMN);
  const emailColumnIndex = getRequiredColumnIndexByName(registrySheet, AMBASSADOR_EMAIL_COLUMN);

  // Filter eligible ambassadors
  const eligibleAmbassadors = registryData
    .filter((row) => !row[statusColumnIndex - 1]?.toLowerCase().includes('expelled')) // Exclude expelled ambassadors
    .map((row) => normalizeEmail(row[emailColumnIndex - 1])) // Extract valid emails
    .filter((email) => email && !recentCRTMembers.map(normalizeEmail).includes(email)); // Exclude empty emails and recent CRT members

  Logger.log(`Eligible ambassadors emails: ${JSON.stringify(eligibleAmbassadors)}`);

  if (eligibleAmbassadors.length < 5) {
    alertAndLog('Failed to select CRT: not enough eligible ambassadors.');
    throw new Error('Not enough eligible ambassadors.');
  }

  // Select 5 random ambassadors
  const selectedAmbassadors = getRandomSelection(eligibleAmbassadors, 5);

  Logger.log(`Selected CRT Members: ${selectedAmbassadors.join(', ')}`);

  // Log selected ambassadors and date in CRT sheet
  const selectionDate = new Date();
  crtSheet.appendRow([selectionDate, ...selectedAmbassadors]);
}

/**
 * Retrieves CRT members from the last 2 months in the Conflict Resolution Team sheet.
 * @param {Sheet} crtSheet - The Conflict Resolution Team sheet.
 * @returns {Array} - List of recent CRT members.
 */
function getRecentCRTMembers(crtSheet) {
  const selectionDateIndex = getRequiredColumnIndexByName(crtSheet, CRT_SELECTION_DATE_COLUMN);
  const today = new Date();
  const twoMonthsAgo = new Date(today.setMonth(today.getMonth() - 2));
  const data = crtSheet.getDataRange().getValues();
  const recentMembers = [];

  data.forEach((row) => {
    const date = row[selectionDateIndex - 1]; // Assuming the date is in the first column
    if (date instanceof Date && date >= twoMonthsAgo) {
      recentMembers.push(...row.slice(1)); // Add CRT members from the row
    }
  });
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
