/**
 * Notifies ambassadors of an upcoming peer review, excluding those with 'Expelled' status in Registry.
 */
function notifyUpcomingPeerReview() {
  try {
    Logger.log('Starting upcoming peer review notification process.');

    // Access the Registry sheet
    const registrySheet = getRegistrySheet();

    // Fetch all rows from the Registry sheet
    const registryData = registrySheet
      .getRange(2, 1, registrySheet.getLastRow() - 1, registrySheet.getLastColumn())
      .getValues();

    const statusColumnIndex = getRequiredColumnIndexByName(registrySheet, AMBASSADOR_STATUS_COLUMN);
    const emailColumnIndex = getRequiredColumnIndexByName(registrySheet, AMBASSADOR_EMAIL_COLUMN);

    // Filter out ambassadors with 'Expelled' in their status
    const eligibleEmails = registryData
      .filter((row) => isActiveAmbassador(row, emailColumnIndex - 1, statusColumnIndex - 1)) // Exclude expelled ambassadors
      .map((row) => normalizeEmail(row[emailColumnIndex - 1])) // Extract valid emails
      .filter((email) => email); // Exclude empty emails

    // Get the email template
    const upcomingPeerReviewTemplate = NOTIFY_UPCOMING_PEER_REVIEW;
    // Send notification to each eligible ambassador
    eligibleEmails.forEach((email) => {
      sendEmailNotification(email, 'Upcoming Peer Review Notification', upcomingPeerReviewTemplate);
    });

    Logger.log('Upcoming peer review notifications completed.');
  } catch (error) {
    Logger.log(`Error in notifyUpcomingPeerReview: ${error.message}`);
  }
}
