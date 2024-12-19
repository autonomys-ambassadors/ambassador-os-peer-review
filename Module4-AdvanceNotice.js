/**
 * Notifies ambassadors of an upcoming peer review, excluding those with 'Expelled' status in Registry.
 */
function notifyUpcomingPeerReview() {
  try {
    Logger.log('Starting upcoming peer review notification process.');

    // Access the Registry sheet
    const registrySheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(
      REGISTRY_SHEET_NAME
    );

    // Fetch all rows from the Registry sheet
    const registryData = registrySheet.getRange(2, 1, registrySheet.getLastRow() - 1, registrySheet.getLastColumn()).getValues();

    // Filter out ambassadors with 'Expelled' in their status
    const eligibleEmails = registryData
      .filter(row => !row[2]?.includes('Expelled')) // Exclude expelled ambassadors
      .map(row => row[0]?.trim()) // Extract valid emails
      .filter(email => email); // Exclude empty emails
    
        // Get the email template
    const upcomingPeerReviewTemplate = NOTIFY_UPCOMING_PEER_REVIEW;
    // Send notification to each eligible ambassador
    eligibleEmails.forEach(email => {
      MailApp.sendEmail({
        to: email,
        subject: 'Upcoming Peer Review Notification',
        body: upcomingPeerReviewTemplate, // Use plain text template
      });
      Logger.log(`Notification sent to: ${email}`);
    });

    Logger.log('Upcoming peer review notifications completed.');
  } catch (error) {
    Logger.log(`Error in notifyUpcomingPeerReview: ${error.message}`);
  }
}
