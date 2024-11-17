// Function to notify ambassadors of an upcoming peer review, excluding those whose emails start with "(EXPELLED)"
function notifyUpcomingPeerReview() {
  try {
    Logger.log('Starting upcoming peer review notification process.');

    const registrySheet = SpreadsheetApp.openById(AMBASSADOR_REGISTRY_SPREADSHEET_ID).getSheetByName(
      REGISTRY_SHEET_NAME
    );
    const data = registrySheet.getRange(2, 1, registrySheet.getLastRow() - 1, 1).getValues(); // Get all emails from column A, skipping the header

    const upcomingPeerReviewTemplate = NOTIFY_UPCOMING_PEER_REVIEW; // Using the string template stored in variables

    // Loop through the emails in column A
    for (let i = 0; i < data.length; i++) {
      const email = data[i][0];

      // Skip ambassadors with emails starting with "(EXPELLED)"
      if (email && email.toUpperCase().startsWith('(EXPELLED)')) {
        Logger.log(`Skipping expelled ambassador: ${email}`);
        continue;
      }

      // Send notification to valid ambassadors
      if (email) {
        MailApp.sendEmail({
          to: email,
          subject: 'Upcoming Peer Review Notification',
          body: upcomingPeerReviewTemplate, // Use plain text template
        });
        Logger.log(`Notification sent to: ${email}`);
      } else {
        Logger.log(`Skipping row ${i + 2}: Email not found.`); // Add 2 to row to match sheet row (because of header)
      }
    }

    Logger.log('Upcoming peer review notifications completed.');
  } catch (error) {
    Logger.log(`Error in notifyUpcomingPeerReview: ${error.message}`);
  }
}
