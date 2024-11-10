// Function to select CRT members
function selectCRTMembers() {
	const registrySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Registry');
	const crtSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Conflict Resolution Team');
  
	// Fetch all ambassador emails from Registry
	const registryData = registrySheet.getRange('A:A').getValues().flat().filter(Boolean);
	
	// Exclude expelled ambassadors and recent CRT members
	const recentCRTMembers = getRecentCRTMembers(crtSheet); // Helper function to get the last 2 months of CRT members
	const eligibleAmbassadors = registryData.filter(email => 
	  email && !email.toUpperCase().startsWith('(EXPELLED)') && !recentCRTMembers.includes(email)
	);
  
	if (eligibleAmbassadors.length < 5) {
	  Logger.log('Not enough eligible ambassadors to form the CRT.');
	  if (!testing) {
		SpreadsheetApp.getUi().alert('Failed to select CRT: not enough eligible ambassadors.');
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
	  selectedAmbassadors.forEach(ambassador => {
		sendCRTNotification(ambassador, CRT_SELECTING_NOTIFICATION_TEMPLATE); // Helper function for sending emails
	  });
  
	  Logger.log('CRT members notified.');
	} else {
	  Logger.log('Test mode: no changes made to the spreadsheet or emails sent.');
	}
  }
  
  // Helper function to get CRT members from the past 2 months
  function getRecentCRTMembers(crtSheet) {
	const today = new Date();
	const twoMonthsAgo = new Date(today.setMonth(today.getMonth() - 2));
	const data = crtSheet.getDataRange().getValues();
	
	const recentMembers = [];
	data.forEach(row => {
	  const date = row[0]; // Assuming the date is in the first column
	  if (date instanceof Date && date >= twoMonthsAgo) {
		recentMembers.push(...row.slice(1)); // Add CRT members from the row
	  }
	});
	
	return recentMembers;
  }
  
  // Helper function to select random members
  function getRandomSelection(array, num) {
	const selected = [];
	while (selected.length < num && array.length > 0) {
	  const randomIndex = Math.floor(Math.random() * array.length);
	  selected.push(array.splice(randomIndex, 1)[0]); // Remove and select random element
	}
	return selected;
  }
  
  // Helper function to send CRT notification email
  function sendCRTNotification(email, template) {
	MailApp.sendEmail(email, 'CRT Selection Notification', template);
  }
  