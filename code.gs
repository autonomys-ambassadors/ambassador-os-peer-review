// Hard code the references to the forms and sheets
const evaluationForm = "https://forms.gle/i89pkwmJeMcCWGvo7";

const RegistrySheetName = "Registry";
const RegistrySpreadsheet =
	"https://docs.google.com/spreadsheets/d/1FSTQKb9_GWQ7HxuwKlwrv6yRf68yjaqeE8zrEbL8btU/edit#gid=146718602";
const RegistryEmailColumn = 0;
const RegistryDiscordHandleColumn = 1;

const submissionForm = "https://forms.gle/74kT61GGWHfAjfT6A";

const SubmissionsSheetName = "Responses";
const SubmissionsSpreadsheet =
	"https://docs.google.com/spreadsheets/d/1FSTQKb9_GWQ7HxuwKlwrv6yRf68yjaqeE8zrEbL8btU/edit#gid=146718602";
const SubmissionEmailAddressColumn = 1;
const SubmissionDiscordHandleColumn = 2;
const SubmissionConstributionColumn = 3;
const SubmissionLinksColumn = 4;

function addAmbassadorPeerReviewMenus() {
	var ui = SpreadsheetApp.getUi();
	ui.createMenu("Ambassador OS")
		.addItem("Request Submissions", "requestSubmissions")
		.addItem("Request Evaluations", "requestEvaluations")
		.addItem("Select Conflict Resolution Team", "selectCRT")
		.addToUi();
}

function getConfiguration() {
	var scriptProperties = PropertiesService.getScriptProperties();
	var config = {
		ambassadorRegistrySpreadsheet:
			scriptProperties.getProperty("ambassadorRegistrySpreadsheet") ||
			RegistrySpreadsheet,
		ambassadorRegistrySheet:
			scriptProperties.getProperty("ambassadorRegistrySpreadsheetName") ||
			RegistrySheetName,
		submissionsForm:
			scriptProperties.getProperty("submissionForm") || submissionForm,
		evaluationsForm:
			scriptProperties.getProperty("evaluationForm") || evaluationForm,
		submissionsSpreadsheet:
			scriptProperties.getProperty("submissionsSpreadsheet") ||
			SubmissionsSpreadsheet,
		submissionsSheet:
			scriptProperties.getProperty("submissionsSheet") || SubmissionsSheetName,
	};
	return config;
}

// 'Request Submissions' button handler
function requestSubmissions() {
	var config = getConfiguration();
	var ambassadorRegistrySpreadsheet = SpreadsheetApp.openByUrl(
		config.ambassadorRegistrySpreadsheet
	);
	var ambassadorRegistry = ambassadorRegistrySpreadsheet.getSheetByName(
		config.ambassadorRegistrySheet
	);

	SpreadsheetApp.setActiveSheet(ambassadorRegistry);
	var data = ambassadorRegistry.getDataRange().getValues();

	validateEmails(data, RegistryEmailColumn);

	// assumes that the first row is the header row starting with i=1 instaed of 0
	for (var i = 1; i < data.length; i++) {
		var email = data[i][RegistryEmailColumn];
		var discordHandle = data[i][RegistryDiscordHandleColumn];
		var subject = "Request for Submissions";
		var body =
			"Dear " +
			discordHandle +
			",\n\n" +
			"Below, find a link to the form for submitting your deliverables for " +
			getDeliverableMonth() +
			", " +
			getDeliverableYear() +
			":\n\n" +
			config.submissionsForm +
			"\n\n" +
			"Please note that the submission deadline falls on " +
			getExpectedResponseDate() +
			".\n\n" +
			"Thank You,\n\n" +
			"Fradique";
		MailApp.sendEmail(email, subject, body);
		console.log("RFS Email sent to " + email + " for " + discordHandle);
	}
}

// 'Request Evaluations' button handler
function requestEvaluations() {
	var config = getConfiguration();
	// array to keep track of how many times each ambassador has been selected
	var ambassadorCount = {};
	var ambassadorRegistrySpreadsheet = SpreadsheetApp.openByUrl(
		config.ambassadorRegistrySpreadsheet
	);
	var ambassadorRegistry = ambassadorRegistrySpreadsheet.getSheetByName(
		config.ambassadorRegistrySheet
	);
	var leftColumn = 0;
	var rightColumn = 0;
	var tempRegistryEmailColumn = 0;
	if (RegistryEmailColumn < RegistryDiscordHandleColumn) {
		leftColumn = RegistryEmailColumn + 1;
		tempRegistryEmailColumn = 1;
		rightColumn = RegistryDiscordHandleColumn + 1;
	} else {
		leftColumn = RegistryDiscordHandleColumn + 1;
		rightColumn = RegistryEmailColumn + 1;
		tempRegistryEmailColumn =
			RegistryEmailColumn - RegistryDiscordHandleColumn + 1;
	}
	var ambassadors = ambassadorRegistry
		//assumes first row is header (get range is 1-indexed not 0-indexed)
		.getRange(2, leftColumn, ambassadorRegistry.getLastRow() - 1, rightColumn)
		.getValues();

	// Triple the list of ambassadors to select from
	var potentialEvaluators = [...ambassadors, ...ambassadors, ...ambassadors];

	var submissionSpreadsheet = SpreadsheetApp.openByUrl(
		config.submissionsSpreadsheet
	);
	var submissionSheet = submissionSpreadsheet.getSheetByName(
		config.submissionsSheet
	);
	var submissions = submissionSheet.getDataRange().getValues();

	// Create a new sheet for logging the reviewers
	var reviewLogSheet = submissionSpreadsheet.getSheetByName("Review Log");
	if (!reviewLogSheet) {
		reviewLogSheet = submissionSpreadsheet.insertSheet("Review Log");
		reviewLogSheet.appendRow([
			"Submitter",
			"Reviewer 1",
			"Reviewer 2",
			"Reviewer 3",
		]);
	}
	// temp registry email is 1-indexed, array is 0-indexed
	validateEmails(ambassadors, tempRegistryEmailColumn - 1);

	// start at index 1 to skip the header row
	for (var i = 1; i < submissions.length; i++) {
		var discordHandle = submissions[i][SubmissionDiscordHandleColumn];
		var submitterEmail = submissions[i][SubmissionEmailAddressColumn];
		var contribution = submissions[i][SubmissionConstributionColumn];
		var links = submissions[i][SubmissionLinksColumn];

		var subject = "Request for Evaluation";
		var body =
			"Dear Ambassador,\n\n" +
			"When possible, please take a moment to examine the deliverables presented for the month of " +
			getDeliverableMonth() +
			" by:\n\n" +
			discordHandle +
			"\n\n" +
			contribution +
			"\n\n" +
			links +
			"\n\n" +
			"Please assign a grade on a scale of 1 to 5 in the form linked below.\n\n" +
			config.evaluationsForm;

		// Select three unique ambassadors from each list
		var selectedAmbassadors = selectUniqueAmbassadors(
			potentialEvaluators,
			tempRegistryEmailColumn - 1,
			submitterEmail,
			ambassadorCount
		);

		// Send the email to each selected ambassador
		selectedAmbassadors.forEach((ambassador) => {
			var email = ambassador[0]; // Assuming emails are in the first column
			MailApp.sendEmail(email, subject, body);
			console.log("RFE Email sent to " + email + " for " + discordHandle);
		});

		// Log the reviewers in the new sheet
		var logRow = [discordHandle];
		selectedAmbassadors.forEach((ambassador) => {
			logRow.push(ambassador[1]); // Assuming discord handles are in the second column
		});
		reviewLogSheet.appendRow(logRow);
	}
}

// 'Select Conflict Resolution Team' button handler
function selectCRT() {
	var config = getConfiguration();
	var ambassadorRegistrySpreadsheet = SpreadsheetApp.openByUrl(
		config.ambassadorRegistrySpreadsheet
	);
	var ambassadorRegistry = ambassadorRegistrySpreadsheet.getSheetByName(
		config.ambassadorRegistrySheet
	);
	var ambassadors = ambassadorRegistry.getDataRange().getValues();
	ambassadors.sort();

	var CRTRegistry = ambassadorRegistrySpreadsheet.getSheetByName(
		"Conflict Resolution Team"
	);
	if (!CRTRegistry) {
		CRTRegistry = ambassadorRegistrySpreadsheet.insertSheet(
			"Conflict Resolution Team"
		);
		CRTRegistry.appendRow([
			"Selection Date",
			"Ambassador 1",
			"Ambassador 2",
			"Ambassador 3",
			"Ambassador 4",
			"Ambassador 5",
		]);
	}
	var lastRow = CRTRegistry.getLastRow();
	if (lastRow > 1) {
		if (lastRow < 4) {
			var lastRows = CRTRegistry.getRange(2, 2, lastRow - 1, 5).getValues();
		} else {
			var lastRows = CRTRegistry.getRange(lastRow - 3, 2, 4, 5).getValues();
		}
	}

	var recentAmbassadors = [].concat.apply([], lastRows);
	ambassadors = ambassadors.filter(function (ambassador) {
		return !recentAmbassadors.includes(ambassador[RegistryEmailColumn]);
	});

	if (ambassadors.length < 5) {
		SpreadsheetApp.getUi().alert(
			"There are fewer than 5 ambassadors to choose from - selection is failed."
		);
		return;
	}
	var selectedAmbassadors = [];
	for (var i = 0; i < 5; i++) {
		var randomIndex = Math.floor(Math.random() * ambassadors.length);
		selectedAmbassadors.push(ambassadors[randomIndex][RegistryEmailColumn]);
		sendCRTElectionEmail(
			ambassadors[randomIndex][RegistryEmailColumn],
			ambassadors[randomIndex][RegistryDiscordHandleColumn]
		);
		ambassadors.splice(randomIndex, 1);
	}

	var currentTime = new Date();
	selectedAmbassadors.unshift(currentTime);
	CRTRegistry.appendRow(selectedAmbassadors);
}

function sendCRTElectionEmail(email, discordHandle) {
	var subject = "Subspace Ambassadors Conflict Resolution Team selection";
	var body =
		"Dear " +
		discordHandle +
		",\n\n" +
		"You have been selected to serve on the Conflict Resolution Team for 3 months.\n\n" +
		"Please review the How To Participate guide here: https://coda.io/d/_dORJu5J0YW4/Participate-on-Conflict-Resolution-Team_suVvl .\n\n" +
		"If there is some reason you think you must be excused from this service, please notify the Sponsor or the Governance Team as soon as possible.\n\n" +
		"Thank You,\n\n" +
		"Fradique";
	MailApp.sendEmail(email, subject, body);
	console.log("CRT Email sent to " + email + " for " + discordHandle);
	return;
}

// Function to select three unique ambassadors from each list
function selectUniqueAmbassadors(
	potentialEvaluators,
	emailColumn,
	revieweeEmail,
	ambassadorCount
) {
	var selectedAmbassadors = [];

	for (let i = 0; i < 3; i++) {
		// Filter out the discord handle and any ambassador that has been selected more than 3 times
		// or has already been selected for this submission
		var availableAmbassadors = potentialEvaluators.filter(
			(ambassador) =>
				ambassador[emailColumn] !== revieweeEmail &&
				(ambassadorCount[ambassador[emailColumn]] || 0) < 3 &&
				!selectedAmbassadors.includes(ambassador)
		);

		if (availableAmbassadors.length === 0) {
			console.error("Could not find an available ambassador");
			return selectedAmbassadors;
		}

		var randomIndex = Math.floor(Math.random() * availableAmbassadors.length);
		var selectedAmbassador = availableAmbassadors[randomIndex];

		// Update the count for the selected ambassador
		ambassadorCount[selectedAmbassador[emailColumn]] =
			(ambassadorCount[selectedAmbassador[emailColumn]] || 0) + 1;

		// Add the selected ambassador to the list of selected ambassadors for this submission
		selectedAmbassadors.push(selectedAmbassador);

		// Remove the selected ambassador from the list of all ambassadors
		availableAmbassadors = availableAmbassadors.filter(
			(ambassador) => ambassador !== selectedAmbassador
		);
	}

	return selectedAmbassadors;
}

// Add the Ambassador OS menu to the spreadsheet
function onOpen() {
	addAmbassadorPeerReviewMenus();
}

// Function to validate all email addresses in an array of arrays, default to assume email is first column (0-indexed)
function validateEmails(listOfEmails, emailColumn = 0) {
	for (var i = 1; i < listOfEmails.length; i++) {
		var email = listOfEmails[i][emailColumn];
		if (!isValidEmail(email)) {
			throw new Error("Invalid email: " + email);
		}
	}
}

// Function to validate email addresses
function isValidEmail(email) {
	var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
	return emailRegex.test(email);
}

//date functions to get the deliverable date month and year, and a deadline for responses.
// assume that deliverable date is always last month, and that deadline is alwasy +7 days from when the requests are sent.
function getDeliverableMonth() {
	// Get the previous month
	var date = new Date();
	var monthNames = [
		"January",
		"February",
		"March",
		"April",
		"May",
		"June",
		"July",
		"August",
		"September",
		"October",
		"November",
		"December",
	];
	return (previousMonth = monthNames[(date.getMonth() - 1 + 12) % 12]);
}

function getDeliverableYear() {
	var date = new Date();
	var year = date.getFullYear();
	if (date.getMonth() === 0) {
		year--;
	}
	return year;
}

function getExpectedResponseDate() {
	var date = new Date();
	date.setDate(date.getDate() + 7);
	return Utilities.formatDate(
		date,
		Session.getScriptTimeZone(),
		"MMMM dd, yyyy"
	);
}
