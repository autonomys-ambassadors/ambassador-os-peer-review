// Hard code the references to the forms and sheets
const evaluationForm = "https://forms.gle/i89pkwmJeMcCWGvo7";

const RegistrySheetName = "Registry";
const RegistrySpreadsheet =
	"https://docs.google.com/spreadsheets/d/1FSTQKb9_GWQ7HxuwKlwrv6yRf68yjaqeE8zrEbL8btU/edit#gid=146718602";

const submissionForm = "https://forms.gle/74kT61GGWHfAjfT6A";

const SubmissionsSheetName = "Responses";
const SubmissionsSpreadsheet =
	"https://docs.google.com/spreadsheets/d/1FSTQKb9_GWQ7HxuwKlwrv6yRf68yjaqeE8zrEbL8btU/edit#gid=146718602";

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

	var data = ambassadorRegistry.getDataRange().getValues();

	SpreadsheetApp.setActiveSheet(ambassadorRegistry);

	validateEmails(data);

	// Assuming emails are in the first column and discord handles are in the second
	for (var i = 1; i < data.length; i++) {
		var email = data[i][0];
		var discordHandle = data[i][1];
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
	}
}

// 'Request Evaluations' button handler
function requestEvaluations() {
	var config = getConfiguration();
	var ambassadorCount = {};
	var ambassadorRegistrySpreadsheet = SpreadsheetApp.openByUrl(
		config.ambassadorRegistrySpreadsheet
	);
	var ambassadorRegistry = ambassadorRegistrySpreadsheet.getSheetByName(
		config.ambassadorRegistrySheet
	);
	var ambassadors = ambassadorRegistry
		.getRange(2, 1, ambassadorRegistry.getLastRow() - 1, 2)
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

	validateEmails(ambassadors);

	// Assuming Discord Handle and Contribution are in the third and fourth columns
	// start at index 1 to skip the header row
	for (var i = 1; i < submissions.length; i++) {
		var discordHandle = submissions[i][2];
		var contribution = submissions[i][3];
		var links = submissions[i][4];

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
			discordHandle,
			ambassadorCount
		);

		// Send the email to each selected ambassador
		selectedAmbassadors.forEach((ambassador) => {
			var email = ambassador[0]; // Assuming emails are in the first column
			MailApp.sendEmail(email, subject, body);
			console.log("Email sent to " + email + " for " + discordHandle);
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
	var ambassadors = ambassadorRegistry
		.getRange(2, 1, ambassadorRegistry.getLastRow() - 1, 2)
		.getValues();
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
		return !recentAmbassadors.includes(ambassador[0]);
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
		selectedAmbassadors.push(ambassadors[randomIndex][0]);
		sendCRTElectionEmail(
			ambassadors[randomIndex][0],
			ambassadors[randomIndex][1]
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
	return;
}

// Function to select three unique ambassadors from each list
function selectUniqueAmbassadors(
	potentialEvaluators,
	discordHandle,
	ambassadorCount
) {
	var selectedAmbassadors = [];

	for (let i = 0; i < 3; i++) {
		// Filter out the discord handle and any ambassador that has been selected more than 3 times
		// or has already been selected for this submission
		var availableAmbassadors = potentialEvaluators.filter(
			(ambassador) =>
				ambassador[1] !== discordHandle &&
				(ambassadorCount[ambassador[1]] || 0) < 3 &&
				!selectedAmbassadors.includes(ambassador)
		);

		if (availableAmbassadors.length === 0) {
			console.error("Could not find an available ambassador");
			return selectedAmbassadors;
		}

		var index = Math.floor(Math.random() * availableAmbassadors.length);
		var selectedAmbassador = availableAmbassadors[index];

		// Update the count for the selected ambassador
		ambassadorCount[selectedAmbassador[1]] =
			(ambassadorCount[selectedAmbassador[1]] || 0) + 1;

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

// Function to validate all email addresses in an array assuming email is the first column
function validateEmails(listOfEmails) {
	for (var i = 1; i < listOfEmails.length; i++) {
		var email = listOfEmails[i][0];
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
// assume that deliverable date is always last month, and that deadline is alwasy +10 days from when the requests are sent.
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
