# Ambassador OS Peer Review

This app script code (6 filesis) to be added to a google sheet to run the Ambassador OS.

## Some assumptions:

1. You have spreadsheet "Ambassador Registry" with sheets: Review Log, Registry, Conflict Resolution Team.
	"Registry" is a list of all ambassadors (columns: "Ambassador Email Address" , "Ambassador Discord Handle").
2. You have spreadsheet "Ambassadors' Scores" with sheets: Overall score, month-sheets.
	The "Ambassadors' Discord Handles" column of the "Overall score" sheet is a dynamically linked column that gets data from the Registry, sorted alphabetically by the "Ambassador Discord Handle" column. (In other words there should be the same list of ambassadors as in Registry).


## To install the Ambassador OS Peer Review process script:

In a google sheet, On the Extensions menu, choose Apps Script.

in the Google Script Editor manu create files and paste the content of this repository files in to them correspondingly:
 SharedUtilities 
 Module1
 Module2
 Module3
 Module4
 Module5
Google will automatically create a .gs extension for them.


  ## Recommendations to installing and using the script.

Backup Ambassadors' Scores spreadsheet.
Rename sheet 'Overall score ' to 'Overall score' (remove space at the end).
To work ideally "Ambassadors' Discord Handles" column in Overall score sheet should be an exact copy of "Ambassador Discord Handle" column in Registry sheet. Note: If you need, it could be done programmatically in nearest update.
Delete "Sheet 1" sheet in Ambassadors' Scores sprdsht, if you don't rly need it, to ease calculations.
Ensure JS version compatibility:
    Go Settings->"edit appsscript.json" tick = ON,
    Open appsscript.json 
	it should look similar to:
	{
	  "timeZone": "America/Los_Angeles",
	  "exceptionLogging": "STACKDRIVER",
	  "runtimeVersion": "V8",
	  "oauthScopes": [
	    "https://www.googleapis.com/auth/spreadsheets",
	    "https://www.googleapis.com/auth/script.send_mail",
	    "https://www.googleapis.com/auth/forms",
	    "https://www.googleapis.com/auth/script.external_request",
	    "https://www.googleapis.com/auth/script.scriptapp"
	  ]
	}


For installing the script in your environment you have to replace all letted variables in SharedUtilities.gs Global Variables section (and in the following refreshGlobalVariables function, to prevent Google cache problems. May be more simple way will be found in next update).
For production mode ensure const testing = false.
Through all the code the setMinutes and getMinutes are used. Edit Triggers and Delays section, using minutes. For ex. 7 days is 10080 minutes , possible can use format like: 60*24*7.
Columns "Penalty Points" and "Max 6-Month PP" will be added automatically if don't exist.
Every month current reporting month column will be added automatically.
If everything is working as expected - no any manually editions will be needed.
Expelled ambassadors are not deleted from Registry - their e-mail address string is concatenated with '(EXPELLED)' string.
Expelled ambassadors should be removed from Registry though, but not necessary, they anyway will not be added to next cycles.
⚠️ Use "Processing past months" option only if you want to count all "didn't submit" and "later submissoin" events in calculating total penalty points based on past months violations (already included in automatic batch).
If "too many triggers" error happens, use "Delete existing triggers" menu item to resolve this (already implemented in code). Still the option can be helpful is some cases.
Do not allow multi selecting options in Evaluation Form. Limit submitting to only one time. Editing in fact creates two forms (can lead to errors too).
Recomendation to ensure formula for "Average Score" column covers the entire range of column-months from the first to the last, adapting as new columns are added "Cummulative Score". Can be done programmatically, if needed.

## TESTING NOTES:
The script has huge size preliminary due to excessive logging and comments. It should be reduced and only important steps logging left as we progressing in testing.

The full description of all functions is in the functions_descr file.

There are 2 objects in the script properties evaluationStartTime and evaluationWindowStart. This will be optimized in the next update. However, it is not critical for operation.


## CHANGES - pls refer to "Changes" file