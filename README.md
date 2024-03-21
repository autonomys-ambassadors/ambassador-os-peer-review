# Ambassador OS Peer Review

This app script code can be added to a google sheet to run the Ambassador OS.

## Some assumptions:

1. You have a google sheet with a list of ambassadors with two columns, Ambassador Email Address, and Ambassador Discord Handle.
2. You have a google sheet with a list of ambassador contributions with 4 columns: Timestamp, Email Address, Your Discord Handle, "Dear Ambassador,
   Please add text, inputs or links to your contributions during the month of February, 2024:"

Note that column headings are not verified, but **column order is assumed**.

## To install the Ambassador OS Peer Review process script:

In a google sheet, On the Extensions menu, choose Apps Script.

Copy and paste in the contents of the code.gs file.

Edit these values at the top of the script to refer to your spreadsheets and forms. You can get the sheet urls from url bar when browsing the relevant sheet, and the sheet names are the tab names.

```
// Hard code the references to the forms and sheets
const evaluationForm = "https://forms.gle/i89pkwmJeMcCWGvo7";

const RegistrySheetName = "Registry";
const RegistrySpreadsheet =
	"https://docs.google.com/spreadsheets/d/1FSTQKb9_GWQ7HxuwKlwrv6yRf68yjaqeE8zrEbL8btU/edit#gid=146718602";

const submissionForm = "https://forms.gle/74kT61GGWHfAjfT6A";

const SubmissionsSheetName = "Responses";
const SubmissionsSpreadsheet =
	"https://docs.google.com/spreadsheets/d/1FSTQKb9_GWQ7HxuwKlwrv6yRf68yjaqeE8zrEbL8btU/edit#gid=146718602";
```

In the header run menu, choose function addAmbassadorPeerReviewMenus and click run.

## To run the process:

From the spreadsheet, you should now see a menu called "Ambassador OS" with menu items for Request Submissions and Request Evaluations.

Choose the relevant menu option to send out emails requesting evaluations or submissions.

When requesting Evaluations, the script will add a new sheet to the Submissions spreadsheet called Review Log to record which ambassadors received which evaluation request.
