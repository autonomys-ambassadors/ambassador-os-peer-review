# Ambassador OS Peer Review

This app script code can be added to a google sheet to run the Ambassador OS.

## Some assumptions:

1. You have a google sheet with a list of ambassadors with two columns, Ambassador Email Address, and Ambassador Discord Handle.
2. You have a google sheet with a list of ambassador contributions with 4 columns: Timestamp, Email Address, Your Discord Handle, "Dear Ambassador,
   Please add text, inputs or links to your contributions during the month of February, 2024:"

Note that column headings are not verified, but **column order is assumed**.

## To install the Ambassador OS Peer Review process script:

### Clone the project

Git clone this project `git clone https://github.com/autonomys-ambassadors/ambassador-os-peer-review.git`

### Install Node

If you are not running node, you'll need to install it. You can google, or use nvm:

```
curl -o- https://raw.githubusercontent.com/nvm-sh/nvm/v0.39.5/install.sh | bash
nvm install node
```

### Install clasp and create a project

Clasp is a google project that can push and pull scripts to sheets.
https://github.com/google/clasp

```
npm install -g @google/clasp
clasp login
clasp create --type sheets --title "Autonomys Ambassador OS Peer Review" --rootDir "./"
```

You should see output similar to this, with references to the new sheet:

```
Created new Google Sheet: https://drive.google.com/open?id=1eNXOCChZ7nvETE4mvBwi9nYairnP6DJwu50VkB-ppB4
Created new Google Sheets Add-on script: https://script.google.com/d/16Sp1Jg9hKZWi7bwuQKtRKLlWibiyqOCptNomLvAzK93ngHm1dT3fzD4t/edit
```

### Push Ambassador OS Peer Review to your new sheet

`clasp push`

Open your sheet to verify it is setup correctly.
Your sheet should be named "Autonomys Ambassador OS Peer Review" and you should have a menu labeled "Ambassador Program" on the menu bar.
You can verify the code is present by choosing Apps Script from the Extensions menu to enter the apps script editor.

### Clone the prototype sheets and update your sheet ids:

Make a copy of the google sheets for your own testing - e.g. clone one of these pre-existing sheets:

```
// Wilyam test sheets
// const AMBASSADOR_REGISTRY_SPREADSHEET_ID = '1AIMD61YKfk-JyP6Aia3UWke9pW15bPYTvWc1C46ofkU';  //"Ambassador Registry"
// const AMBASSADORS_SCORES_SPREADSHEET_ID = '1RJzCo1FgGWkx0UCYhaY3SIOR0sl7uD6vCUFw55BX0iQ';   // "Ambassadors' Scores"

// Jonathan test sheets
//const AMBASSADOR_REGISTRY_SPREADSHEET_ID = "15J5-F2_FxNJf6X2P7umiwOxJN9FckJYjIzDp3ydtZf8"; //"Ambassador Registry"
//const AMBASSADORS_SCORES_SPREADSHEET_ID = "1p6SUyoinRl9DtQ5ESQZz-wb5PpdNL6wtucrVOf20vVM"; // "Ambassadors' Scores"
```

Add your own sheets - using your cloned sheets identity - to SharedUtilities.js. Please comment them out when pushing to github, and just un-comment your relevant contants for local testing. The canonical sheets maintained by the foundation should be the const set when the code is merged to main.

```
const AMBASSADOR_REGISTRY_SPREADSHEET_ID = "MyNewRegistrySheetId";
const AMBASSADORS_SCORES_SPREADSHEET_ID = "MyNewScoreSheetId";
```

### Prepare test data

Modify the Registry and Scores actual data to suit your test cases.

## To run the process:

From the spreadsheet, you should now see a menu called "Ambassador OS" with menu items for Request Submissions and Request Evaluations.

Choose the relevant menu option to send out emails requesting evaluations or submissions.

When requesting Evaluations, the script will add a new sheet to the Submissions spreadsheet called Review Log to record which ambassadors received which evaluation request.
