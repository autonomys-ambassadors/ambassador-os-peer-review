# Ambassador OS Peer Review

This app script code can be added to a google sheet to run the Ambassador OS.

## Some assumptions:

1. You have a google sheet with a list of ambassadors with two columns, Ambassador Email Address, and Ambassador Discord Handle.
2. You have a google sheet with a list of ambassador contributions with 4 columns: Timestamp, Email Address, Your Discord Handle, "Dear Ambassador,
   Please add text, inputs or links to your contributions during the month:"

Note that column headings are not verified, but **column order is assumed**.

## To install the Ambassador OS Peer Review process script:

### Clone the project

Git clone this project `git clone https://github.com/autonomys-ambassadors/ambassador-os-peer-review.git`

### Install Node

If you are not running node, you'll need to install it. You can google, or use nvm:

```
curl -o- https://raw.githubusercontent.com/nvm-sh/nvm/v0.39.5/install.sh | bash
source ~/.bashrc
nvm install node
```

### Install clasp

Clasp is a google project that can push and pull scripts to sheets.
https://github.com/google/clasp

```
npm install -g @google/clasp
clasp login
```

### If You Already Have Google Apps Script Project

Enable API in your Google Apps Interface:

https://script.google.com/home/usersettings

1. Убедитесь, что у вас есть scriptId существующего проекта

   open your project in Google Apps Script Editor:
   File > Project Settings.
   Copy Script ID.

2. Link your local directory \*where you pulled the repository to) with your Google Apps Script

cd <path-to-local-procet-dir>
clasp settings set scriptId <SCRIPT_ID>
Manually create .clasp.json:
nano .clasp.json
{
"scriptId": "<your script ID>",
"rootDir": "./"
}

3. clasp push

### if You Haven't Project yet - Create a Project

```
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

### Clone the prototype sheets and update your sheet ids (if you already have Google Apps Project):

Make a copy of the google sheets for your own testing - e.g. open the example testing Registry and Scoring sheets below and choose File/Make a copy to create your own copy of the testing sheets.

Registry:
`https://docs.google.com/spreadsheets/d/15J5-F2_FxNJf6X2P7umiwOxJN9FckJYjIzDp3ydtZf8/edit?gid=368768780#gid=368768780`

Scoring:
`https://docs.google.com/spreadsheets/d/1p6SUyoinRl9DtQ5ESQZz-wb5PpdNL6wtucrVOf20vVM/edit?usp=sharing`

Add your own sheets - using your cloned sheets' identity - to EnvironmentVariablesTest.js. Please comment them out when pushing to github, and just un-comment your relevant contants for local testing. The canonical sheets maintained by the foundation should be the const set when the code is merged to main.

```
const AMBASSADOR_REGISTRY_SPREADSHEET_ID = "MyNewRegistrySheetId";
const AMBASSADORS_SCORES_SPREADSHEET_ID = "MyNewScoreSheetId";
```

The sheet id is the string after /spreadsheets/d - for example, the bolded portion here: https://docs.google.com/spreadsheets/d/**1p6SUyoinRl9DtQ5ESQZz-wb5PpdNL6wtucrVOf20vVM**/edit?usp=sharing

### Create google forms to capture submissions and evaluate results

This application will edit the google forms to update the month to the relevant month when it is run. Google doesn't have an easy way to share admins of forms, so you'll have to create your own test forms.

#### Submitter Form

Create a new form, with the following questions (all should be required):

- Email
- Your Discord Handle
- "Dear Ambassador,
  Please add text to your contributions during the month"
- "Dear Ambassador,
  Please add links your contributions during the month"

Click on the "Responses" tab to change the form Responses to write to your test submission sheet by choosing Select destination for responses (and choosing your Registry sheet cloned above, or another sheet based on your testing needs) from the ellipsis menu.
Update EnvironmentVariablesTest.js with your form ids and links:

Again, the form id is after /forms/d in the url - for example: https://docs.google.com/forms/d/**13oDRgD2qjryfhv992ZS99zCTOHPXBxsqKAXijupHbfE**/edit
You can get the submitter links by clicking Send then the Send Via Link option. You may want to select the "shorten url" chekbox for a shorter link. (e.g. https://forms.gle/44BW8t2aWhLTrS7i6)

The Id and Link should be put in the `SUBMISSION_FORM_ID` and `SUBMISSION_FORM_URL` in EnvironmentVariablesTest.js. You can open the sheet you are writing to to get the sheet name for the response data. It will be something like `Form Response 1`, and the sheet name shoudl be populated in `FORM_RESPONSES_SHEET_NAME`. The worksheet the responses are written to should be updated in `AMBASSADORS_SUBMISSIONS_SPREADSHEET_ID`.

#### Evaluator Form

Create a new form, with the following questions (all but the last should be required.) The question text must match the below exactly.

- Email
- Discord handle of the ambassador you are evaluating?
- Please assign a grade on a scale of 0 to 5
- Remarks (optional)

Update EnvironmentVariablesTest.js with your form ids and links for `EVALUATION_FORM_ID` and `EVALUATION_FORM_URL`. You can open the sheet you are writing to to get the sheet name for the response data. It will be something like `Form Response 2`, and the sheet name should be populated in `EVAL_FORM_RESPONSES_SHEET_NAME`. The worksheet the responses are written to should be updated in `EVALUATION_RESPONSES_SPREADSHEET_ID`.

### Prepare test data

Modify your Registry sheet with data to suit your test cases.
If you are not going to fill out the google forms manually, also update the From Responses sheets with the relevant test data. Be careful to ensure proper formatting for the Timestamp column.

## To run the process:

From vscode, you can simply run `clasp push` and then `clasp open` to open the code editor and run in debug.

From the spreadsheet, that was created on your initial clasp setup, you should now see a menu called "Ambassador OS" with menu items to run the process from the google sheet.

Choose the relevant menu option to send out emails requesting evaluations or submissions.

When requesting Evaluations, the script will add a new sheet to the Submissions spreadsheet called Review Log to record which ambassadors received which evaluation request.
