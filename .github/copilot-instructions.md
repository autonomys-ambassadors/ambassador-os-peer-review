# Google Apps Script - Autonomys Ambassador Peer Review

This is a Google Apps Script project hosted directly in a Google Sheet. **All code must be compatible with Apps Script runtime** (no Node.js APIs, file system access, etc.).

## Architecture Overview

This system manages a monthly peer review workflow for Autonomys Network Ambassadors using Google Sheets, Forms, and automated email triggers. The process flows through three major stages:

1. **Submissions** (`Module1-Submissions.js`) - Request monthly contribution submissions from active ambassadors via Google Form
2. **Evaluations** (`Module2-Evaluations.js`) - Randomly assign 3 peer evaluators per submission, send evaluation requests via Google Form
3. **Compliance** (`Module3-Compliance.js`) - Calculate scores, apply penalty points, check expulsion criteria, sync data to Notion

### Data Flow

```
Notion Database → Registry Sheet → Submission Requests → Form Responses
                                  ↓
                     Evaluation Assignments → Evaluation Responses
                                  ↓
                        Scoring & Compliance → Overall Score Sheet
```

### Key Integration Points

- **Notion API**: Source of truth for ambassador data (sync via `NotionSync.js`)
- **Google Forms**: Submission and evaluation collection (forms auto-updated with current month)
- **Google Sheets**: 4 spreadsheets (Registry, Scores, Submission Responses, Evaluation Responses)
- **Time-based Triggers**: Automated reminders, deadline enforcement, form submission processing

## Configuration System

Configuration uses a multi-file pattern to support testing and production environments:

1. **`Config-Initialize.js`** - Declares ALL configuration variables as `var` (required for global scope in Apps Script)
2. **`Config-[Name].js`** - Sets values for specific environment (Production, Jonathan, Wilyam, etc.)
3. **`Main.js`** - Sets `CONFIG_NAME` constant to select which configuration to load

**Adding new configuration:**

```javascript
// 1. Declare in Config-Initialize.js
var NEW_CONFIG_VAR;

// 2. Set in EVERY Config-*.js file
function setProductionVariables() {
  NEW_CONFIG_VAR = 'production-value';
  // ...
}

// 3. Use constants for values that don't vary between environments
// Constants.js - for static values like email templates, UI constants
```

**Configuration control flow:**

- `CONFIG_NAME` in `Main.js` → calls `set[Name]Variables()` → populates all global vars
- `TESTING` flag controls email redirection to `TESTER_EMAIL`
- `SEND_EMAIL` flag can disable emails entirely for dry runs

## Development Patterns

### Column Access Pattern

**Never hardcode column indices or exact header strings.** Always use dynamic column lookup:

```javascript
// ✅ Correct - dynamic column lookup
const emailColIndex = getRequiredColumnIndexByName(sheet, AMBASSADOR_EMAIL_COLUMN);
const row = sheet.getRange(2, emailColIndex).getValue();

// ❌ Wrong - hardcoded indices or column names
const row = sheet.getRange(2, 3).getValue(); // Brittle
const headers = sheet.getRange(1, 1, 1, 10).getValues()[0];
const emailIndex = headers.indexOf('Ambassador Email Address'); // Duplicates config
```

Column names defined in configuration files as constants (e.g., `AMBASSADOR_EMAIL_COLUMN = 'Ambassador Email Address'`).

### Shared Utilities Pattern

All cross-module utilities live in `SharedUtilities.js`:

- Sheet access functions: `getRegistrySheet()`, `getOverallScoreSheet()`, etc.
- Email sending: `sendEmailNotification()` with test mode support
- Column lookup: `getRequiredColumnIndexByName()`, `getColumnIndexByName()`
- Time utilities: `minutesToMilliseconds()`, date formatting helpers

### Month Sheet Management

- Monthly scores stored in sheets named "MMMM yyyy" (e.g., "January 2026")
- Overall Score sheet has date-based columns (first day of month) formatted as "MMMM yyyy"
- Always use `getReportingMonthFromRequestLog()` to determine which month is being processed
- Creates month sheets dynamically during evaluation requests

## Deployment & Testing

### Setup Commands

```bash
# Install clasp globally
npm install -g @google/clasp
clasp login

# Clone existing project (requires Script ID from Google Apps Script)
# Create .clasp.json with: {"scriptId": "YOUR_SCRIPT_ID", "rootDir": "./"}

# Push code changes
clasp push

# Open in Apps Script editor for debugging
clasp open

# Pull latest from Apps Script
clasp pull
```

### Testing Workflow

1. Copy prototype Registry/Scores sheets for testing (links in README.md)
2. Create `Config-[YourName].js` based on `Config-Template.js`
3. Set your sheet IDs, form IDs, and `TESTER_EMAIL`
4. In `Main.js`, set `CONFIG_NAME = 'YourName'`
5. Set `TESTING = true` and optionally `SEND_EMAIL = false` in your config
6. Run `clasp push` to deploy
7. Use menu items in spreadsheet to trigger workflows

### Debugging

- **Execution logs**: Apps Script editor → Executions tab (shows all Logger.log() output)
- **No traditional debugger**: Use `Logger.log()` extensively
- **Email testing**: `TESTING=true` redirects all emails to `TESTER_EMAIL`
- **Trigger management**: Menu → 🔧️Delete Existing Triggers (clears orphaned time-based triggers)

## Critical Patterns

### Trigger Management

- Time-based triggers created programmatically for reminders and deadlines
- Always delete existing triggers before creating new ones (`deleteExistingTriggers()`)
- Form submission triggers set up via `setupEvaluationResponseTrigger()`
- Trigger cleanup function exposed in UI menu for manual intervention

### Error Handling & User Interaction

```javascript
// Use UI alert/prompt wrappers that also log
alertAndLog('Error title', 'Message');
const response = promptAndLog('Title', 'Message', ButtonSet.YES_NO);

if (response === ButtonResponse.YES) {
  // User confirmed
}
```

### Penalty Points & Compliance

- 6-month rolling window for penalty calculations (`COMPLIANCE_PERIOD_MONTHS`)
- Missed submission: 1 point, Missed evaluation: 1 point, Both: 2 points
- Expulsion criteria checked in `Module3-Compliance.js::expelAmbassadors()`
- Status synced back to Notion after compliance updates

### Data Synchronization Order

Critical sequence in `runComplianceAudit()`:

1. Check evaluation window status
2. Sync Notion → Registry (latest ambassador data)
3. Sync Registry → Overall Score (ensures all ambassadors present)
4. Copy Final Scores → Overall Score month column
5. Calculate penalty points
6. Check expulsion criteria
7. Sync status back to Registry and Overall Score
8. Publish anonymous scores

## Menu Structure

Main UI entry points (`onOpen()` in `Main.js`):

- Request Submissions → `requestMonthlySubmissions()`
- Request Evaluations → `requestEvaluationsModule()`
- Compliance Audit → `runComplianceAudit()`
- 🔧️ prefix = maintenance/admin functions

## Common Gotchas

- **Global scope**: All `.js` files loaded together; declare vars in `Config-Initialize.js` only
- **No async/await**: Apps Script uses older JavaScript; use synchronous patterns
- **Flush required**: Call `SpreadsheetApp.flush()` between dependent sheet operations
- **Time zones**: Always use `getProjectTimeZone()` for date formatting consistency
- **Form questions**: Must match exact question text from Google Forms (set in config constants)
