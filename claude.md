# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Google Apps Script project that implements the Autonomys Ambassador OS Peer Review system. It manages ambassador submissions, evaluations, and scoring through Google Sheets, Forms, and automated email workflows. The process has 3 major stages:

1. Each month, all active ambassadors are asked to submit evidence via a google form of their work for the prior month. Module1-Submissions.js handles the submission processes.
2. After submissions are recevied (or the deadline passes), 3 ambassadors are randomly selected to evaluate each submission (again via a google form). Module2-Evaluations.js manages the evaluation process.
3. After evaluations are received, a scoring process is run to update the scores and check compliance with the rules. This is managed by Module3-Compliance.js.

## Development Commands

- `clasp push` - Deploy code to Google Apps Script
- `clasp open` - Open the Apps Script editor for debugging
- `clasp pull` - Pull latest code from Google Apps Script
- No traditional build/test commands - this is a Google Apps Script project

## Architecture

### Core Modules

- **SharedUtilities.js** - Global variables, utility functions, and shared configuration
- **Module1-Submissions.js** - Handles ambassador contribution submissions
- **Module2-Evaluations.js** - Manages peer evaluation workflows
- **Module3-Compliance.js** - Compliance checking and penalty point management
- **Module4-AdvanceNotice.js** - Advance notice handling for submissions
- **Module5-ConflictResolutionTeam.js** - CRT member management
- **Module6-NotionSync.js** - Notion integration for data synchronization

### Configuration System

- **EnvironmentVariablesProd.js** - Production configuration (sheet IDs, form IDs, etc.)
- **EnvironmentVariablesTest.js** - Test configuration for development
- Global `testing` constant in SharedUtilities.js controls which environment is loaded

### Data Sources

The system integrates with multiple Google resources:

- Ambassador Registry spreadsheet
- Ambassadors' Scores spreadsheet
- Submission and Evaluation forms
- Form response collection sheets

## Development Guidelines

### Column Access Pattern

- Never use magic strings or hard-coded indices
- Use `getRequiredColumnIndexByName()` for required columns
- Use `getColumnIndexByName()` for optional columns
- All column names are defined as constants in SharedUtilities.js

### Ambassador Lookup

- Use `lookupEmailAndDiscord(identifier)` to get email/Discord from either identifier
- Use `getCurrentCRTMemberEmails()` for CRT member information

### Shared Functions and Configuration Management

SharedUtilities.js is used for utiltiy functions used across

Configuration is managed in Config-_.js files. Config-Initialize.js contains the vars, then Config-[Name].js contains the specific configuration parameters used for different people, with Config-Production.js being the production values. When adding new configuration, it should be added to Config-Initialize.js as an empty declaration, then added to all other Config-_.js files as a placeholder. Config-Template.js is a template used to create new Config-[name].js files in case of a new tester.

Constants.js should be used to contain all constants that don't vary by tester.

- Declare global variables using `var` in SharedUtilities.js
- Static values go in SharedUtilities.js
- Environment-specific values go in Config-\*.js:
- - Config-Initialize initialized the variables
- - Config-Template is cloned for different users that may test the system
- - Config-Production has production values
- Replicate new config variables in all environment files

### Code Standards

- Document new functions with JSDoc-style comments
- Prefer utility functions for reusable logic - kept in SharedUtilities.js
- Never hard-coded column indices or header names. Insteaad, keep sheet/column access dynamic and robust to schema changes. Use Constants for column names. Set Constants in Config.\* files.
- Use descriptive variable names (no abbreviations)

### Testing

- Set CONFIG_NAME is main.js to set the the configuration to the relevant tester.
- Each config file can control test mode, which will send eamils to a tester email instead. The config may also turn off email sendign altogether with these constants: TESTING, SEND_EMAIL, TESTER_EMAIL.

## Deployment

1. Configure appropriate environment variables, set config to "Production" or the relevant Tester in main.js.
2. Run `clasp push` to deploy to Google Apps Script
3. For production: ensure `testing = false` and `SEND_EMAIL = true` in the Config-Production.js.
