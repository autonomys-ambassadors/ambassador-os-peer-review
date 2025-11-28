claude.md

This a a google appscript project hosted direclty in a google sheet.
All code must be limited to what can run in appscript.

This script manages a peer review process for the autonomys network ambassador program. The process has 3 major stages:

1. Each month, all active ambassadors are asked to submit evidence via a google form of their work for the prior month. Module1-Submissions.js handles the submission processes.
2. After submissions are recevied (or the deadline passes), 3 ambassadors are randomly selected to evaluate each submission (again via a google form). Module2-Evaluations.js manages the evaluation process.
3. After evaluations are received, a scoring process is run to update the scores and check compliance with the rules. This is managed by Module3-Compliance.js.

SharedUtilities.js is used for utiltiy functions used across

Configuration is managed in Config-_.js files. Config-Initialize.js contains the vars, then Config-[Name].js contains the specific configuration parameters used for different people, with Config-Production.js being the production values. When adding new configuration, it should be added to Config-Initialize.js as an empty declaration, then added to all other Config-_.js files as a placeholder. Config-Template.js is a template used to create new Config-[name].js files in case of a new tester.

Constants.js should be used to contain all constants that don't vary by tester.
