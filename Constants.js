/**
 * Application Constants
 *
 * This file contains all application constants that do not vary between environments.
 * These are static values used throughout the application.
 *
 * Note: Configuration variables that change between Production/Test environments
 * are declared in Config-Initialize.js and set by Config-[Name].js files.
 */

// ===== UI Constants =====

const ButtonSet = {
  OK: 'OK',
  OK_CANCEL: 'OK_CANCEL',
  YES_NO: 'YES_NO',
  YES_NO_CANCEL: 'YES_NO_CANCEL',
};

const ButtonResponse = {
  OK: 'ok',
  CANCEL: 'cancel',
  YES: 'yes',
  NO: 'no',
};

// ===== Email Templates =====

// Request Submission Email Template
const REQUEST_SUBMISSION_EMAIL_TEMPLATE = `
<p>Dear {AmbassadorDiscordHandle},</p>

<p>Please submit your deliverables for {Month} {Year} using the link below:</p>
<p><a href="{SubmissionFormURL}">Submission Form</a></p>

<p>The deadline is {SUBMISSION_DEADLINE_DATE}.</p>

<p>Thank you,<br>
Ambassador Program Team</p>
`;

// Request Evaluation Email Template
const REQUEST_EVALUATION_EMAIL_TEMPLATE = `
<p>Dear {AmbassadorDiscordHandle},</p>
<p>Please review the following deliverables for the month of <strong>{Month}</strong> by:</p>

<p>
<strong>{AmbassadorSubmitter}<br><br>
Primary Team:  {PrimaryTeam}<br><br>
Primary Team Responsibilities:</strong><br>{PrimaryTeamResponsibilities}<br><br>
</p>

<strong>Work Submitted:</strong><br>
<p>{SubmissionsList}</p>

<p>Assign a grade using the form:</p>
<p><a href="{EvaluationFormURL}">Evaluation Form</a></p>

<p>The deadline is {EVALUATION_DEADLINE_DATE}.</p>

<p>Thank you,<br>Ambassador Program Team</p>
`;

// Reminder Email Template
const REMINDER_EMAIL_TEMPLATE = `
Hi there! Just a friendly reminder that we are still waiting for your response to the Request for Submission/Evaluation. Please respond soon to avoid any penalties. Thank you!.
`;

// Penalty Warning Email Template
const PENALTY_WARNING_EMAIL_TEMPLATE = `
Dear Ambassador,
You have been assessed one penalty point for failing to meet Submission or Evaluation deadlines. Further penalties may result in expulsion from the program. Please be vigillant.
`;

// Expulsion Email Template
const EXPULSION_EMAIL_TEMPLATE = `
Dear {Discord Handle},
We regret to inform you that you have been expelled from the program for Failure to Participate according to Article 2, Section 10 of the Bylaws as of {Expulsion Date}.

If you believe the expulsion is incorrect, you have the right to appeal your expulsion through the Conflict Resolution Team. Please email the Sponsor Representative at {Sponsor Email} including the reason for your appeal and any supporting documentation that the CRT should consider if you choose to appeal.

We acknowledge and thank you for your contributions to the project as an Ambassador from {Start Date} to {Expulsion Date}.

Autonomys Community Team`;

// Notify Upcoming Peer Review Email Template
const NOTIFY_UPCOMING_PEER_REVIEW = `
Dear Ambassador,
By this we notify you about upcoming Peer Review mailing, please be vigilant!
`;

// Exemption from Evaluation Email Template
const EXEMPTION_FROM_EVALUATION_TEMPLATE = `
Dear Ambassador, you have been relieved of the obligation to evaluate your colleagues this month.
`;

// CRT Referral for Inadequate Contribution Email Template
const CRT_INADEQUATE_CONTRIBUTION_EMAIL_TEMPLATE = `
To: CRT Members and accused Ambassador and Sponsor,<br><br>
Ambassador {discordHandle} is being referred to the CRT due to Inadequate Contribution as defined in the bylaws in Article 2.<br>
{discordHandle} has scored below {inadequateContributionScoreThreshold} a total of {inadequateContributionCount} times in the last 6 evaluation months.<br>
{crtNote}
`;

// Inadequate Contribution Notification Email Template (sent directly to ambassador)
const INADEQUATE_CONTRIBUTION_NOTIFICATION_EMAIL_TEMPLATE = `
Hello Ambassador,<br><br>
I write to inform you that the AmbassadorOS process has lodged a formal case to the Conflict Resolution Team based on {monthName} DELIVERABLES triggering Inadequate Contribution. You have scored below 3 in more than 2 of the last 6 months.<br><br>
Peer ambassadors noticing deceptive or low-quality contributions often feel disappointed by the lack of fairness and accountability expected in the Ambassador Program.<br><br>
I look forward to your response within 3 business days ({deadlineDate}).<br><br>
Thank you for your attention to this matter.<br><br>
The Autonomys Community Team
`;

// Primary team Responsibilities
const PrimaryTeamResponsibilities = {
  support: `Provide peer-to-peer support and create support materials (e.g., articles),<br>
      Gather information and help investigate and solve technical issues,<br>
      Assist or directly participate in technical development of the project,<br>
      Answer questions in Discord, Telegram, and the Networks forum,<br>
      Moderate Telegram and Discord channels,<br>
      Communicate about current releases and important events`,
  content: `Create and improve an educational plan for onboarding new Apprentices,<br>
      Develop materials, resources, and documentation on the protocol, Program, and community,<br>
      Create high-quality content to educate the community about the Network,<br>
      Cultivate content creators by recognizing and promoting users with the Content Creator role`,
  engagement: `Promote the growth of the Network by establishing connections with the community,<br>
      Identify target audiences and develop strategies to attract them to the Network,<br>
      Create and disseminate high-quality content across various platforms,<br>
      Increase user engagement and encourage active community participation,<br>
      Act as a voice for the Network, ensuring smooth communication among stakeholders`,
  onboarding: `Create and administer Ambassador selection processes,<br>
      Introduce and integrate Apprentices and new Ambassadors to the Program,<br>
      Recruit new ambassador cohorts and host events/workshops,<br>
      Mentor Apprentice Ambassadors and develop peer relationships,<br>
      Collaborate with the Content & Education team to keep Ambassadors updated`,
  governance: `Create and maintain the Bylaws and facilitate General Assembly operations,<br>
      Develop transparent systems and processes to implement the Bylaws,<br>
      Administer processes and evaluate adherence to Ambassador Rights and Obligations`,
};

// ===== Pattern Constants =====

/**
 * ISO 8601 timestamp pattern used to identify supplemental evaluation window columns.
 * Format: YYYY-MM-DDTHH:mm:ss±HH:MM (e.g., "2024-11-01T14:30:00-08:00")
 * This pattern is used in Review Log column headers to mark supplemental evaluation windows.
 * All supplemental column headers are created with this full format including timezone offset.
 */
const ISO_8601_TIMESTAMP_PATTERN = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}[+-]\d{2}:\d{2}$/;

/**
 * Tests if a string matches the ISO 8601 timestamp pattern for supplemental columns.
 * @param {string} str - The string to test
 * @returns {boolean} - True if the string matches the pattern
 */
function isSupplementalColumnHeader(str) {
  if (!str) return false;
  return ISO_8601_TIMESTAMP_PATTERN.test(str.toString());
}
