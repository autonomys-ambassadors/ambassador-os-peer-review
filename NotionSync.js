// MODULE: Notion Sync
// Handles synchronization of ambassador data from Notion to Google Sheets Registry

/**
 * Main entry point for Notion sync process.
 * Fetches ambassador data from Notion and synchronizes it with the Google Sheets Registry.
 * Throws exception on critical errors to halt the overall score sync process.
 */
function syncRegistryWithNotion() {
  Logger.log('=== Starting Notion Sync Process ===');

  try {
    // Get Notion API key from script properties
    const notionApiKey = getNotionApiKey();
    if (!notionApiKey) {
      throw new Error('NOTION_API_KEY not found in script properties');
    }

    // Fetch all ambassadors from Notion
    const notionAmbassadors = fetchNotionAmbassadors(notionApiKey);
    Logger.log(`Fetched ${notionAmbassadors.length} ambassadors from Notion`);

    // Get registry sheet and data
    const registrySheet = getRegistrySheet();
    if (!registrySheet) {
      throw new Error('Registry sheet not found');
    }

    const registryData = registrySheet.getDataRange().getValues();
    const registryHeaders = registryData[0];

    // Get all required column indices
    const columnIndices = getNotionSyncColumnIndices(registrySheet, registryHeaders);

    // Track changes for summary
    let addedCount = 0;
    let updatedCount = 0;
    let expelledCount = 0;

    // Process each Notion record
    notionAmbassadors.forEach((notionRecord) => {
      try {
        const discordHandle = notionRecord.discord || 'Unknown';
        const status = notionRecord.status;

        // Find matching row in registry
        const rowIndex = findMatchingRegistryRow(notionRecord, registryData, columnIndices);

        if (status === AMBASSADOR_STATUS_INACTIVE) {
          // Handle inactive ambassadors (mark as expelled if not already)
          const wasExpelled = handleInactiveAmbassador(notionRecord, registrySheet, columnIndices, rowIndex);
          if (wasExpelled) expelledCount++;
        } else {
          // Sync active ambassadors
          if (rowIndex === -1) {
            // Add new ambassador
            addNewAmbassadorRow(notionRecord, registrySheet, columnIndices);
            addedCount++;
            Logger.log(`Added new ambassador: ${discordHandle}`);
          } else {
            // Update existing ambassador
            const wasUpdated = syncAmbassadorFromNotion(notionRecord, registrySheet, columnIndices, rowIndex);
            if (wasUpdated) updatedCount++;
          }
        }
      } catch (error) {
        Logger.log(`Error processing ambassador ${notionRecord.discord || notionRecord.email}: ${error.message}`);
        // Continue processing other ambassadors
      }
    });

    // Log summary
    Logger.log('=== Notion Sync Summary ===');
    Logger.log(`Added: ${addedCount} ambassadors`);
    Logger.log(`Updated: ${updatedCount} ambassadors`);
    Logger.log(`Expelled: ${expelledCount} ambassadors`);
    Logger.log('=== Notion Sync Completed Successfully ===');
  } catch (error) {
    const errorMsg = `CRITICAL: Notion sync failed: ${error.message}`;
    Logger.log(errorMsg);
    Logger.log(`Stack trace: ${error.stack || 'Not available'}`);
    throw new Error(errorMsg);
  }
}

/**
 * Retrieves the Notion API key from script properties.
 * @returns {string} The Notion API key
 */
function getNotionApiKey() {
  return PropertiesService.getScriptProperties().getProperty('NOTION_API_KEY');
}

/**
 * Fetches all ambassador records from Notion database.
 * @param {string} apiKey - Notion API key
 * @returns {Array} Array of ambassador objects
 */
function fetchNotionAmbassadors(apiKey) {
  const url = `https://api.notion.com/v1/databases/${NOTION_DATABASE_ID}/query`;

  const options = {
    method: 'post',
    headers: {
      Authorization: `Bearer ${apiKey}`,
      'Notion-Version': '2022-06-28',
      'Content-Type': 'application/json',
    },
    payload: JSON.stringify({
      page_size: 100,
    }),
    muteHttpExceptions: true,
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const statusCode = response.getResponseCode();

    if (statusCode !== 200) {
      throw new Error(`Notion API returned status ${statusCode}: ${response.getContentText()}`);
    }

    const data = JSON.parse(response.getContentText());
    return parseNotionResponse(data);
  } catch (error) {
    throw new Error(`Failed to fetch from Notion API: ${error.message}`);
  }
}

/**
 * Parses Notion API response into usable ambassador objects.
 * @param {Object} response - Notion API response
 * @returns {Array} Array of parsed ambassador objects
 */
function parseNotionResponse(response) {
  const ambassadors = [];

  if (!response.results || !Array.isArray(response.results)) {
    Logger.log('Warning: No results found in Notion response');
    return ambassadors;
  }

  response.results.forEach((page) => {
    try {
      const props = page.properties;

      const ambassador = {
        notionId: page.id,
        number: getNotionPropertyValue(props[NOTION_NUMBER_COLUMN], 'number'),
        email: getNotionPropertyValue(props[NOTION_EMAIL_COLUMN], 'email'),
        discord: getNotionPropertyValue(props[NOTION_DISCORD_COLUMN], 'rich_text'),
        status: getNotionPropertyValue(props[NOTION_STATUS_COLUMN], 'select'),
        primaryTeam: getNotionPropertyValue(props[NOTION_PRIMARY_TEAM_COLUMN], 'select'),
        secondaryTeam: getNotionPropertyValue(props[NOTION_SECONDARY_TEAM_COLUMN], 'select'),
        startDate: getNotionPropertyValue(props[NOTION_START_DATE_COLUMN], 'date'),
      };

      ambassadors.push(ambassador);
    } catch (error) {
      Logger.log(`Error parsing Notion record: ${error.message}`);
    }
  });

  return ambassadors;
}

/**
 * Extracts value from Notion property based on its type.
 * @param {Object} property - Notion property object
 * @param {string} type - Property type (number, email, rich_text, select, date)
 * @returns {*} Extracted value or null
 */
function getNotionPropertyValue(property, type) {
  if (!property) return null;

  try {
    switch (type) {
      case 'number':
        return property.number;
      case 'email':
        return property.email;
      case 'rich_text':
        return property.rich_text && property.rich_text[0] ? property.rich_text[0].plain_text : null;
      case 'select':
        return property.select ? property.select.name : null;
      case 'date':
        return property.date ? property.date.start : null;
      default:
        return null;
    }
  } catch (error) {
    Logger.log(`Error extracting property value of type ${type}: ${error.message}`);
    return null;
  }
}

/**
 * Maps Notion team name to Google Sheet team name.
 * @param {string} notionTeam - Team name from Notion
 * @returns {string} Mapped team name for sheet
 */
function mapNotionTeamToSheet(notionTeam) {
  if (!notionTeam) return '';

  return NOTION_TO_SHEET_TEAM_MAPPING[notionTeam] || notionTeam;
}

/**
 * Gets all required column indices for Notion sync.
 * @param {Sheet} registrySheet - Registry sheet
 * @param {Array} headers - Header row array
 * @returns {Object} Object containing all column indices
 */
function getNotionSyncColumnIndices(registrySheet, headers) {
  return {
    notionId: getColumnIndexByName(registrySheet, REGISTRY_NOTION_ID_COLUMN),
    number: getRequiredColumnIndexByName(registrySheet, AMBASSADOR_ID_COLUMN),
    email: getRequiredColumnIndexByName(registrySheet, AMBASSADOR_EMAIL_COLUMN),
    discord: getRequiredColumnIndexByName(registrySheet, AMBASSADOR_DISCORD_HANDLE_COLUMN),
    status: getRequiredColumnIndexByName(registrySheet, AMBASSADOR_STATUS_COLUMN),
    primaryTeam: getRequiredColumnIndexByName(registrySheet, AMBASSADOR_PRIMARY_TEAM_COLUMN),
    secondaryTeam: getColumnIndexByName(registrySheet, REGISTRY_SECONDARY_TEAM_COLUMN),
    startDate: getColumnIndexByName(registrySheet, REGISTRY_START_DATE_COLUMN),
  };
}

/**
 * Finds matching registry row for a Notion record.
 * Tries matching by: Number → Email → Discord Handle
 * @param {Object} notionRecord - Notion ambassador record
 * @param {Array} registryData - Registry sheet data
 * @param {Object} columnIndices - Column indices object
 * @returns {number} Row index (1-based) or -1 if not found
 */
function findMatchingRegistryRow(notionRecord, registryData, columnIndices) {
  const normalizedNotionEmail = normalizeEmail(notionRecord.email);
  const normalizedNotionDiscord = normalizeDiscordHandle(notionRecord.discord);

  for (let i = 1; i < registryData.length; i++) {
    const row = registryData[i];

    // Try match by Number (Unique ID)
    if (
      notionRecord.number &&
      String(row[columnIndices.number - 1]) === String(notionRecord.number)
    ) {
      Logger.log(`Matched by Number: ${notionRecord.number} at row ${i + 1}`);
      return i + 1;
    }

    // Try match by Email
    const sheetEmail = normalizeEmail(row[columnIndices.email - 1]);
    if (normalizedNotionEmail && sheetEmail === normalizedNotionEmail) {
      Logger.log(`Matched by Email: ${notionRecord.email} at row ${i + 1}`);
      return i + 1;
    }

    // Try match by Discord Handle
    const sheetDiscord = normalizeDiscordHandle(row[columnIndices.discord - 1]);
    if (normalizedNotionDiscord && sheetDiscord === normalizedNotionDiscord) {
      Logger.log(`Matched by Discord: ${notionRecord.discord} at row ${i + 1}`);
      return i + 1;
    }
  }

  return -1; // No match found
}

/**
 * Syncs an existing ambassador row from Notion data.
 * @param {Object} notionRecord - Notion ambassador record
 * @param {Sheet} registrySheet - Registry sheet
 * @param {Object} columnIndices - Column indices object
 * @param {number} rowIndex - Row index to update
 * @returns {boolean} True if any updates were made
 */
function syncAmbassadorFromNotion(notionRecord, registrySheet, columnIndices, rowIndex) {
  let wasUpdated = false;
  const discordHandle = notionRecord.discord || 'Unknown';

  // Update Notion ID if column exists
  if (columnIndices.notionId > 0) {
    const currentNotionId = registrySheet.getRange(rowIndex, columnIndices.notionId).getValue();
    if (currentNotionId !== notionRecord.notionId) {
      registrySheet.getRange(rowIndex, columnIndices.notionId).setValue(notionRecord.notionId);
      Logger.log(`Updated Notion ID for ${discordHandle}: ${currentNotionId} → ${notionRecord.notionId}`);
      wasUpdated = true;
    }
  }

  // Update Email (only if not null)
  if (notionRecord.email) {
    const currentEmail = registrySheet.getRange(rowIndex, columnIndices.email).getValue();
    if (normalizeEmail(currentEmail) !== normalizeEmail(notionRecord.email)) {
      registrySheet.getRange(rowIndex, columnIndices.email).setValue(notionRecord.email);
      Logger.log(`Updated Email for ${discordHandle}: ${currentEmail} → ${notionRecord.email}`);
      wasUpdated = true;
    }
  }

  // Update Discord Handle (only if not null)
  if (notionRecord.discord) {
    const currentDiscord = registrySheet.getRange(rowIndex, columnIndices.discord).getValue();
    if (normalizeDiscordHandle(currentDiscord) !== normalizeDiscordHandle(notionRecord.discord)) {
      registrySheet.getRange(rowIndex, columnIndices.discord).setValue(notionRecord.discord);
      Logger.log(`Updated Discord Handle for ${discordHandle}: ${currentDiscord} → ${notionRecord.discord}`);
      wasUpdated = true;
    }
  }

  // Update Primary Team (with mapping, only if not empty)
  const mappedPrimaryTeam = mapNotionTeamToSheet(notionRecord.primaryTeam);
  if (mappedPrimaryTeam) {
    const currentPrimaryTeam = registrySheet.getRange(rowIndex, columnIndices.primaryTeam).getValue();
    if (currentPrimaryTeam !== mappedPrimaryTeam && currentPrimaryTeam !== TEAM_VALUE_EXPELLED) {
      registrySheet.getRange(rowIndex, columnIndices.primaryTeam).setValue(mappedPrimaryTeam);
      Logger.log(`Updated Primary Team for ${discordHandle}: ${currentPrimaryTeam} → ${mappedPrimaryTeam}`);
      wasUpdated = true;
    }
  }

  // Update Secondary Team if column exists (with mapping)
  if (columnIndices.secondaryTeam > 0) {
    const mappedSecondaryTeam = mapNotionTeamToSheet(notionRecord.secondaryTeam);
    const currentSecondaryTeam = registrySheet.getRange(rowIndex, columnIndices.secondaryTeam).getValue();
    // Allow empty string to clear the secondary team
    if (currentSecondaryTeam !== mappedSecondaryTeam) {
      registrySheet.getRange(rowIndex, columnIndices.secondaryTeam).setValue(mappedSecondaryTeam || '');
      Logger.log(`Updated Secondary Team for ${discordHandle}: ${currentSecondaryTeam} → ${mappedSecondaryTeam || '(empty)'}`);
      wasUpdated = true;
    }
  }

  // Update Start Date if column exists
  if (columnIndices.startDate > 0 && notionRecord.startDate) {
    const currentStartDate = registrySheet.getRange(rowIndex, columnIndices.startDate).getValue();
    
    // Normalize both dates to YYYY-MM-DD format for comparison
    const notionDateStr = notionRecord.startDate; // Already in YYYY-MM-DD format from Notion
    const currentDateStr = currentStartDate 
      ? Utilities.formatDate(new Date(currentStartDate), getProjectTimeZone(), 'yyyy-MM-dd')
      : '';

    if (currentDateStr !== notionDateStr) {
      registrySheet.getRange(rowIndex, columnIndices.startDate).setValue(notionRecord.startDate);
      Logger.log(`Updated Start Date for ${discordHandle}: ${currentDateStr} → ${notionDateStr}`);
      wasUpdated = true;
    }
  }

  return wasUpdated;
}

/**
 * Handles inactive ambassadors from Notion (marks as expelled if not already).
 * @param {Object} notionRecord - Notion ambassador record
 * @param {Sheet} registrySheet - Registry sheet
 * @param {Object} columnIndices - Column indices object
 * @param {number} rowIndex - Row index or -1 if not found
 * @returns {boolean} True if ambassador was newly marked as expelled
 */
function handleInactiveAmbassador(notionRecord, registrySheet, columnIndices, rowIndex) {
  const discordHandle = notionRecord.discord || notionRecord.email || 'Unknown';

  // If no matching row found, skip (don't add inactive ambassadors)
  if (rowIndex === -1) {
    Logger.log(`Inactive ambassador not in registry, skipping: ${discordHandle}`);
    return false;
  }

  // Check current primary team
  const currentPrimaryTeam = registrySheet.getRange(rowIndex, columnIndices.primaryTeam).getValue();

  if (currentPrimaryTeam !== TEAM_VALUE_EXPELLED) {
    // Mark as expelled
    registrySheet.getRange(rowIndex, columnIndices.primaryTeam).setValue(TEAM_VALUE_EXPELLED);

    // Append to status column
    const currentStatus = registrySheet.getRange(rowIndex, columnIndices.status).getValue() || '';
    const today = Utilities.formatDate(new Date(), getProjectTimeZone(), 'yyyy-MM-dd');
    const updatedStatus = `${currentStatus} | expelled by notion sync [${today}]`;
    registrySheet.getRange(rowIndex, columnIndices.status).setValue(updatedStatus);

    Logger.log(`Marked as expelled by Notion sync: ${discordHandle}`);
    return true;
  }

  Logger.log(`Already expelled: ${discordHandle}`);
  return false;
}

/**
 * Adds a new ambassador row to the registry from Notion data.
 * @param {Object} notionRecord - Notion ambassador record
 * @param {Sheet} registrySheet - Registry sheet
 * @param {Object} columnIndices - Column indices object
 */
function addNewAmbassadorRow(notionRecord, registrySheet, columnIndices) {
  const newRowIndex = registrySheet.getLastRow() + 1;
  const discordHandle = notionRecord.discord || 'Unknown';

  // Validate required fields
  if (!notionRecord.email && !notionRecord.number) {
    throw new Error('Cannot add ambassador without email or number');
  }

  // Map team names
  const mappedPrimaryTeam = mapNotionTeamToSheet(notionRecord.primaryTeam);
  const mappedSecondaryTeam = mapNotionTeamToSheet(notionRecord.secondaryTeam);

  // Set Notion ID if column exists
  if (columnIndices.notionId > 0) {
    registrySheet.getRange(newRowIndex, columnIndices.notionId).setValue(notionRecord.notionId);
  }

  // Set Number (Unique ID) - generate from email if not provided
  const ambassadorId = notionRecord.number || generateMD5Hash(notionRecord.email);
  registrySheet.getRange(newRowIndex, columnIndices.number).setValue(ambassadorId);

  // Set Email (or empty string if null)
  registrySheet.getRange(newRowIndex, columnIndices.email).setValue(notionRecord.email || '');

  // Set Discord Handle (or empty string if null)
  registrySheet.getRange(newRowIndex, columnIndices.discord).setValue(notionRecord.discord || '');

  // Set Status
  registrySheet.getRange(newRowIndex, columnIndices.status).setValue(AMBASSADOR_STATUS_ACTIVE);

  // Set Primary Team (or empty string if null)
  registrySheet.getRange(newRowIndex, columnIndices.primaryTeam).setValue(mappedPrimaryTeam || '');

  // Set Secondary Team if column exists
  if (columnIndices.secondaryTeam > 0) {
    registrySheet.getRange(newRowIndex, columnIndices.secondaryTeam).setValue(mappedSecondaryTeam || '');
  }

  // Set Start Date if column exists
  if (columnIndices.startDate > 0 && notionRecord.startDate) {
    registrySheet.getRange(newRowIndex, columnIndices.startDate).setValue(notionRecord.startDate);
  }

  Logger.log(
    `Added new ambassador at row ${newRowIndex}: ${discordHandle} (${notionRecord.email || 'no email'}) - Team: ${mappedPrimaryTeam || 'no team'}`
  );
}
