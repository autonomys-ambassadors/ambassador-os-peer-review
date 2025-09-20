/**
 * Setup script to configure secure properties via clasp parameters
 *
 * Usage from command line:
 * clasp run setCodaCredentials -p '["your_api_token", "your_doc_id"]'
 * clasp run setCustomProperty -p '["PROPERTY_NAME", "property_value"]'
 * clasp run checkAllProperties
 *
 * This will securely store credentials in Google Apps Script's PropertiesService
 */

/**
 * Sets Coda credentials in PropertiesService
 * @param {string} apiToken - Your Coda API token
 * @param {string} docId - Your Coda document ID
 */
function setCodaCredentials(apiToken, docId) {
  try {
    if (!apiToken || !docId) {
      throw new Error('Both apiToken and docId parameters are required');
    }

    // Store in PropertiesService
    PropertiesService.getScriptProperties().setProperties({
      'CODA_API_TOKEN': apiToken,
      'CODA_DOC_ID': docId
    });

    console.log('✅ Coda credentials successfully stored in PropertiesService');
    console.log(`   API Token: ${apiToken.substring(0, 8)}...`);
    console.log(`   Doc ID: ${docId}`);

    // Test the configuration by loading it
    loadCodaConfiguration();

  } catch (error) {
    console.error('❌ Error setting up Coda credentials:', error.toString());
    throw error;
  }
}

/**
 * Sets a custom property in PropertiesService
 * @param {string} propertyName - The property name
 * @param {string} propertyValue - The property value
 */
function setCustomProperty(propertyName, propertyValue) {
  try {
    if (!propertyName || !propertyValue) {
      throw new Error('Both propertyName and propertyValue parameters are required');
    }

    PropertiesService.getScriptProperties().setProperty(propertyName, propertyValue);

    // Mask sensitive-looking values in logs
    const isSensitive = /token|key|secret|password|credential|api/i.test(propertyName);
    const displayValue = isSensitive && propertyValue.length > 8
      ? `${propertyValue.substring(0, 8)}...`
      : propertyValue;

    console.log(`✅ Property '${propertyName}' set successfully`);
    console.log(`   Value: ${displayValue}`);

  } catch (error) {
    console.error(`❌ Error setting property '${propertyName}':`, error.toString());
    throw error;
  }
}

/**
 * Displays current Coda configuration status
 */
function checkCodaConfiguration() {
  const properties = PropertiesService.getScriptProperties();
  const token = properties.getProperty('CODA_API_TOKEN');
  const docId = properties.getProperty('CODA_DOC_ID');

  if (token && docId) {
    console.log('✅ Coda integration is configured');
    console.log(`   API Token: ${token.substring(0, 8)}...`);
    console.log(`   Doc ID: ${docId}`);
  } else {
    console.log('❌ Coda integration not configured');
    console.log('   Run: clasp run setCodaCredentials -p \'["your_token", "your_doc_id"]\'');
  }
}

/**
 * Displays all stored properties (masks sensitive values)
 */
function checkAllProperties() {
  try {
    const properties = PropertiesService.getScriptProperties().getProperties();
    const propertyNames = Object.keys(properties);

    if (propertyNames.length === 0) {
      console.log('❌ No properties found in PropertiesService');
      console.log('   Use setCodaCredentials or setCustomProperty to add some');
      return;
    }

    console.log(`✅ Found ${propertyNames.length} stored properties:`);

    propertyNames.forEach(name => {
      const value = properties[name];
      const isSensitive = /token|key|secret|password|credential|api/i.test(name);
      const displayValue = isSensitive && value && value.length > 8
        ? `${value.substring(0, 8)}...`
        : value;

      console.log(`   ${name}: ${displayValue}`);
    });

  } catch (error) {
    console.error('❌ Error checking properties:', error.toString());
    throw error;
  }
}

/**
 * Removes a property from PropertiesService
 * @param {string} propertyName - The property name to remove
 */
function removeProperty(propertyName) {
  try {
    if (!propertyName) {
      throw new Error('propertyName parameter is required');
    }

    PropertiesService.getScriptProperties().deleteProperty(propertyName);
    console.log(`✅ Property '${propertyName}' removed successfully`);

  } catch (error) {
    console.error(`❌ Error removing property '${propertyName}':`, error.toString());
    throw error;
  }
}