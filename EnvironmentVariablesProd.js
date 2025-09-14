/**
 * DEPRECATED: This file has been replaced by Config-Production.js
 * 
 * The configuration system has been unified. Production configuration is now in:
 * Config-Production.js with setProductionVariables() function
 * 
 * To use production configuration:
 * Set CONFIG_NAME = 'Production' in SharedUtilities.js
 * 
 * This file is kept for reference but is no longer used by the system.
 */

function setProductionVariables() {
  throw new Error('This function is deprecated. Production configuration is now in Config-Production.js. Set CONFIG_NAME = \'Production\' in SharedUtilities.js to use it.');
}