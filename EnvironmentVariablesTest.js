/**
 * DEPRECATED: This file has been replaced by the unified configuration system
 * 
 * The configuration system has been moved to SharedUtilities.js and individual Config-[Name].js files.
 * 
 * To use the new system:
 * 1. Set CONFIG_NAME in SharedUtilities.js to your desired configuration
 * 2. Available configurations: 'Production', 'Jonathan', 'Wilyam'
 * 3. To add new configurations, create Config-[YourName].js with a set[YourName]Variables() function
 * 
 * This file is kept for reference but is no longer used by the system.
 */

function setTestVariables() {
  throw new Error('This function is deprecated. Configuration is now handled in SharedUtilities.js using the CONFIG_NAME constant. Set CONFIG_NAME to your desired configuration name.');
}