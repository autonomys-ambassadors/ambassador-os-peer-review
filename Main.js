// (( Configuration System ))
// Configuration selection - set this to the name of your configuration:
// 'Production' for live environment, or any tester name like 'Jonathan', 'Wilyam', etc.
const CONFIG_NAME = 'Jonathan'; // Available: 'Production', 'Jonathan', 'Wilyam' - add more in Config-[Name].js files

// Note: All configuration variables are declared in Config-Initialize.js
// Their values are set by the configuration functions in Config-[Name].js files

// Unified configuration loader - calls the appropriate configuration function based on CONFIG_NAME
switch (CONFIG_NAME) {
  case 'Production':
    setProductionVariables();
    break;
  case 'Jonathan':
    setJonathanVariables();
    break;
  case 'Wilyam':
    setWilyamVariables();
    break;
  default:
    throw new Error(
      `Unknown configuration: "${CONFIG_NAME}". Available configurations: 'Production', 'Jonathan', 'Wilyam'. To add a new configuration, create a Config-[Name].js file with a set[Name]Variables() function.`
    );
}

// Log which configuration is active
Logger.log(`Configuration loaded: ${CONFIG_NAME}`);
if (typeof TESTER_EMAIL !== 'undefined') {
  Logger.log(`Tester email: ${TESTER_EMAIL}`);
}
