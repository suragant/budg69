// app_errors.gs
// Centralized error capture and sheet logging.

function logErrorToSheet(errorObj) {
  try {
    const ss = getSpreadsheet();
    let errorLogSheet = resolveSheet(CONFIG.SHEETS.ERROR_LOG);
    if (!errorLogSheet) {
      errorLogSheet = ss.insertSheet(CONFIG.SHEETS.ERROR_LOG);
      errorLogSheet.appendRow(['Timestamp', 'User Email', 'Function', 'Error Message', 'Stack Trace', 'Context']);
    }
    errorLogSheet.appendRow([
      new Date(),
      getUserEmail() || 'system',
      errorObj.functionName || 'unknown',
      errorObj.message || '',
      errorObj.stack || '',
      JSON.stringify(errorObj.context || {})
    ]);
  } catch (e) {
    Logger.log('Error logging failed: ' + e.toString());
  }
}

function handleError(functionName, error, context = {}) {
  const errorDetails = {
    functionName,
    message: error?.message || error?.toString() || 'Unknown error',
    stack: error?.stack || '',
    timestamp: new Date().toISOString(),
    userId: getUserEmail(),
    context
  };
  Logger.log(JSON.stringify(errorDetails));
  logErrorToSheet(errorDetails);
  return errorDetails;
}
