// app_config.gs
// Shared application configuration.

const CONFIG = {
  ITEM_ID_PREFIX: 'BG69-',
  ITEM_ID_LENGTH: 3,
  SHEETS: {
    BUDGET: 'Budget',
    USERS: 'Users',
    TRANSACTION_LOG: 'Transaction_Log',
    ERROR_LOG: 'Error_Log'
  },
  ADMIN_EMAIL: 'admin@example.com',
  WEB_APP_URL: 'https://script.google.com/macros/s/AKfycbwZO3UoovGBEnXi_JsepXCrAySgWJyU-RIAuczOaLKNn-6itAVMGa4-w0jcROnMey3F/exec',
  ALERT_THRESHOLD: {
    critical: 95,
    high: 90,
    medium: 80
  },
  TIMEZONE: 'Asia/Bangkok',
  LOCK_TIMEOUT_MS: 5000,
  MAX_LOCK_RETRIES: 3
};
