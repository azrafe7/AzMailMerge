const DISABLE_TOASTS = false;
const LOG_TOASTS = true;

const TOASTER = {
  DEFAULT_TIMEOUT: 7,
  LOG: "",
  ON_EDIT: "✏️ ON_EDIT",
  DEBUG: "🐞 DEBUG",
  INFO: "ℹ️ INFO",
  WARN: "⚠️ WARNING",
  ERROR: "⛔ ERROR",
  
  log: (title, message, timeout) => {
    title = title ?? TOASTER.LOG;
    timeout = timeout ?? TOASTER.DEFAULT_TIMEOUT;
    if (LOG_TOASTS) Logger.log(`[${title}] ` + message);
    if (!DISABLE_TOASTS) SpreadsheetApp.getActive().toast(message, title, timeout);
  }
}

function testToaster() {
  TOASTER.log(TOASTER.ERROR, "message", 10);
}

