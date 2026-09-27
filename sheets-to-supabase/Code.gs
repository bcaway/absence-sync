/**
 * BCA Google Sheets to Supabase Sync (Event-Driven)
 *
 * Reads staged absence data from Google Sheets, normalizes the payload,
 * detects changes using content fingerprinting (MD5 hash), and synchronizes
 * directly with the Supabase sync-absences Edge Function.
 *
 * Automatically triggered on every change in the Google Sheet using an
 * installable onChange trigger (zero polling latency, zero redundant API calls).
 */

const CONFIG = {
  get SUPABASE_URL() {
    return PropertiesService.getScriptProperties().getProperty("SUPABASE_URL");
  },

  get SYNC_SECRET() {
    return PropertiesService.getScriptProperties().getProperty("SYNC_SECRET");
  },

  get SHEET_NAME() {
    return PropertiesService.getScriptProperties().getProperty("SHEET_NAME");
  },
};


/**
 * Event handler executed automatically whenever a change occurs in the Google Sheet.
 *
 * Configured via createSheetChangeTrigger().
 *
 * @param {Object} [e] Google Apps Script change event object.
 */
function handleChange(e) {
  const changeType = e && e.changeType ? e.changeType : "CHANGE";
  console.log(`Detected spreadsheet change event: ${changeType}`);
  syncAbsences({ changeType: changeType });
}


/**
 * Main sync function.
 * Reads the sheet, checks if the data changed since last sync, and pushes to Supabase.
 *
 * Google Sheet format:
 * A1: Date (e.g. 9/20/2026 or Date object)
 * A2: Teacher (e.g. Bian)
 * B2: Periods (e.g. all, or 1-3, 7-8, 5, igs)
 *
 * @param {Object} [options]
 * @param {boolean} [options.force=false] If true, bypasses the hash check and forces sync.
 * @param {string} [options.changeType] Description of the trigger event type.
 */
function syncAbsences(options) {
  options = options || {};
  const force = Boolean(options.force);

  if (!CONFIG.SUPABASE_URL) {
    throw new Error(
      "SUPABASE_URL is not configured. Please set 'SUPABASE_URL' in Project Settings > Script Properties."
    );
  }

  if (!CONFIG.SYNC_SECRET) {
    throw new Error(
      "SYNC_SECRET is not configured. Please set 'SYNC_SECRET' in Project Settings > Script Properties."
    );
  }

  // Prevent concurrent executions if multiple rapid edits occur
  const lock = LockService.getScriptLock();
  const hasLock = lock.tryLock(15000);
  if (!hasLock) {
    console.warn("Could not acquire script lock: another sync is already in progress.");
    return { skipped: true, reason: "locked" };
  }

  try {
    const data = readSheet();
    const currentHash = computeDataHash(data);
    const props = PropertiesService.getScriptProperties();
    const lastHash = props.getProperty("LAST_SYNCED_HASH");

    if (!force && lastHash === currentHash) {
      console.log(
        `No change detected in absences table since last sync (hash: ${currentHash}). Skipping Supabase sync.`
      );
      return { skipped: true, hash: currentHash };
    }

    console.log(
      `Detected change in absences table (old hash: ${lastHash || "none"}, new hash: ${currentHash}). ` +
      `Syncing ${data.absences.length} record(s) for ${data.date} to Supabase...`
    );

    const response = sendToSupabase(data);

    props.setProperty("LAST_SYNCED_HASH", currentHash);
    props.setProperty("LAST_SYNCED_AT", new Date().toISOString());

    console.log(`Supabase sync completed successfully (hash: ${currentHash}).`);
    return { skipped: false, hash: currentHash, response: response };
  } finally {
    lock.releaseLock();
  }
}


/**
 * Computes a deterministic MD5 hash for the sheet absence data snapshot.
 * Used to avoid redundant HTTP requests when the sheet contents have not changed.
 *
 * @param {Object} data Parsed sheet data.
 * @returns {string} Base64-encoded MD5 hash string.
 */
function computeDataHash(data) {
  const content = JSON.stringify(data);
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5,
    content,
    Utilities.Charset.UTF_8
  );
  return Utilities.base64Encode(digest);
}


/**
 * Gets the active or configured target sheet.
 */
function getTargetSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheet) {
    throw new Error(
      "No active spreadsheet found. Make sure this Apps Script is bound to the Google Sheet (Extensions > Apps Script)."
    );
  }

  const configuredName = CONFIG.SHEET_NAME;
  if (configuredName) {
    const found = spreadsheet.getSheetByName(configuredName);
    if (found) return found;
    console.warn(`Configured SHEET_NAME "${configuredName}" not found; falling back to active sheet.`);
  }

  return spreadsheet.getActiveSheet() || spreadsheet.getSheets()[0];
}


/**
 * Reads and parses the Google Sheet.
 */
function readSheet() {
  const sheet = getTargetSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < 1) {
    throw new Error("The Google Sheet is empty.");
  }

  /*
   * A1 must contain the date.
   */
  const dateValue = sheet.getRange("A1").getValue();
  const date = parseSheetDate(dateValue);

  /*
   * There are no teacher rows.
   * Return empty absences snapshot.
   */
  if (lastRow < 2) {
    return {
      date: date,
      absences: [],
    };
  }

  /*
   * Read columns A and B starting at row 2.
   * We read both displayValues (for formatted text strings) and rawValues (to detect and unwrap coerced dates).
   */
  const range = sheet.getRange(2, 1, lastRow - 1, 2);
  const displayValues = range.getDisplayValues();
  const rawValues = range.getValues();

  const absences = [];

  for (let i = 0; i < displayValues.length; i++) {
    const teacher = String(displayValues[i][0] ?? "").trim();
    let rawPeriods = String(displayValues[i][1] ?? "").trim();

    // If Google Sheets coerced "1-4" into a Date (e.g. Jan 4), recover the period range
    const rawVal = rawValues[i][1];
    if (rawVal instanceof Date && (!rawPeriods || /^\d{1,2}\/\d{1,2}/.test(rawPeriods))) {
      rawPeriods = `${rawVal.getMonth() + 1}-${rawVal.getDate()}`;
    }

    /*
     * Ignore completely empty rows.
     */
    if (!teacher && !rawPeriods) {
      continue;
    }

    /*
     * A teacher without periods is invalid and therefore ignored.
     */
    if (!teacher || !rawPeriods) {
      continue;
    }

    const periodsImpacted = parsePeriods(rawPeriods);

    /*
     * If nothing in the period field was parsable,
     * ignore the row.
     */
    if (!periodsImpacted) {
      continue;
    }

    absences.push({
      teacher: teacher,
      periods_impacted: periodsImpacted,
    });
  }

  /*
   * Sort teachers here so the payload is deterministic for hashing.
   */
  absences.sort((a, b) =>
    a.teacher.localeCompare(b.teacher)
  );

  return {
    date: date,
    absences: absences,
  };
}


/**
 * Parses a Google Sheets date value (Date object or text string) into:
 * yyyy-MM-dd
 */
function parseSheetDate(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );
  }

  if (typeof value === "string" && value.trim()) {
    const trimmed = value.trim();

    // Standard yyyy-MM-dd
    if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) {
      return trimmed;
    }

    // M/D/YYYY or MM/DD/YYYY
    const slashMatch = trimmed.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (slashMatch) {
      const month = slashMatch[1].padStart(2, "0");
      const day = slashMatch[2].padStart(2, "0");
      const year = slashMatch[3];
      return `${year}-${month}-${day}`;
    }

    // Fallback: new Date(trimmed)
    const parsed = new Date(trimmed);
    if (!isNaN(parsed.getTime())) {
      return Utilities.formatDate(
        parsed,
        Session.getScriptTimeZone(),
        "yyyy-MM-dd"
      );
    }
  }

  throw new Error(
    "Cell A1 must contain a valid date (e.g. 9/20/2026 or Date object). Received: " + value
  );
}


/**
 * Parses the raw periods string.
 *
 * Rules:
 * - "all" / "all day" -> "igs, 1, 2, 3, 4, 5, 6, 7, 8, 9"
 * - "igs" -> "igs"
 * - 2 -> "2"
 * - 1-3 -> "1, 2, 3"
 * - Duplicate periods are automatically removed.
 */
function parsePeriods(rawValue) {
  if (!rawValue) return "";

  // Normalize en-dashes, em-dashes, non-breaking spaces, and whitespace
  let text = String(rawValue)
    .replace(/[\u2013\u2014]/g, "-")
    .replace(/\u00A0/g, " ")
    .trim();

  // "all" or "all day" -> full day
  if (/\ball\b/i.test(text)) {
    return "igs, 1, 2, 3, 4, 5, 6, 7, 8, 9";
  }

  // Pre-normalize common connectors and words:
  // e.g. "Periods 1-4" -> "1-4", "7 & 8" -> "7, 8", "1 to 4" -> "1-4", "7 and 8" -> "7, 8"
  text = text
    .replace(/\b(?:through|thru|to)\b/gi, "-")
    .replace(/&|\band\b|\+|\/|;/gi, ",")
    .replace(/\b(?:periods?|mods?|p\.?)\b/gi, " ")
    .replace(/\s*-\s*/g, "-");

  const tokens = text
    .split(",")
    .map((token) => token.trim().toLowerCase())
    .filter((token) => token.length > 0);

  const periods = new Set();

  for (const token of tokens) {
    // IGS
    if (token === "igs") {
      periods.add("igs");
      continue;
    }

    // Numerical range: e.g. 1-4, 7-9
    const rangeMatch = token.match(/^([1-9])-([1-9])$/);
    if (rangeMatch) {
      const start = Number(rangeMatch[1]);
      const end = Number(rangeMatch[2]);
      if (start <= end) {
        for (let period = start; period <= end; period++) {
          periods.add(String(period));
        }
      }
      continue;
    }

    // Single numerical period (1-9)
    if (/^[1-9]$/.test(token)) {
      periods.add(token);
      continue;
    }

    // Fallback: extract any standalone digits 1-9 within the token
    const digits = token.match(/\b[1-9]\b/g);
    if (digits) {
      for (const d of digits) {
        periods.add(d);
      }
    }
  }

  if (periods.size === 0) {
    return "";
  }

  /*
   * IGS must always come first.
   * Numerical periods follow in ascending order.
   */
  const sortedPeriods = [];

  if (periods.has("igs")) {
    sortedPeriods.push("igs");
  }

  for (let period = 1; period <= 9; period++) {
    if (periods.has(String(period))) {
      sortedPeriods.push(String(period));
    }
  }

  return sortedPeriods.join(", ");
}


/**
 * Sends the normalized snapshot to Supabase.
 */
function sendToSupabase(data) {
  const baseUrl = (CONFIG.SUPABASE_URL || "").trim().replace(/\/+$/, "");
  const endpoint = baseUrl + "/functions/v1/sync-absences";

  const response = UrlFetchApp.fetch(endpoint, {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + CONFIG.SYNC_SECRET,
    },
    payload: JSON.stringify(data),
    muteHttpExceptions: true,
  });

  const status = response.getResponseCode();
  const responseBody = response.getContentText();

  if (status < 200 || status >= 300) {
    throw new Error(
      "Supabase sync failed (" +
      status +
      "): " +
      responseBody
    );
  }

  console.log("Supabase response:", responseBody);
  return responseBody;
}


/**
 * Creates the installable onChange trigger.
 *
 * Removes any existing triggers first so running this function multiple
 * times does not create duplicates.
 */
function createSheetChangeTrigger() {
  deleteTriggers();

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error(
      "No active spreadsheet found. Open the Google Sheet and open Apps Script from Extensions > Apps Script."
    );
  }

  ScriptApp.newTrigger("handleChange")
    .forSpreadsheet(spreadsheet)
    .onChange()
    .create();

  console.log(
    "Installable sheet change trigger created successfully! " +
    "Every edit, addition, or deletion in the Google Sheet will now trigger sync to Supabase."
  );
}


/**
 * Removes existing triggers associated with this sync script.
 */
function deleteTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  let count = 0;

  for (const trigger of triggers) {
    const handler = trigger.getHandlerFunction();
    if (
      handler === "handleChange" ||
      handler === "syncAbsences" ||
      handler === "onSheetChange"
    ) {
      ScriptApp.deleteTrigger(trigger);
      count++;
    }
  }

  console.log(`Removed ${count} existing trigger(s).`);
}


/**
 * Lists all active project triggers for inspection.
 */
function listTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  console.log(`Active triggers for this project (${triggers.length}):`);
  for (const trigger of triggers) {
    console.log(
      `- Handler: ${trigger.getHandlerFunction()} | Type: ${trigger.getEventType()} | ID: ${trigger.getUniqueId()}`
    );
  }
}


/**
 * Forces a manual sync to Supabase, bypassing the change hash check.
 */
function forceSyncAbsences() {
  console.log("Starting forced manual sync to Supabase...");
  return syncAbsences({ force: true });
}


/**
 * Manually test the entire sync process.
 */
function testSync() {
  return forceSyncAbsences();
}


/**
 * @deprecated Legacy function kept for backward compatibility.
 * Replaced by createSheetChangeTrigger().
 */
function createFiveMinuteTrigger() {
  console.warn("createFiveMinuteTrigger is deprecated. Creating real-time sheet change trigger instead...");
  createSheetChangeTrigger();
}