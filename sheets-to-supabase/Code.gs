const CONFIG = {
  SUPABASE_URL: PropertiesService
    .getScriptProperties()
    .getProperty("SUPABASE_URL"),

  SYNC_SECRET: PropertiesService
    .getScriptProperties()
    .getProperty("SYNC_SECRET"),
};


/**
 * Main sync function.
 *
 * Google Sheet format:
 *
 * A1: Date
 *
 * A2: Teacher
 * B2: Periods
 *
 * A3: Teacher
 * B3: Periods
 *
 * Example:
 *
 * A1 = 9/20/2026
 *
 * A2 = Bian
 * B2 = all
 *
 * A3 = Xu
 * B3 = 1-3, 7-8, 5, igs
 */
function syncAbsences() {
  if (!CONFIG.SUPABASE_URL) {
    throw new Error("SUPABASE_URL is not configured.");
  }

  if (!CONFIG.SYNC_SECRET) {
    throw new Error("SYNC_SECRET is not configured.");
  }

  const data = readSheet();

  sendToSupabase(data);
}


/**
 * Reads and parses the active Google Sheet.
 */
function readSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheet) {
    throw new Error(
      "No active spreadsheet found. Make sure this Apps Script is bound to the Google Sheet."
    );
  }

  const sheet = spreadsheet.getActiveSheet();

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
   *
   * We still return the date. The Supabase function can decide
   * how to handle an empty snapshot.
   */
  if (lastRow < 2) {
    return {
      date: date,
      absences: [],
    };
  }

  /*
   * Read columns A and B starting at row 2.
   */
  const values = sheet
    .getRange(2, 1, lastRow - 1, 2)
    .getValues();

  const absences = [];

  for (const row of values) {
    const teacher = String(row[0] ?? "").trim();
    const rawPeriods = String(row[1] ?? "").trim();

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
   * Sort teachers here so the payload is deterministic.
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
 * Parses a Google Sheets date value into:
 *
 * yyyy-MM-dd
 */
function parseSheetDate(value) {
  if (!(value instanceof Date) || isNaN(value.getTime())) {
    throw new Error(
      "Cell A1 must contain a valid date."
    );
  }

  return Utilities.formatDate(
    value,
    Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );
}


/**
 * Parses the raw periods string.
 *
 * Rules:
 *
 * all
 *   -> igs, 1, 2, 3, 4, 5, 6, 7, 8, 9
 *
 * igs
 *   -> igs
 *
 * 2
 *   -> 2
 *
 * 1-3
 *   -> 1, 2, 3
 *
 * Invalid values are ignored.
 *
 * Duplicate periods are automatically removed.
 */
function parsePeriods(rawValue) {
  const tokens = String(rawValue)
    .split(",")
    .map((token) => token.trim().toLowerCase())
    .filter((token) => token.length > 0);

  const periods = new Set();

  for (const token of tokens) {
    /*
     * "all" always wins.
     */
    if (token === "all") {
      return "igs, 1, 2, 3, 4, 5, 6, 7, 8, 9";
    }

    /*
     * IGS.
     */
    if (token === "igs") {
      periods.add("igs");
      continue;
    }

    /*
     * Single numerical period.
     *
     * Only 1 through 9 are valid.
     */
    if (/^\d+$/.test(token)) {
      const period = Number(token);

      if (period >= 1 && period <= 9) {
        periods.add(String(period));
      }

      continue;
    }

    /*
     * Numerical range.
     *
     * Examples:
     * 1-3
     * 2-9
     */
    const rangeMatch = token.match(/^(\d+)\s*-\s*(\d+)$/);

    if (rangeMatch) {
      const start = Number(rangeMatch[1]);
      const end = Number(rangeMatch[2]);

      /*
       * Invalid ranges are ignored.
       */
      if (
        start >= 1 &&
        start <= 9 &&
        end >= 1 &&
        end <= 9 &&
        start <= end
      ) {
        for (let period = start; period <= end; period++) {
          periods.add(String(period));
        }
      }

      continue;
    }

    /*
     * Anything else is intentionally ignored.
     */
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
  const endpoint =
    CONFIG.SUPABASE_URL +
    "/functions/v1/sync-absences";

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

  console.log(responseBody);
}


/**
 * Creates the five-minute trigger.
 *
 * Deletes existing sync triggers first so running this
 * function multiple times does not create duplicates.
 */
function createFiveMinuteTrigger() {
  const triggers = ScriptApp.getProjectTriggers();

  for (const trigger of triggers) {
    if (
      trigger.getHandlerFunction() === "syncAbsences"
    ) {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  ScriptApp
    .newTrigger("syncAbsences")
    .timeBased()
    .everyMinutes(5)
    .create();

  console.log(
    "Five-minute sync trigger created."
  );
}


/**
 * Manually test the entire sync process.
 */
function testSync() {
  syncAbsences();
}