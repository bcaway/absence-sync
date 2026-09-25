/**
 * BCA Class Cancellation List to Google Sheets Sync
 *
 * Fetches the published BCA Class Cancellation Google Doc, extracts the
 * date and teacher cancellations table, formats according to specified rules,
 * and populates the Google Sheet.
 */

const CONFIG = {
  DOC_URL:
    PropertiesService.getScriptProperties().getProperty("DOC_URL")

  SPREADSHEET_ID:
    PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID"),

  SHEET_NAME:
    PropertiesService.getScriptProperties().getProperty("SHEET_NAME"),
};

const MONTH_MAP = {
  january: 1, jan: 1,
  february: 2, feb: 2,
  march: 3, mar: 3,
  april: 4, apr: 4,
  may: 5,
  june: 6, jun: 6,
  july: 7, jul: 7,
  august: 8, aug: 8,
  september: 9, sep: 9, sept: 9,
  october: 10, oct: 10,
  november: 11, nov: 11,
  december: 12, dec: 12,
};


/**
 * Main sync orchestrator.
 */
function syncDocToSheets() {
  const url = CONFIG.DOC_URL;
  if (!url) {
    throw new Error("DOC_URL is not configured.");
  }

  console.log("Fetching published doc from: " + url);
  const html = fetchPublishedDocHtml(url);

  console.log("Parsing document date...");
  const dateInfo = parseDocDate(html);
  console.log("Parsed date: " + dateInfo.formattedDate);

  console.log("Parsing cancellation table...");
  const rows = parseDocTable(html);
  console.log(`Parsed ${rows.length} teacher absence row(s).`);

  const sheet = getTargetSheet();
  writeToSheet(sheet, dateInfo, rows);

  console.log("Sync completed successfully.");
}


/**
 * Fetches published Google Doc HTML.
 * Handles both public published documents and domain-restricted published documents
 * by trying an authenticated request with the user's OAuth token and falling back
 * to a direct public fetch.
 */
function fetchPublishedDocHtml(url) {
  let lastError = null;

  // Attempt 1: Fetch with OAuth token (required if published within a Google Workspace domain)
  try {
    const token = ScriptApp.getOAuthToken();
    const responseWithAuth = UrlFetchApp.fetch(url, {
      headers: {
        Authorization: "Bearer " + token,
      },
      followRedirects: true,
      muteHttpExceptions: true,
    });

    const status = responseWithAuth.getResponseCode();
    const text = responseWithAuth.getContentText();

    if (status === 200 && !isLoginOrBlockedPage(text)) {
      return text;
    }
  } catch (e) {
    lastError = e;
  }

  // Attempt 2: Direct public fetch (without Authorization header)
  try {
    const response = UrlFetchApp.fetch(url, {
      followRedirects: true,
      muteHttpExceptions: true,
    });

    const status = response.getResponseCode();
    const text = response.getContentText();

    if (status === 200 && !isLoginOrBlockedPage(text)) {
      return text;
    }

    if (isLoginOrBlockedPage(text)) {
      throw new Error(
        "Received Google login redirect. Make sure the executing account has access to the published document."
      );
    }

    throw new Error(`HTTP ${status}: ${text.slice(0, 300)}`);
  } catch (e) {
    lastError = e;
  }

  throw new Error(
    "Failed to fetch published document: " + (lastError ? lastError.message : "Unknown error")
  );
}


/**
 * Checks if the fetched HTML is a Google login wall or access-blocked page.
 */
function isLoginOrBlockedPage(html) {
  if (!html) return true;
  return (
    html.includes("accounts.google.com/ServiceLogin") ||
    html.includes("Sign in to your Google Account") ||
    html.includes("class=\"request-storage-access\"") ||
    html.includes("class=\"document-root loading\"")
  );
}


/**
 * Parses the date from the document header.
 *
 * Expected format:
 * "BCA Class Cancellation List
 * {Month} {Day}, {YYYY}"
 *
 * Handles single digit days with and without leading zero:
 * "September 6, 2026" or "September 06, 2026"
 *
 * Returns:
 * {
 *   year: number,
 *   month: number,       // 1-12 (no leading zero)
 *   day: number,         // 1-31 (no leading zero)
 *   formattedDate: string, // "M/D/YYYY" (e.g. "9/6/2026")
 *   dateObj: Date        // Native Date object for Sheets
 * }
 */
function parseDocDate(html) {
  // Normalize HTML tags and whitespace
  const text = normalizeHtmlToText(html);

  // Match header and date
  const headerRegex =
    /BCA\s+Class\s+Cancellation\s+List[\s\S]*?([a-zA-Z]+)\s+(\d{1,2}),?\s+(\d{4})/i;

  let match = text.match(headerRegex);

  // Fallback: look for "Month Day, Year" anywhere before the table
  if (!match) {
    const fallbackRegex = /([a-zA-Z]+)\s+(\d{1,2}),?\s+(\d{4})/i;
    match = text.match(fallbackRegex);
  }

  if (!match) {
    throw new Error("Could not find cancellation list date in document.");
  }

  const rawMonth = match[1].toLowerCase();
  const rawDay = match[2];
  const rawYear = match[3];

  const month = MONTH_MAP[rawMonth];
  if (!month) {
    throw new Error(`Unrecognized month name: "${match[1]}"`);
  }

  const day = parseInt(rawDay, 10);
  const year = parseInt(rawYear, 10);

  if (isNaN(day) || day < 1 || day > 31) {
    throw new Error(`Invalid day value: "${rawDay}"`);
  }

  if (isNaN(year) || year < 2000 || year > 2100) {
    throw new Error(`Invalid year value: "${rawYear}"`);
  }

  // Format without leading zeroes: M/D/YYYY (e.g. 9/6/2026)
  const formattedDate = `${month}/${day}/${year}`;

  // Date object at noon (12:00:00) to protect against any timezone day shifts
  const dateObj = new Date(year, month - 1, day, 12, 0, 0);

  return {
    year: year,
    month: month,
    day: day,
    formattedDate: formattedDate,
    dateObj: dateObj,
  };
}


/**
 * Parses the 3-column cancellation table from the document HTML.
 *
 * Rules:
 * - Row 1 (index 0) is the header row: ignored.
 * - Column 3 (index 2) is the rightmost column: ignored.
 * - Column 1 (index 0): Teacher name string.
 * - Column 2 (index 1): Periods string, formatted according to rules.
 *
 * Returns array of [teacher, formattedPeriods].
 */
function parseDocTable(html) {
  const tableMatch = html.match(/<table[^>]*>([\s\S]*?)<\/table>/i);
  if (!tableMatch) {
    throw new Error("Could not find cancellation table in document HTML.");
  }

  const tableHtml = tableMatch[1];
  const rowRegex = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
  const cellRegex = /<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/gi;

  const rawRows = [];
  let rMatch;

  while ((rMatch = rowRegex.exec(tableHtml)) !== null) {
    const rowHtml = rMatch[1];
    const cells = [];
    let cMatch;

    while ((cMatch = cellRegex.exec(rowHtml)) !== null) {
      cells.push(cleanCellHtml(cMatch[1]));
    }

    if (cells.length > 0) {
      rawRows.push(cells);
    }
  }

  if (rawRows.length <= 1) {
    // Only header row or empty table
    return [];
  }

  const parsedRows = [];

  // Start at row index 1 (skipping header row 0)
  for (let i = 1; i < rawRows.length; i++) {
    const row = rawRows[i];
    const teacher = (row[0] || "").trim();
    const rawCol2 = row[1] || "";

    // Ignore completely empty rows
    if (!teacher && !rawCol2.trim()) {
      continue;
    }

    // If teacher exists, format periods according to rules
    const formattedPeriods = parsePeriodsCell(rawCol2);

    parsedRows.push([teacher, formattedPeriods]);
  }

  return parsedRows;
}


/**
 * Formats Column 2 (Periods) per requirements:
 *
 * 1. If word "all" (case-insensitive) appears -> "All"
 * 2. Otherwise:
 *    - Search for any pair of numbers that appear with a hyphen between them
 *      (e.g. 1-3, 1-8, 1-9). If multiple, separate by commas.
 *    - Then, if any non-hyphenated numbers appear, append separated by commas.
 *    - Then, if the text "igs" (case-insensitive) appears, append "igs"
 *      separated by comma.
 */
function parsePeriodsCell(rawText) {
  if (!rawText) return "";

  // Normalize en-dashes, em-dashes, and whitespace
  const text = rawText
    .replace(/[\u2013\u2014]/g, "-")
    .replace(/\u00A0/g, " ")
    .trim();

  // Rule 1: If the word "all" in any case appears
  if (/\ball\b/i.test(text)) {
    return "All";
  }

  const parts = [];

  // Rule 2: Pair of numbers with a hyphen between them (e.g. 1-3, 1-8, 1-9)
  const hyphenRegex = /(\d+)\s*-\s*(\d+)/g;
  let match;
  while ((match = hyphenRegex.exec(text)) !== null) {
    parts.push(`${match[1]}-${match[2]}`);
  }

  // Rule 3: Non-hyphenated numbers
  // Mask out the hyphenated pairs so their digits aren't treated as standalone numbers
  const maskedText = text.replace(/(\d+)\s*-\s*(\d+)/g, " ");
  const singleNumbers = maskedText.match(/\b\d+\b/g);
  if (singleNumbers) {
    for (const num of singleNumbers) {
      parts.push(num);
    }
  }

  // Rule 4: "igs"
  if (/\bigs\b/i.test(text)) {
    parts.push("igs");
  }

  return parts.join(", ");
}


/**
 * Cleans individual HTML table cell contents into plain text.
 */
function cleanCellHtml(cellHtml) {
  if (!cellHtml) return "";

  return cellHtml
    .replace(/<br\s*[\/]?>/gi, " ")
    .replace(/<\/p>/gi, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/&#160;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&#39;|&apos;/gi, "'")
    .replace(/&quot;/gi, '"')
    .replace(/&#8211;|&ndash;/gi, "-")
    .replace(/&#8212;|&mdash;/gi, "-")
    .replace(/[\u2013\u2014]/g, "-")
    .replace(/\u00A0/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}


/**
 * Normalizes full document HTML into plain text for header and date parsing.
 */
function normalizeHtmlToText(html) {
  return html
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, " ")
    .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, " ")
    .replace(/<br\s*[\/]?>/gi, " ")
    .replace(/<\/p>/gi, "\n")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/&#160;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&#8211;|&ndash;/gi, "-")
    .replace(/[\u2013\u2014]/g, "-")
    .replace(/\u00A0/g, " ")
    .replace(/[ \t]+/g, " ")
    .trim();
}


/**
 * Retrieves the target Google Sheet.
 */
function getTargetSheet() {
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // If running as a standalone script with SPREADSHEET_ID configured
  if (!spreadsheet && CONFIG.SPREADSHEET_ID) {
    spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  }

  if (!spreadsheet) {
    throw new Error(
      "No active spreadsheet found. Make sure this Apps Script is bound to the Google Sheet, " +
      "or configure SPREADSHEET_ID in Script Properties."
    );
  }

  if (CONFIG.SHEET_NAME) {
    const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) {
      throw new Error(`Sheet tab "${CONFIG.SHEET_NAME}" was not found.`);
    }
    return sheet;
  }

  return spreadsheet.getActiveSheet() || spreadsheet.getSheets()[0];
}


/**
 * Writes the date and teacher cancellations to the target sheet.
 *
 * Structure:
 * A1: Date (formatted as M/D/YYYY without leading zeroes, e.g. 9/6/2026)
 * A2:B{N+1}: Teacher and formatted periods
 */
function writeToSheet(sheet, dateInfo, rows) {
  // 1. Set A1 date with Date object and format as "m/d/yyyy"
  const cellA1 = sheet.getRange("A1");
  cellA1.setValue(dateInfo.dateObj);
  cellA1.setNumberFormat("m/d/yyyy");

  // 2. Clear any old data in columns A and B from row 2 downwards
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, 2).clearContent();
  }

  // 3. Write rows starting at A2
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  }

  SpreadsheetApp.flush();
}


/**
 * Creates a recurring 5-minute time-driven trigger.
 * Removes existing triggers for syncDocToSheets first to avoid duplicates.
 */
function createFiveMinuteTrigger() {
  createSyncTrigger(5);
}


/**
 * Creates a recurring 1-minute time-driven trigger.
 */
function createOneMinuteTrigger() {
  createSyncTrigger(1);
}


/**
 * Generic trigger creator for periodic sync.
 *
 * @param {number} minutes Trigger interval in minutes (1, 5, 10, 15, or 30).
 */
function createSyncTrigger(minutes) {
  deleteTriggers();

  ScriptApp.newTrigger("syncDocToSheets")
    .timeBased()
    .everyMinutes(minutes || 5)
    .create();

  console.log(`Sync trigger created: runs every ${minutes || 5} minute(s).`);
}


/**
 * Deletes all existing triggers for syncDocToSheets.
 */
function deleteTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === "syncDocToSheets") {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  console.log("Existing sync triggers removed.");
}


/**
 * Manual test function for direct execution in Apps Script editor.
 */
function testSync() {
  syncDocToSheets();
}
