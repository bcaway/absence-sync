/**
 * BCA Class Cancellation List to Google Sheets Sync
 *
 * Fetches the published BCA Class Cancellation Google Doc, extracts the
 * date and teacher cancellations table, formats according to specified rules,
 * and populates the Google Sheet.
 */

const DOC_CONFIG = {
  get DOC_URL() {
    return (
      PropertiesService.getScriptProperties().getProperty("DOC_URL")
    );
  },

  // Optional: If the original Google Doc ID is known and accessible
  get DOC_ID() {
    return PropertiesService.getScriptProperties().getProperty("DOC_ID");
  },

  // Optional: Browser session cookie if published doc is domain-restricted
  get DOC_COOKIE() {
    return PropertiesService.getScriptProperties().getProperty("DOC_COOKIE");
  },
};

const DOC_MONTH_MAP = {
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
 * Main sync orchestrator for doc-to-sheets.
 */
function syncDocToSheets() {
  let dateInfo;
  let rows;

  // Path A: Direct DocumentApp access if DOC_ID is configured
  if (DOC_CONFIG.DOC_ID) {
    console.log("Using DocumentApp with DOC_ID: " + DOC_CONFIG.DOC_ID);
    const docData = readFromDocumentApp(DOC_CONFIG.DOC_ID);
    dateInfo = docData.dateInfo;
    rows = docData.rows;
  } else {
    // Path B: Fetch published web document
    const url = DOC_CONFIG.DOC_URL;
    if (!url) {
      throw new Error("DOC_URL is not configured.");
    }

    console.log("Fetching published doc from: " + url);
    const html = fetchPublishedDocHtml(url);

    console.log("Parsing document date...");
    dateInfo = parseDocDate(html);
    console.log("Parsed date: " + dateInfo.formattedDate);

    console.log("Parsing cancellation table...");
    rows = parseDocTable(html);
    console.log(`Parsed ${rows.length} teacher absence row(s).`);
  }

  const sheet = getDocTargetSheet();
  writeDocDataToSheet(sheet, dateInfo, rows);

  console.log("Doc to Sheets sync completed successfully.");
}


/**
 * Reads directly from Google Docs if DOC_ID is provided.
 */
function readFromDocumentApp(docId) {
  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();
  const text = body.getText();
  const dateInfo = parseDocDate(text);

  const tables = body.getTables();
  if (tables.length === 0) {
    throw new Error("No table found in Google Document.");
  }

  const table = tables[0];
  const numRows = table.getNumRows();
  const rows = [];

  for (let i = 1; i < numRows; i++) {
    const tableRow = table.getRow(i);
    const teacher = tableRow.getCell(0).getText().trim();
    const rawCol2 = tableRow.getCell(1).getText().trim();

    if (!teacher && !rawCol2) {
      continue;
    }

    const formattedPeriods = parsePeriodsCell(rawCol2);
    rows.push([teacher, formattedPeriods]);
  }

  return { dateInfo, rows };
}


/**
 * Fetches published Google Doc HTML.
 */
function fetchPublishedDocHtml(url) {
  // Strategy 0: If DOC_COOKIE is configured in Script Properties
  if (DOC_CONFIG.DOC_COOKIE) {
    try {
      console.log("Attempting fetch with provided DOC_COOKIE...");
      const cleanCookie = DOC_CONFIG.DOC_COOKIE.replace(/^Cookie:\s*/i, "").trim();
      const response = UrlFetchApp.fetch(url, {
        headers: {
          Cookie: cleanCookie,
          "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
          Accept: "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
          "Accept-Language": "en-US,en;q=0.9",
        },
        followRedirects: true,
        muteHttpExceptions: true,
      });
      const status = response.getResponseCode();
      const text = response.getContentText();
      console.log(`Cookie fetch HTTP ${status}, length: ${text.length}`);

      if (hasDocCancellationContent(text)) {
        console.log("Cancellation content detected with session cookie.");
        return text;
      } else {
        console.warn(`Cookie fetch did not return cancellation content. HTTP ${status}, preview: ${text.slice(0, 150)}`);
      }
    } catch (e) {
      console.warn("Cookie fetch exception: " + e.message);
    }
  }

  // Strategy 1: Direct public fetch
  try {
    console.log("Attempting direct fetch of published doc URL...");
    const response = UrlFetchApp.fetch(url, {
      followRedirects: true,
      muteHttpExceptions: true,
    });
    const status = response.getResponseCode();
    const text = response.getContentText();
    console.log(`Direct fetch HTTP ${status}, length: ${text.length}`);

    if (hasDocCancellationContent(text)) {
      console.log("Cancellation content detected in direct fetch response.");
      return text;
    }
  } catch (e) {
    console.warn("Direct fetch exception: " + e.message);
  }

  // Strategy 2: Authenticated fetch with OAuth Bearer token
  try {
    console.log("Attempting fetch with OAuth bearer token...");
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
    console.log(`OAuth fetch HTTP ${status}, length: ${text.length}`);

    if (hasDocCancellationContent(text)) {
      console.log("Cancellation content detected in OAuth fetch response.");
      return text;
    }
  } catch (e) {
    console.warn("OAuth fetch exception: " + e.message);
  }

  // Strategy 3: Try with ?embedded=true
  try {
    console.log("Attempting fetch with ?embedded=true parameter...");
    const embedUrl = url.includes("?") ? url + "&embedded=true" : url + "?embedded=true";
    const embedResponse = UrlFetchApp.fetch(embedUrl, {
      followRedirects: true,
      muteHttpExceptions: true,
    });
    const status = embedResponse.getResponseCode();
    const text = embedResponse.getContentText();
    console.log(`Embedded fetch HTTP ${status}, length: ${text.length}`);

    if (hasDocCancellationContent(text)) {
      console.log("Cancellation content detected in embedded fetch response.");
      return text;
    }
  } catch (e) {
    console.warn("Embedded fetch exception: " + e.message);
  }

  // Strategy 4: Fallback to searching Google Drive for the document
  try {
    console.log("Searching user's Google Drive for cancellation document...");
    const driveDocHtml = tryFetchDocFromDrive();
    if (driveDocHtml) {
      console.log("Cancellation document found and exported from Google Drive.");
      return driveDocHtml;
    }
  } catch (e) {
    console.warn("Drive search exception: " + e.message);
  }

  throw new Error(
    "Failed to fetch cancellation list content from: " + url + "\n\n" +
    "Google returned HTTP 401/403 because the document was published with 'Require viewers to sign in with their domain account' enabled.\n" +
    "Server-side scripts cannot perform interactive Google logins without session cookies.\n\n" +
    "To resolve this, choose one of the following options:\n" +
    "1. (Recommended) In the Google Doc, go to File > Share > Publish to web, expand 'Published content & settings', and uncheck 'Require viewers to sign in'.\n" +
    "2. If you have the standard Google Doc ID, set DOC_ID in Script Properties.\n" +
    "3. Set DOC_COOKIE in Script Properties with your browser session cookie from viewing the page."
  );
}


/**
 * Checks whether the HTML contains the cancellation document content.
 */
function hasDocCancellationContent(html) {
  if (!html) return false;
  return (
    html.includes("BCA Class Cancellation List") ||
    (html.includes("Cancellation List") && /<table[^>]*>/i.test(html)) ||
    (html.includes("Cancellation") && /<table[^>]*>/i.test(html)) ||
    (/<table[^>]*>/i.test(html) && /Teacher/i.test(html))
  );
}


/**
 * Fallback to search Drive for files matching "Cancellation" and export HTML.
 */
function tryFetchDocFromDrive() {
  if (typeof DriveApp === "undefined") {
    return null;
  }

  const query = "title contains 'Cancellation' and trashed = false";
  const files = DriveApp.searchFiles(query);

  while (files.hasNext()) {
    const file = files.next();
    const fileId = file.getId();
    console.log(`Found Drive file candidate: "${file.getName()}" (${fileId})`);

    try {
      const exportUrl = `https://docs.google.com/document/d/${fileId}/export?format=html`;
      const token = ScriptApp.getOAuthToken();
      const resp = UrlFetchApp.fetch(exportUrl, {
        headers: {
          Authorization: "Bearer " + token,
        },
        muteHttpExceptions: true,
      });

      if (resp.getResponseCode() === 200) {
        const content = resp.getContentText();
        if (hasDocCancellationContent(content)) {
          return content;
        }
      }
    } catch (err) {
      console.warn(`Could not export file ${fileId}: ${err.message}`);
    }
  }

  return null;
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
function parseDocDate(htmlOrText) {
  const text = normalizeHtmlToText(htmlOrText);

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

  const month = DOC_MONTH_MAP[rawMonth];
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
 * Retrieves the target Google Sheet for doc-to-sheets sync.
 */
function getDocTargetSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheet) {
    throw new Error(
      "No active spreadsheet found. Make sure this Apps Script is bound to the Google Sheet."
    );
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
function writeDocDataToSheet(sheet, dateInfo, rows) {
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
 * Creates a recurring 5-minute time-driven trigger for doc-to-sheets.
 * Removes existing triggers for syncDocToSheets first to avoid duplicates.
 */
function createDocToSheetsFiveMinuteTrigger() {
  createDocToSheetsTrigger(5);
}


/**
 * Creates a recurring 1-minute time-driven trigger for doc-to-sheets.
 */
function createDocToSheetsOneMinuteTrigger() {
  createDocToSheetsTrigger(1);
}


/**
 * Generic trigger creator for periodic doc-to-sheets sync.
 *
 * @param {number} minutes Trigger interval in minutes (1, 5, 10, 15, or 30).
 */
function createDocToSheetsTrigger(minutes) {
  deleteDocToSheetsTriggers();

  ScriptApp.newTrigger("syncDocToSheets")
    .timeBased()
    .everyMinutes(minutes || 5)
    .create();

  console.log(`Doc-to-sheets sync trigger created: runs every ${minutes || 5} minute(s).`);
}


/**
 * Deletes all existing triggers for syncDocToSheets.
 */
function deleteDocToSheetsTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === "syncDocToSheets") {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  console.log("Existing doc-to-sheets sync triggers removed.");
}


/**
 * Manual test function for doc-to-sheets execution in Apps Script editor.
 */
function testDocToSheetsSync() {
  syncDocToSheets();
}


/**
 * Web App HTTP POST handler.
 * Receives session cookie updates from the Chrome Cookie Tool extension.
 *
 * Payload format (JSON or URL-encoded):
 * {
 *   "cookie": "SID=...; HSID=...",
 *   "secret": "optional-secret-key",
 *   "triggerSync": false
 * }
 */
function doPost(e) {
  try {
    let payload = {};

    if (e && e.postData && e.postData.contents) {
      try {
        payload = JSON.parse(e.postData.contents);
      } catch (jsonErr) {
        // Fallback for form-encoded or plain string
        if (e.postData.contents.indexOf("cookie=") !== -1) {
          payload = e.parameter || {};
        } else {
          payload = { cookie: e.postData.contents };
        }
      }
    } else if (e && e.parameter) {
      payload = e.parameter;
    }

    const scriptProps = PropertiesService.getScriptProperties();
    const expectedSecret = scriptProps.getProperty("SYNC_SECRET");

    if (expectedSecret && payload.secret !== expectedSecret) {
      return ContentService.createTextOutput(
        JSON.stringify({
          status: "error",
          message: "Unauthorized: Invalid or missing sync secret.",
        })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    const cookie = (payload.cookie || "").trim();
    if (!cookie) {
      return ContentService.createTextOutput(
        JSON.stringify({
          status: "error",
          message: "Missing 'cookie' field in request body.",
        })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    // Save cookie and timestamp
    const nowIso = new Date().toISOString();
    scriptProps.setProperty("DOC_COOKIE", cookie);
    scriptProps.setProperty("DOC_COOKIE_UPDATED_AT", nowIso);

    console.log(`Updated DOC_COOKIE via Web App at ${nowIso}, length: ${cookie.length}`);

    let syncMessage = "Cookie successfully updated.";
    if (payload.triggerSync === true || payload.triggerSync === "true") {
      try {
        syncDocToSheets();
        syncMessage = "Cookie updated and sync executed successfully.";
      } catch (syncErr) {
        console.error("Immediate sync error after cookie update: " + syncErr.message);
        syncMessage = "Cookie updated, but sync execution failed: " + syncErr.message;
      }
    }

    return ContentService.createTextOutput(
      JSON.stringify({
        status: "success",
        message: syncMessage,
        updatedAt: nowIso,
      })
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    console.error("doPost error: " + err.toString());
    return ContentService.createTextOutput(
      JSON.stringify({
        status: "error",
        message: err.toString(),
      })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}


/**
 * Web App HTTP GET handler.
 * Provides a health check endpoint for testing the Web App deployment.
 */
function doGet(e) {
  const scriptProps = PropertiesService.getScriptProperties();
  const hasCookie = !!scriptProps.getProperty("DOC_COOKIE");
  const updatedAt = scriptProps.getProperty("DOC_COOKIE_UPDATED_AT") || null;
  const hasSecret = !!scriptProps.getProperty("SYNC_SECRET");

  return ContentService.createTextOutput(
    JSON.stringify({
      status: "ok",
      service: "BCA Absence Sync Web App",
      hasActiveCookie: hasCookie,
      cookieUpdatedAt: updatedAt,
      requiresSecret: hasSecret,
      serverTime: new Date().toISOString(),
    })
  ).setMimeType(ContentService.MimeType.JSON);
}

