/**
 * BCA Class Cancellation List to Google Sheets Sync
 *
 * Fetches the domain-restricted published BCA Class Cancellation Google Doc
 * using active session credentials synchronized by the BCA Absence Sync Chrome Extension.
 * Formats the cancellation table and stages data in Google Sheets for downstream sync.
 */

const DOC_CONFIG = {
  DOC_URL:
    PropertiesService.getScriptProperties().getProperty("DOC_URL") ||
    "https://docs.google.com/document/d/e/2PACX-1vRkhySmwAiTtY88tcshckpV4F0vRrULccaGrYl_Sf2ubWpyyXA4l8c-KAOuMzSwFe-qyAQhLqXzVsbA/pub",

  AUTH_USER:
    PropertiesService.getScriptProperties().getProperty("AUTH_USER") || "kabsek30@bergen.org",

  // Cookie synchronized automatically from the Chrome Extension
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
 * Fetches published doc HTML with the extension's session cookie, parses data, and writes to Sheets.
 */
function syncDocToSheets() {
  const url = DOC_CONFIG.DOC_URL;
  console.log("Starting BCA absence sync from: " + url);

  const html = fetchPublishedDocHtml(url);

  console.log("Parsing document date...");
  const dateInfo = parseDocDate(html);
  console.log("Parsed date: " + dateInfo.formattedDate);

  console.log("Parsing cancellation table...");
  const rows = parseDocTable(html);
  console.log(`Parsed ${rows.length} teacher absence row(s).`);

  const sheet = getDocTargetSheet();
  writeDocDataToSheet(sheet, dateInfo, rows);

  console.log("Doc to Sheets sync completed successfully.");
}


/**
 * Fetches published Google Doc HTML using the session cookie provided by the Chrome extension.
 *
 * @param {string} url Published doc URL.
 * @return {string} Document HTML content.
 */
function fetchPublishedDocHtml(url) {
  const cookie = DOC_CONFIG.DOC_COOKIE;

  if (!cookie || !cookie.trim()) {
    // Check if we have valid HTML synchronized directly from the Chrome extension
    const cachedHtml = PropertiesService.getScriptProperties().getProperty("LAST_VALID_HTML");
    if (cachedHtml && hasDocCancellationContent(cachedHtml)) {
      console.log("No DOC_COOKIE found, but using latest HTML synchronized directly from Chrome extension.");
      return cachedHtml;
    }

    throw new Error(
      "Missing DOC_COOKIE in Script Properties.\n\n" +
      "The BCA Class Cancellation document requires domain authentication.\n" +
      "Please open the BCA Absence Sync Chrome extension and click 'Sync Cookie Now' to connect your session."
    );
  }

  const cleanCookie = cookie.replace(/^Cookie:\s*/i, "").trim();

  // Filter out known problematic cookies that cause Google to redirect to Account Chooser
  const disallowedCookies = [
    "ACCOUNT_CHOOSER",
    "PLAY_ACTIVE_ACCOUNT",
    "GG_ACTIVE_ACCOUNT",
    "GG_XSRF",
    "GMAIL_AT",
    "__Host-GAPS",
    "LSID",
    "__Host-1PLSID",
    "__Host-3PLSID",
    "LSOLH",
    "SNID",
    "SMSV",
    "COMPASS",
  ];

  const filteredCookies = cleanCookie
    .split(";")
    .map(function(s) { return s.trim(); })
    .filter(function(cookiePair) {
      const name = cookiePair.split("=")[0].trim();
      if (!name) return false;
      if (disallowedCookies.indexOf(name) !== -1) return false;
      if (name.indexOf("__Host-GMAIL") === 0 || name.indexOf("GMAIL") === 0) return false;
      return true;
    });

  const cookieHeader = filteredCookies.join("; ");

  // Log cookie diagnostic summary
  const cookieNames = filteredCookies.map(function(s) { return s.split("=")[0]; });
  console.log(`Using ${cookieNames.length} session cookie(s): ${cookieNames.join(", ")}`);
  console.log(`Total cookie payload length: ${cookieHeader.length} characters`);

  // Ensure published doc URL includes not_in_iframe=true and authuser parameter
  let fetchUrl = url;
  if (fetchUrl.indexOf("not_in_iframe=true") === -1) {
    fetchUrl += (fetchUrl.indexOf("?") === -1 ? "?" : "&") + "not_in_iframe=true";
  }
  if (DOC_CONFIG.AUTH_USER && fetchUrl.indexOf("authuser=") === -1) {
    fetchUrl += "&authuser=" + encodeURIComponent(DOC_CONFIG.AUTH_USER);
  }

  console.log("Fetching cancellation doc from: " + fetchUrl);

  const response = UrlFetchApp.fetch(fetchUrl, {
    headers: {
      Cookie: cookieHeader,
      "User-Agent":
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
      Accept:
        "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
      "Accept-Language": "en-US,en;q=0.9",
    },
    followRedirects: true,
    muteHttpExceptions: true,
  });

  const status = response.getResponseCode();
  const text = response.getContentText();
  console.log(`Doc fetch response HTTP ${status}, length: ${text.length}`);

  if (status === 200 && hasDocCancellationContent(text)) {
    // Cache valid HTML for resilient trigger execution
    PropertiesService.getScriptProperties().setProperty("LAST_VALID_HTML", text);
    PropertiesService.getScriptProperties().setProperty("LAST_VALID_HTML_DATE", new Date().toISOString());
    return text;
  }

  // Diagnostic warning for inspection
  const headers = response.getAllHeaders();
  console.warn("Response headers: " + JSON.stringify(headers));
  console.warn("Response preview: " + text.slice(0, 300).replace(/\s+/g, " "));

  const isAccountChooserOrLogin =
    text.indexOf("accounts.google.com") !== -1 ||
    text.indexOf("accountchooser") !== -1 ||
    text.indexOf("AccountsSignInUi") !== -1 ||
    text.indexOf("ServiceLogin") !== -1;

  // Fallback to recent HTML synced directly from the Chrome extension
  const cachedHtml = PropertiesService.getScriptProperties().getProperty("LAST_VALID_HTML");
  const cachedDate = PropertiesService.getScriptProperties().getProperty("LAST_VALID_HTML_DATE");
  if (cachedHtml && hasDocCancellationContent(cachedHtml)) {
    console.warn(
      `Direct UrlFetchApp encountered Google auth restriction (${isAccountChooserOrLogin ? "Account Chooser" : "HTTP " + status}). ` +
      `Falling back to cancellation HTML captured directly via Chrome extension (${cachedDate}).`
    );
    return cachedHtml;
  }

  if (isAccountChooserOrLogin) {
    throw new Error(
      `Failed to access BCA Class Cancellation List (Google redirected to Sign-in / Accounts).\n\n` +
      `Google Apps Script's UrlFetchApp automatically strips the 'Cookie' header when requesting Google-owned services (docs.google.com) for platform security. As a result, Google Docs receives the cloud request as unauthenticated and redirects to the sign-in page.\n\n` +
      `How to sync:\n` +
      `1. In Chrome, open the BCA cancellation list (or click 'Open Doc in Tab' in the extension).\n` +
      `2. In the BCA Absence Sync Chrome extension, click 'Sync Cookie Now'.\n` +
      `The extension extracts the cancellation document directly from your authenticated Chrome browser session and pushes it to Google Sheets.`
    );
  }

  throw new Error(
    `Failed to access BCA Class Cancellation List (HTTP ${status}).\n\n` +
    `Response length: ${text.length}. Preview: ${text.slice(0, 100).replace(/\s+/g, ' ')}\n\n` +
    "The session cookie from the Chrome extension has expired or is invalid.\n" +
    "Please open Chrome and click 'Sync Cookie Now' in the extension to refresh your credentials."
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
 * If cached HTML from the Chrome extension is available, stages it to Sheets.
 * Otherwise, attempts direct fetch.
 */
function testDocToSheetsSync() {
  const cachedHtml = PropertiesService.getScriptProperties().getProperty("LAST_VALID_HTML");
  const cachedDate = PropertiesService.getScriptProperties().getProperty("LAST_VALID_HTML_DATE");

  if (cachedHtml && hasDocCancellationContent(cachedHtml)) {
    console.log(`Using cancellation document HTML staged from Chrome extension (${cachedDate})...`);
    const dateInfo = parseDocDate(cachedHtml);
    console.log("Parsed date: " + dateInfo.formattedDate);
    const rows = parseDocTable(cachedHtml);
    console.log(`Parsed ${rows.length} teacher absence row(s).`);
    const sheet = getDocTargetSheet();
    writeDocDataToSheet(sheet, dateInfo, rows);
    console.log("Doc to Sheets sync completed successfully using Chrome extension staged data.");
    return;
  }

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
    const nowIso = new Date().toISOString();

    if (cookie) {
      scriptProps.setProperty("DOC_COOKIE", cookie);
      scriptProps.setProperty("DOC_COOKIE_UPDATED_AT", nowIso);
      console.log(`Updated DOC_COOKIE via Web App at ${nowIso}, length: ${cookie.length}`);
    }

    if (payload.authUser && payload.authUser.trim()) {
      scriptProps.setProperty("AUTH_USER", payload.authUser.trim());
    }

    // Direct HTML ingestion: if Chrome extension fetched or extracted the document HTML directly
    if (payload.html && hasDocCancellationContent(payload.html)) {
      console.log(`Received full document HTML directly from Chrome extension (${payload.html.length} chars).`);
      try {
        const dateInfo = parseDocDate(payload.html);
        const rows = parseDocTable(payload.html);
        const sheet = getDocTargetSheet();
        writeDocDataToSheet(sheet, dateInfo, rows);

        // Cache valid HTML for subsequent trigger runs
        scriptProps.setProperty("LAST_VALID_HTML", payload.html);
        scriptProps.setProperty("LAST_VALID_HTML_DATE", nowIso);

        console.log(`Directly staged ${rows.length} teacher absence row(s) for ${dateInfo.formattedDate}.`);
        return ContentService.createTextOutput(
          JSON.stringify({
            status: "success",
            message: `Document synced directly from Chrome! Staged ${rows.length} absence(s) for ${dateInfo.formattedDate}.`,
            date: dateInfo.formattedDate,
            rowCount: rows.length,
            updatedAt: nowIso,
          })
        ).setMimeType(ContentService.MimeType.JSON);
      } catch (directErr) {
        console.warn("Direct HTML parsing warning: " + directErr.message);
      }
    }

    if (!cookie && !payload.html) {
      return ContentService.createTextOutput(
        JSON.stringify({
          status: "error",
          message: "Missing 'cookie' or 'html' field in request body.",
        })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    let syncMessage = "Cookie successfully updated.";
    const cachedHtml = scriptProps.getProperty("LAST_VALID_HTML");
    if (payload.triggerSync === true || payload.triggerSync === "true") {
      if (cachedHtml && hasDocCancellationContent(cachedHtml)) {
        try {
          const dateInfo = parseDocDate(cachedHtml);
          const rows = parseDocTable(cachedHtml);
          const sheet = getDocTargetSheet();
          writeDocDataToSheet(sheet, dateInfo, rows);
          syncMessage = `Cookie updated and staged ${rows.length} absence(s) from document cache.`;
        } catch (cacheErr) {
          console.warn("Cached HTML staging error: " + cacheErr.message);
        }
      } else {
        syncMessage = "Cookie successfully synchronized. Keep the cancellation document open in Chrome for instant table sync.";
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

