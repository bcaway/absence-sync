/**
 * BCA Class Cancellation List to Google Sheets Sync (Native Google Workspace)
 *
 * Reads the BCA Class Cancellation Google Doc natively using DocumentApp,
 * parses the date and teacher absences table, and stages formatted data in Google Sheets.
 *
 * Runs autonomously 24/7 via Apps Script time-driven triggers with zero cookies,
 * zero HTTP requests, and zero browser extensions required.
 */

const DOC_CONFIG = {
  // Google Drive File ID or full URL of the BCA Class Cancellation document.
  // Set in Script Properties under 'DOC_ID', or paste directly below:
  get DOC_ID() {
    const raw = PropertiesService.getScriptProperties().getProperty("DOC_ID") || "";
    return extractDocId(raw);
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
 * Natively opens the cancellation doc, parses the date and absences table, and populates the Sheet.
 */
function syncDocToSheets() {
  const docId = DOC_CONFIG.DOC_ID;

  if (!docId) {
    throw new Error(
      "Missing DOC_ID in Script Properties.\n\n" +
      "Please set 'DOC_ID' in Project Settings > Script Properties (or run searchDriveForCancellationDoc() to find it).\n" +
      "You can provide either the 44-character document ID or the full Google Docs URL."
    );
  }

  console.log("Opening BCA Class Cancellation document (ID: " + docId + ")...");
  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();

  console.log("Parsing document date...");
  const fullText = body.getText();
  const dateInfo = parseDocDate(fullText);
  console.log(`Parsed date: ${dateInfo.formattedDate} (${dateInfo.year}-${dateInfo.month}-${dateInfo.day})`);

  console.log("Parsing cancellation table...");
  const rows = parseDocTables(doc);
  console.log(`Parsed ${rows.length} teacher absence row(s).`);

  const sheet = getDocTargetSheet();
  writeDocDataToSheet(sheet, dateInfo, rows);

  console.log(`Doc to Sheets sync completed successfully. Staged ${rows.length} absence(s) for ${dateInfo.formattedDate}.`);
}


/**
 * Manual test function for doc-to-sheets execution in Apps Script editor.
 */
function testDocToSheetsSync() {
  console.log("Starting manual test of BCA Doc to Sheets sync...");
  syncDocToSheets();
}


/**
 * Helper utility to search your Google Drive for the BCA Class Cancellation document.
 * Run this function in the Apps Script editor to discover the Document ID if you don't know it!
 */
function searchDriveForCancellationDoc() {
  console.log("Searching Google Drive for BCA cancellation documents...");

  const queries = [
    'title contains "Cancellation List" and mimeType = "application/vnd.google-apps.document"',
    'title contains "Class Cancellation" and mimeType = "application/vnd.google-apps.document"',
    'title contains "Cancellation" and mimeType = "application/vnd.google-apps.document"',
  ];

  const foundFiles = [];
  const seenIds = new Set();

  for (const query of queries) {
    try {
      const files = DriveApp.searchFiles(query);
      while (files.hasNext()) {
        const file = files.next();
        const id = file.getId();
        if (!seenIds.has(id)) {
          seenIds.add(id);
          foundFiles.push({
            name: file.getName(),
            id: id,
            url: file.getUrl(),
            lastUpdated: file.getLastUpdated(),
          });
        }
      }
    } catch (err) {
      console.warn(`Query "${query}" search error: ${err.message}`);
    }
  }

  if (foundFiles.length === 0) {
    console.warn(
      "No documents found matching 'Cancellation List' in your Drive.\n" +
      "Make sure you are running Apps Script with your @bergen.org account and have opened/viewed the document in Google Drive."
    );
    return;
  }

  console.log(`Found ${foundFiles.length} candidate file(s):`);
  for (let i = 0; i < foundFiles.length; i++) {
    const f = foundFiles[i];
    console.log(`[${i + 1}] Name: "${f.name}"`);
    console.log(`    Doc ID: ${f.id}`);
    console.log(`    URL: ${f.url}`);
    console.log(`    Last Updated: ${f.lastUpdated}`);
  }

  // Automatically save the first matching file ID if DOC_ID is not configured
  const currentDocId = PropertiesService.getScriptProperties().getProperty("DOC_ID");
  if (!currentDocId && foundFiles.length > 0) {
    const chosen = foundFiles[0];
    PropertiesService.getScriptProperties().setProperty("DOC_ID", chosen.id);
    console.log(`Automatically set Script Property 'DOC_ID' to: ${chosen.id} ("${chosen.name}")`);
  }
}


/**
 * Extracts a Google Drive document ID from either a raw ID string or a full URL.
 *
 * Examples handled:
 * - "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms"
 * - "https://docs.google.com/document/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit"
 * - "https://docs.google.com/document/u/1/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/preview"
 */
function extractDocId(urlOrId) {
  if (!urlOrId) return "";
  const trimmed = urlOrId.trim();

  // If already just the ID (alphanumeric, dashes, underscores, length >= 20)
  if (/^[a-zA-Z0-9_-]{20,}$/.test(trimmed)) {
    return trimmed;
  }

  // Match /document/d/([a-zA-Z0-9_-]+) or /document/u/\d+/d/([a-zA-Z0-9_-]+)
  const match = trimmed.match(/\/document\/(?:u\/\d+\/)?d\/([a-zA-Z0-9_-]+)/i);
  if (match && match[1]) {
    return match[1];
  }

  return trimmed;
}


/**
 * Parses the document date from the text header.
 *
 * Expected format:
 * "BCA Class Cancellation List
 * {Month} {Day}, {YYYY}"
 */
function parseDocDate(text) {
  if (!text) {
    throw new Error("Cannot parse date: document body text is empty.");
  }

  // Normalize en-dashes, em-dashes, and non-breaking spaces
  const cleanText = text
    .replace(/[\u2013\u2014]/g, "-")
    .replace(/\u00A0/g, " ");

  // Match header and date
  const headerRegex =
    /BCA\s+Class\s+Cancellation\s+List[\s\S]*?([a-zA-Z]+)\s+(\d{1,2}),?\s+(\d{4})/i;

  let match = cleanText.match(headerRegex);

  // Fallback: look for "Month Day, Year" anywhere in the document
  if (!match) {
    const fallbackRegex = /([a-zA-Z]+)\s+(\d{1,2}),?\s+(\d{4})/i;
    match = cleanText.match(fallbackRegex);
  }

  if (!match) {
    throw new Error("Could not find cancellation list date in document header.");
  }

  const rawMonth = match[1].toLowerCase();
  const rawDay = match[2];
  const rawYear = match[3];

  const month = DOC_MONTH_MAP[rawMonth];
  if (!month) {
    throw new Error(`Unrecognized month name in document header: "${match[1]}"`);
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
 * Parses teacher cancellations natively from the document's Table element(s).
 *
 * Rules:
 * - Row 0 is the header row: ignored.
 * - Column 0: Teacher name string.
 * - Column 1: Periods string, formatted according to rules.
 * - Column 2: Rightmost column, ignored.
 *
 * Returns array of [teacher, formattedPeriods].
 */
function parseDocTables(doc) {
  const body = doc.getBody();
  const tables = body.getTables();

  if (!tables || tables.length === 0) {
    throw new Error("No tables found in the Google Document.");
  }

  // Locate the cancellation table (table with > 1 row containing Teacher / Period headers or first table)
  let targetTable = tables[0];
  for (const t of tables) {
    if (t.getNumRows() > 1) {
      const headerText = t.getRow(0).getText().toLowerCase();
      if (headerText.includes("teacher") || headerText.includes("period") || headerText.includes("absence")) {
        targetTable = t;
        break;
      }
    }
  }

  const numRows = targetTable.getNumRows();
  if (numRows <= 1) {
    return [];
  }

  const parsedRows = [];

  // Start at row index 1 (skipping header row 0)
  for (let i = 1; i < numRows; i++) {
    const row = targetTable.getRow(i);
    const numCells = row.getNumCells();
    if (numCells < 2) continue;

    const teacher = row.getCell(0).getText().trim();
    const rawCol2 = row.getCell(1).getText().trim();

    // Ignore completely empty rows
    if (!teacher && !rawCol2) {
      continue;
    }

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
 * Retrieves the target Google Sheet for doc-to-sheets sync.
 */
function getDocTargetSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheet) {
    throw new Error(
      "No active spreadsheet found. Make sure this Apps Script is bound to your Google Sheet."
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

  // If sheets-to-supabase is in the same Apps Script project, chain sync immediately
  if (typeof syncAbsences === "function") {
    console.log("Triggering sheets-to-supabase sync...");
    try {
      syncAbsences();
    } catch (err) {
      console.warn("sheets-to-supabase sync notice: " + err.message);
    }
  }
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

  console.log(`Doc-to-sheets sync trigger created: runs natively every ${minutes || 5} minute(s) in Google Cloud.`);
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
