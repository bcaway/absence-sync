const CONFIG = {
  DOCUMENT_ID: PropertiesService
    .getScriptProperties()
    .getProperty("DOCUMENT_ID"),

  SUPABASE_URL: PropertiesService
    .getScriptProperties()
    .getProperty("SUPABASE_URL"),

  SYNC_SECRET: PropertiesService
    .getScriptProperties()
    .getProperty("SYNC_SECRET"),
};


/**
 * Main sync function.
 */
function syncAbsences() {
  if (!CONFIG.DOCUMENT_ID) {
    throw new Error("DOCUMENT_ID is not configured.");
  }

  if (!CONFIG.SUPABASE_URL) {
    throw new Error("SUPABASE_URL is not configured.");
  }

  if (!CONFIG.SYNC_SECRET) {
    throw new Error("SYNC_SECRET is not configured.");
  }

  const document = DocumentApp.openById(CONFIG.DOCUMENT_ID);
  const body = document.getBody();

  const data = parseDocument(body);

  sendToSupabase(data);
}


/**
 * Parses the Google Doc.
 *
 * Expected format:
 *
 * BCA Class Cancellation List
 *
 * September 18, 2026
 *
 * Teacher | Periods Impacted
 * Bian    | All Day
 * Cardenas| All Day
 * Molino  | Periods 2-9
 */
function parseDocument(body) {
  const tables = body.getTables();

  if (tables.length === 0) {
    throw new Error("No table found in Google Doc.");
  }

  const table = tables[0];

  if (table.getNumColumns() < 2) {
    throw new Error("Expected a two-column table.");
  }

  const date = findDate(body);

  if (!date) {
    throw new Error("Could not find a date in the document.");
  }

  const absences = [];

  for (let rowIndex = 1; rowIndex < table.getNumRows(); rowIndex++) {
    const row = table.getRow(rowIndex);

    const teacher = row
      .getCell(0)
      .getText()
      .trim();

    const periodsImpacted = row
      .getCell(1)
      .getText()
      .trim();

    if (!teacher || !periodsImpacted) {
      continue;
    }

    absences.push({
      teacher: teacher,
      periods_impacted: periodsImpacted,
    });
  }

  return {
    date: date,
    absences: absences,
  };
}


/**
 * Finds a date such as:
 *
 * September 18, 2026
 */
function findDate(body) {
  const text = body.getText();

  const match = text.match(
    /\b(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}\b/i
  );

  if (!match) {
    return null;
  }

  const parsed = new Date(match[0]);

  if (isNaN(parsed.getTime())) {
    throw new Error("Found a date but could not parse it.");
  }

  return Utilities.formatDate(
    parsed,
    Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );
}


/**
 * Sends the parsed data to Supabase.
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
  const body = response.getContentText();

  if (status < 200 || status >= 300) {
    throw new Error(
      "Supabase sync failed (" +
      status +
      "): " +
      body
    );
  }

  console.log(body);
}

function createFiveMinuteTrigger() {
  ScriptApp
    .newTrigger("syncAbsences")
    .timeBased()
    .everyMinutes(5)
    .create();
}

function testSync() {
  syncAbsences();
}