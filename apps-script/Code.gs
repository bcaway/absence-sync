const CONFIG = {
  DOCUMENT_URL: PropertiesService
    .getScriptProperties()
    .getProperty("DOCUMENT_URL"),

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
  if (!CONFIG.DOCUMENT_URL) {
    throw new Error("DOCUMENT_URL is not configured.");
  }

  if (!CONFIG.SUPABASE_URL) {
    throw new Error("SUPABASE_URL is not configured.");
  }

  if (!CONFIG.SYNC_SECRET) {
    throw new Error("SYNC_SECRET is not configured.");
  }

  const response = UrlFetchApp.fetch(CONFIG.DOCUMENT_URL, {
    method: "get",
    muteHttpExceptions: true,
  });

  const status = response.getResponseCode();

  if (status !== 200) {
    throw new Error(
      "Failed to fetch Google Doc: HTTP " + status
    );
  }

  const html = response.getContentText();

  const data = parseDocument(html);

  sendToSupabase(data);
}


/**
 * Parses the published Google Doc HTML.
 *
 * Expected document:
 *
 * BCA Class Cancellation List
 *
 * September 18, 2026
 *
 * Teacher | Periods Impacted
 * Bian | All Day
 * Cardenas | All Day
 * Molino | Periods 2-9
 */
function parseDocument(html) {
  const text = htmlToText(html);

  const date = findDate(text);

  if (!date) {
    throw new Error(
      "Could not find a date in the published document."
    );
  }

  const table = extractFirstTable(html);

  if (!table || table.length < 2) {
    throw new Error(
      "Could not find a valid cancellation table."
    );
  }

  const absences = [];

  // Skip the header row.
  for (let i = 1; i < table.length; i++) {
    const row = table[i];

    if (row.length < 2) {
      continue;
    }

    const teacher = cleanText(row[0]);
    const periodsImpacted = cleanText(row[1]);

    if (!teacher || !periodsImpacted) {
      continue;
    }

    absences.push({
      teacher: teacher,
      periods_impacted: periodsImpacted,
    });
  }

  if (absences.length === 0) {
    throw new Error(
      "No teacher cancellation entries were found."
    );
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
function findDate(text) {
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
 * Extracts HTML tables.
 *
 * Returns:
 *
 * [
 *   ["Teacher", "Periods Impacted"],
 *   ["Bian", "All Day"],
 *   ["Molino", "Periods 2-9"]
 * ]
 */
function extractFirstTable(html) {
  const tableMatch = html.match(
    /<table[\s\S]*?<\/table>/i
  );

  if (!tableMatch) {
    return null;
  }

  const tableHtml = tableMatch[0];

  const rowMatches = tableHtml.match(
    /<tr[\s\S]*?<\/tr>/gi
  );

  if (!rowMatches) {
    return null;
  }

  const rows = [];

  for (const rowHtml of rowMatches) {
    const cellMatches = rowHtml.match(
      /<(td|th)[^>]*>[\s\S]*?<\/\1>/gi
    );

    if (!cellMatches) {
      continue;
    }

    const row = cellMatches.map((cell) => {
      return cleanText(cell);
    });

    if (row.length > 0) {
      rows.push(row);
    }
  }

  return rows;
}


/**
 * Converts HTML into readable plain text.
 */
function htmlToText(html) {
  return html
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<\/div>/gi, "\n")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&#39;/gi, "'")
    .replace(/&quot;/gi, '"')
    .replace(/\s+/g, " ")
    .trim();
}


/**
 * Cleans a table cell while preserving the actual text.
 */
function cleanText(value) {
  return value
    .replace(/<br\s*\/?>/gi, " ")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&#39;/gi, "'")
    .replace(/&quot;/gi, '"')
    .replace(/\s+/g, " ")
    .trim();
}


/**
 * Sends the parsed snapshot to Supabase.
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
 * Deletes an existing sync trigger first so running this
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