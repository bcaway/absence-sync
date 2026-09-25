# absence-sync

Automated pipeline syncing BCA teacher attendance and class cancellations from the published Google Doc to Google Sheets and Supabase.

## Architecture

```mermaid
flowchart LR
    Doc["Published Google Doc<br/>(Class Cancellation List)"]
    Sheet["Google Sheet<br/>(A1: Date, A: Teacher, B: Periods)"]
    Supabase["Supabase Edge Function<br/>(sync-absences)"]
    DB[("Supabase DB<br/>(teacher_absences)")]

    Doc -->|"doc-to-sheets (Apps Script)"| Sheet
    Sheet -->|"sheets-to-supabase (Apps Script)"| Supabase
    Supabase --> DB
```

---

## 1. doc-to-sheets

Google Apps Script that fetches the published BCA Class Cancellation List document, parses the date and teacher absences table, and populates the Google Sheet.

### Setup Instructions

1. Open your target Google Sheet.
2. Click **Extensions** > **Apps Script**.
3. Copy the contents of [`doc-to-sheets/Code.gs`](./doc-to-sheets/Code.gs) into `Code.gs`.
4. (Optional) If editing the manifest, enable "Show appsscript.json manifest file in editor" under Project Settings and copy [`doc-to-sheets/appsscript.json`](./doc-to-sheets/appsscript.json).
5. **Script Properties** (optional):
   - `DOC_URL`: The published doc URL (defaults to the BCA cancellation list `/pub` URL).
   - `SHEET_NAME`: Name of the specific tab to write to (defaults to the active/first sheet).
   - `SPREADSHEET_ID`: Only required if running as a standalone script rather than bound to the Google Sheet.
6. **Automation**:
   - Run `createFiveMinuteTrigger()` in the Apps Script editor to create a recurring time-driven trigger that runs every 5 minutes.
   - Alternatively, run `createOneMinuteTrigger()` for 1-minute updates.
   - Run `testSync()` to manually verify.

### Google Sheet Format

- **Cell A1**: Date in `M/D/YYYY` format without leading zeroes (e.g. `9/6/2026`).
- **Row 2 onwards**:
  - **Column A**: Teacher name.
  - **Column B**: Normalized periods impacted:
    - `"All"` if the doc specifies all day.
    - Ranges (`1-3, 7-8`), standalone periods (`5`), and/or `"igs"` separated by commas (e.g. `1-3, 7-8, 5, igs`).

---

## 2. sheets-to-supabase

Google Apps Script bound to the Google Sheet that reads the staged date and teacher absences, validates the data, and sends it to the Supabase Edge Function.

### Setup Instructions

1. In the same (or bound) Google Apps Script project, configure:
   - `SUPABASE_URL`: Your Supabase project URL (e.g. `https://xyz.supabase.co`).
   - `SYNC_SECRET`: Bearer authentication secret matching your Supabase environment.
2. Run `createFiveMinuteTrigger()` to automate sync to Supabase.
3. Run `testSync()` to manually execute.

---

## 3. supabase

Contains the database migrations and Edge Function:
- `supabase/functions/sync-absences`: Validates payload, verifies teacher attendance, and updates database records.
- `supabase/migrations`: SQL migrations setting up `teacher_absences` and row-level security.
