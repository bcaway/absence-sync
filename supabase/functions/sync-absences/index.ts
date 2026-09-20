import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

type Absence = {
  teacher: string;
  periods_impacted: string;
};

type SyncPayload = {
  date: string;
  absences: Absence[];
};

const supabase = createClient(
  Deno.env.get("SUPABASE_URL")!,
  Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!
);

const SYNC_SECRET = Deno.env.get("SYNC_SECRET");

Deno.serve(async (req) => {
  try {
    if (req.method !== "POST") {
      return new Response(
        JSON.stringify({ error: "Method not allowed" }),
        {
          status: 405,
          headers: { "Content-Type": "application/json" },
        }
      );
    }

    const authHeader = req.headers.get("Authorization");

    if (!SYNC_SECRET || authHeader !== `Bearer ${SYNC_SECRET}`) {
      return new Response(
        JSON.stringify({ error: "Unauthorized" }),
        {
          status: 401,
          headers: { "Content-Type": "application/json" },
        }
      );
    }

    const payload = (await req.json()) as SyncPayload;

    if (!payload.date || !Array.isArray(payload.absences)) {
      return new Response(
        JSON.stringify({ error: "Invalid payload" }),
        {
          status: 400,
          headers: { "Content-Type": "application/json" },
        }
      );
    }

    const absences = payload.absences.map((absence) => ({
      teacher: String(absence.teacher).trim(),
      periods_impacted: String(absence.periods_impacted),
    }));

    /*
     * Sort the data before comparing it.
     *
     * This means that changing the order of rows in the Google Doc
     * does not count as a data change.
     */
    absences.sort((a, b) =>
      a.teacher.localeCompare(b.teacher)
    );

    /*
     * Find the latest snapshot for this date.
     */
    const { data: latestRows, error: latestError } = await supabase
      .from("teacher_absences")
      .select("teacher, periods_impacted")
      .eq("date", payload.date)
      .order("synced_at", { ascending: false });

    if (latestError) {
      throw latestError;
    }

    let previousSnapshot: Absence[] = [];

    if (latestRows && latestRows.length > 0) {
      /*
       * Because every snapshot has the same synced_at,
       * the newest rows represent the latest complete snapshot.
       */
      const latestSyncedAt = await getLatestSyncedAt(payload.date);

      const { data: snapshotRows, error: snapshotError } =
        await supabase
          .from("teacher_absences")
          .select("teacher, periods_impacted")
          .eq("date", payload.date)
          .eq("synced_at", latestSyncedAt);

      if (snapshotError) {
        throw snapshotError;
      }

      previousSnapshot = (snapshotRows ?? [])
        .map((row) => ({
          teacher: row.teacher,
          periods_impacted: row.periods_impacted,
        }))
        .sort((a, b) => a.teacher.localeCompare(b.teacher));
    }

    const changed =
      JSON.stringify(absences) !== JSON.stringify(previousSnapshot);

    if (!changed) {
      return new Response(
        JSON.stringify({
          changed: false,
          inserted: 0,
          message: "No changes detected.",
        }),
        {
          status: 200,
          headers: { "Content-Type": "application/json" },
        }
      );
    }

    const syncedAt = new Date().toISOString();

    const rowsToInsert = absences.map((absence) => ({
      date: payload.date,
      synced_at: syncedAt,
      teacher: absence.teacher,
      periods_impacted: absence.periods_impacted,
    }));

    if (rowsToInsert.length > 0) {
      const { error: insertError } = await supabase
        .from("teacher_absences")
        .insert(rowsToInsert);

      if (insertError) {
        throw insertError;
      }
    }

    return new Response(
      JSON.stringify({
        changed: true,
        inserted: rowsToInsert.length,
        synced_at: syncedAt,
      }),
      {
        status: 200,
        headers: { "Content-Type": "application/json" },
      }
    );
  } catch (error) {
    console.error(error);

    return new Response(
      JSON.stringify({
        error: "Internal server error",
      }),
      {
        status: 500,
        headers: { "Content-Type": "application/json" },
      }
    );
  }
});


async function getLatestSyncedAt(date: string): Promise<string> {
  const { data, error } = await supabase
    .from("teacher_absences")
    .select("synced_at")
    .eq("date", date)
    .order("synced_at", { ascending: false })
    .limit(1)
    .single();

  if (error) {
    throw error;
  }

  return data.synced_at;
}