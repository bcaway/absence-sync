import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

type Absence = {
  teacher: string;
  periods_impacted: string;
};

type SyncPayload = {
  date: string;
  absences: Absence[];
};

const SUPABASE_URL = Deno.env.get("SUPABASE_URL")!;
const SUPABASE_SERVICE_ROLE_KEY = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;
const SYNC_SECRET = Deno.env.get("SYNC_SECRET")!;

const supabase = createClient(
  SUPABASE_URL,
  SUPABASE_SERVICE_ROLE_KEY,
);

const ALL_PERIODS = [
  "igs",
  "1",
  "2",
  "3",
  "4",
  "5",
  "6",
  "7",
  "8",
  "9",
];

function normalizePeriods(rawValue: string): string {
  const periods = new Set<string>();

  const tokens = rawValue
    .split(",")
    .map((token) => token.trim().toLowerCase())
    .filter(Boolean);

  // "all" always means every period.
  if (tokens.includes("all")) {
    return ALL_PERIODS.join(", ");
  }

  for (const token of tokens) {
    if (token === "igs") {
      periods.add("igs");
      continue;
    }

    // Single numerical period.
    if (/^[1-9]$/.test(token)) {
      periods.add(token);
      continue;
    }

    // Numerical range, e.g. 2-5.
    const rangeMatch = token.match(/^([1-9])-([1-9])$/);

    if (rangeMatch) {
      const start = Number(rangeMatch[1]);
      const end = Number(rangeMatch[2]);

      if (start <= end) {
        for (let period = start; period <= end; period++) {
          periods.add(String(period));
        }
      }
    }
  }

  return [
    ...(periods.has("igs") ? ["igs"] : []),
    ...Array.from({ length: 9 }, (_, i) => String(i + 1))
      .filter((period) => periods.has(period)),
  ].join(", ");
}

function normalizeAbsences(absences: Absence[]): Absence[] {
  return absences
    .map((absence) => ({
      teacher: absence.teacher.trim(),
      periods_impacted: normalizePeriods(absence.periods_impacted),
    }))
    .filter(
      (absence) =>
        absence.teacher.length > 0 &&
        absence.periods_impacted.length > 0,
    )
    .sort((a, b) => a.teacher.localeCompare(b.teacher));
}

function getTodayInNewYork(): string {
  return new Intl.DateTimeFormat("en-CA", {
    timeZone: "America/New_York",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).format(new Date());
}

function jsonResponse(
  body: Record<string, unknown>,
  status = 200,
): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: {
      "Content-Type": "application/json",
    },
  });
}

Deno.serve(async (req) => {
  try {
    if (req.method !== "POST") {
      return jsonResponse(
        { error: "Method not allowed" },
        405,
      );
    }

    const authorization = req.headers.get("Authorization");

    if (authorization !== `Bearer ${SYNC_SECRET}`) {
      return jsonResponse(
        { error: "Unauthorized" },
        401,
      );
    }

    const payload = (await req.json()) as SyncPayload;

    if (
      !payload ||
      typeof payload.date !== "string" ||
      !Array.isArray(payload.absences)
    ) {
      return jsonResponse(
        { error: "Invalid payload" },
        400,
      );
    }

    const normalizedAbsences = normalizeAbsences(payload.absences);
    const today = getTodayInNewYork();

    /*
     * The Supabase table is a LIVE representation of today's
     * cancellation data.
     *
     * If the Sheet contains anything other than today's date,
     * the database must be empty.
     */
    if (payload.date !== today) {
      const { error } = await supabase
        .from("teacher_absences")
        .delete()
        .not("id", "is", null);

      if (error) {
        console.error("Failed to clear stale data:", error);

        return jsonResponse(
          {
            error: "Failed to clear stale absence data",
            details: error.message,
          },
          500,
        );
      }

      return jsonResponse({
        success: true,
        changed: true,
        cleared: true,
        reason: "Sheet date is not today",
        sheet_date: payload.date,
        today,
      });
    }

    /*
     * Get the current live state.
     */
    const { data: currentRows, error: fetchError } = await supabase
      .from("teacher_absences")
      .select("teacher, periods_impacted")
      .eq("date", today)
      .order("teacher", { ascending: true });

    if (fetchError) {
      console.error("Failed to fetch current data:", fetchError);

      return jsonResponse(
        {
          error: "Failed to fetch current absence data",
          details: fetchError.message,
        },
        500,
      );
    }

    const currentAbsences = normalizeAbsences(
      (currentRows ?? []).map((row) => ({
        teacher: row.teacher,
        periods_impacted: row.periods_impacted,
      })),
    );

    const currentSnapshot = JSON.stringify(currentAbsences);
    const incomingSnapshot = JSON.stringify(normalizedAbsences);

    /*
     * Nothing changed. Leave Supabase completely untouched.
     */
    if (currentSnapshot === incomingSnapshot) {
      return jsonResponse({
        success: true,
        changed: false,
        cleared: false,
        date: today,
        count: normalizedAbsences.length,
      });
    }

    /*
     * Something changed.
     *
     * Replace the entire live snapshot. This handles:
     * - new teachers
     * - removed teachers
     * - changed periods
     * - all -> specific periods
     * - specific periods -> all
     * - going from data -> zero absences
     */
    const { error: deleteError } = await supabase
      .from("teacher_absences")
      .delete()
      .not("id", "is", null);

    if (deleteError) {
      console.error("Failed to clear old snapshot:", deleteError);

      return jsonResponse(
        {
          error: "Failed to clear old absence snapshot",
          details: deleteError.message,
        },
        500,
      );
    }

    /*
     * Empty snapshot is valid. The table simply remains empty.
     */
    if (normalizedAbsences.length === 0) {
      return jsonResponse({
        success: true,
        changed: true,
        cleared: true,
        date: today,
        count: 0,
      });
    }

    const syncedAt = new Date().toISOString();

    const rowsToInsert = normalizedAbsences.map((absence) => ({
      date: today,
      synced_at: syncedAt,
      teacher: absence.teacher,
      periods_impacted: absence.periods_impacted,
    }));

    const { error: insertError } = await supabase
      .from("teacher_absences")
      .insert(rowsToInsert);

    if (insertError) {
      console.error("Failed to insert new snapshot:", insertError);

      return jsonResponse(
        {
          error: "Failed to insert new absence snapshot",
          details: insertError.message,
        },
        500,
      );
    }

    return jsonResponse({
      success: true,
      changed: true,
      cleared: false,
      date: today,
      count: normalizedAbsences.length,
      synced_at: syncedAt,
    });
  } catch (error) {
    console.error("Unexpected error:", error);

    return jsonResponse(
      {
        error: "Internal server error",
        details: error instanceof Error ? error.message : String(error),
      },
      500,
    );
  }
});