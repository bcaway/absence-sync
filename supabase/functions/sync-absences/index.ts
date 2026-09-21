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
      return jsonResponse(
        { error: "Method not allowed" },
        405
      );
    }

    const authHeader = req.headers.get("Authorization");

    if (
      !SYNC_SECRET ||
      authHeader !== `Bearer ${SYNC_SECRET}`
    ) {
      return jsonResponse(
        { error: "Unauthorized" },
        401
      );
    }

    const payload = (await req.json()) as SyncPayload;

    if (
      !payload.date ||
      !Array.isArray(payload.absences)
    ) {
      return jsonResponse(
        { error: "Invalid payload" },
        400
      );
    }

    /*
     * Normalize every absence on the server.
     */
    const absences = payload.absences
      .map((absence) => {
        const teacher = String(absence.teacher).trim();

        const periods = normalizePeriods(
          String(absence.periods_impacted)
        );

        return {
          teacher,
          periods_impacted: periods,
        };
      })
      .filter(
        (absence) =>
          absence.teacher.length > 0 &&
          absence.periods_impacted.length > 0
      );

    /*
     * Sort teachers so row ordering never affects
     * snapshot comparison.
     */
    absences.sort((a, b) =>
      a.teacher.localeCompare(b.teacher)
    );

    /*
     * Get the timestamp of the latest complete snapshot.
     */
    const { data: latestRow, error: latestError } =
      await supabase
        .from("teacher_absences")
        .select("synced_at")
        .eq("date", payload.date)
        .order("synced_at", {
          ascending: false,
        })
        .limit(1)
        .maybeSingle();

    if (latestError) {
      throw latestError;
    }

    let previousSnapshot: Absence[] = [];

    if (latestRow) {
      const { data: snapshotRows, error: snapshotError } =
        await supabase
          .from("teacher_absences")
          .select(
            "teacher, periods_impacted"
          )
          .eq("date", payload.date)
          .eq(
            "synced_at",
            latestRow.synced_at
          );

      if (snapshotError) {
        throw snapshotError;
      }

      previousSnapshot = (snapshotRows ?? [])
        .map((row) => ({
          teacher: row.teacher,
          periods_impacted:
            normalizePeriods(
              row.periods_impacted
            ),
        }))
        .sort((a, b) =>
          a.teacher.localeCompare(b.teacher)
        );
    }

    /*
     * Compare canonical data.
     */
    const changed =
      JSON.stringify(absences) !==
      JSON.stringify(previousSnapshot);

    if (!changed) {
      return jsonResponse({
        changed: false,
        inserted: 0,
        message: "No changes detected.",
      });
    }

    const syncedAt =
      new Date().toISOString();

    const rowsToInsert = absences.map(
      (absence) => ({
        date: payload.date,
        synced_at: syncedAt,
        teacher: absence.teacher,
        periods_impacted:
          absence.periods_impacted,
      })
    );

    /*
     * Insert the entire new snapshot.
     *
     * No existing rows are updated or deleted.
     */
    if (rowsToInsert.length > 0) {
      const { error: insertError } =
        await supabase
          .from("teacher_absences")
          .insert(rowsToInsert);

      if (insertError) {
        throw insertError;
      }
    }

    return jsonResponse({
      changed: true,
      inserted: rowsToInsert.length,
      synced_at: syncedAt,
    });
  } catch (error) {
    console.error(error);

    return jsonResponse(
      {
        error: "Internal server error",
      },
      500
    );
  }
});


/**
 * Normalizes a periods string into canonical form.
 *
 * Examples:
 *
 * all
 * → igs, 1, 2, 3, 4, 5, 6, 7, 8, 9
 *
 * 1-3, 2, igs
 * → igs, 1, 2, 3
 *
 * 2-3, nonsense, igs
 * → igs, 2, 3
 */
function normalizePeriods(
  rawValue: string
): string {
  const tokens = rawValue
    .split(",")
    .map((token) =>
      token.trim().toLowerCase()
    )
    .filter((token) => token.length > 0);

  const periods = new Set<string>();

  for (const token of tokens) {
    /*
     * "all" has absolute priority.
     */
    if (token === "all") {
      return [
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
      ].join(", ");
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
     */
    const rangeMatch =
      token.match(
        /^(\d+)\s*-\s*(\d+)$/
      );

    if (rangeMatch) {
      const start = Number(
        rangeMatch[1]
      );

      const end = Number(
        rangeMatch[2]
      );

      if (
        start >= 1 &&
        start <= 9 &&
        end >= 1 &&
        end <= 9 &&
        start <= end
      ) {
        for (
          let period = start;
          period <= end;
          period++
        ) {
          periods.add(
            String(period)
          );
        }
      }
    }

    /*
     * Anything else is ignored.
     */
  }

  if (periods.size === 0) {
    return "";
  }

  const result: string[] = [];

  if (periods.has("igs")) {
    result.push("igs");
  }

  for (
    let period = 1;
    period <= 9;
    period++
  ) {
    if (
      periods.has(String(period))
    ) {
      result.push(String(period));
    }
  }

  return result.join(", ");
}


/**
 * Creates a JSON response.
 */
function jsonResponse(
  body: unknown,
  status = 200
): Response {
  return new Response(
    JSON.stringify(body),
    {
      status,
      headers: {
        "Content-Type":
          "application/json",
      },
    }
  );
}