// src/utils/excel.ts
import { read, utils } from "xlsx";

/* ============================================================
   Types
============================================================ */
export type Row = Record<string, number>;
export type RowsBySheet = Record<string, Row[]>;

export type Role = "pitcher" | "hitter";

export type Series = {
  /** UI label to display (your left/right column text) */
  label: string;
  /** Exact header key from the Excel sheet (e.g., "/Calc/Pelvis/Twist/Velocity_x") */
  key: string;
  /** Parsed numeric time-series (NaN when missing) */
  values: number[];
};

export type NeededMetrics = {
  time: number[]; // seconds
  series: Series[];
};

export type NeededMetricsByRole = {
  pitcher: NeededMetrics;
  hitter: NeededMetrics;
};

/* ============================================================
   Public API — existing helpers (unchanged signatures)
============================================================ */

/** Parse a File chosen via <input> */
export async function parseExcelToDataSets(
  file: File,
  fpsGuess = 120
): Promise<RowsBySheet> {
  const buf = await file.arrayBuffer();
  return parseWorkbookArrayBuffer(buf, fpsGuess);
}

/** Parse an Excel file fetched from a URL (served from /public) */
export async function parseExcelUrlToDataSets(
  url: string,
  fpsGuess = 120
): Promise<RowsBySheet> {
  const res = await fetch(url);
  if (!res.ok)
    throw new Error(`Failed to fetch Excel: ${res.status} ${res.statusText}`);
  const buf = await res.arrayBuffer();
  return parseWorkbookArrayBuffer(buf, fpsGuess);
}

/** Core parser used by both helpers — returns rows for *all* sheets */
export function parseWorkbookArrayBuffer(
  buf: ArrayBuffer,
  fpsGuess = 120
): RowsBySheet {
  const wb = read(buf, { type: "array" });
  const out: RowsBySheet = {};

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    if (!ws) continue;

    const table: any[][] = utils.sheet_to_json(ws, {
      header: 1,
      defval: null,
      raw: true,
    });
    if (!Array.isArray(table) || table.length === 0) continue;

    const { headers, dataRows } = detectHeaderAndExtract(table);
    if (!headers.length || dataRows.length === 0) continue;

    const rowsRaw = toObjects(headers, dataRows);
    const cleaned = normalizeSheet(rowsRaw, fpsGuess);
    if (
      cleaned.length &&
      cleaned.some((r) => Object.keys(r).some((k) => k !== "t"))
    ) {
      out[sheetName] = cleaned;
    }
  }
  return out;
}

/** Legacy single-sheet helper retained for compatibility. */
export async function parseExcelToRows(
  file: File,
  fpsGuess = 120
): Promise<Row[]> {
  const sets = await parseExcelToDataSets(file, fpsGuess);
  const names = Object.keys(sets);
  const pref =
    names.find((n) => /baseball/i.test(n)) ??
    names.find((n) => /positions|velocity/i.test(n)) ??
    names.find((n) => /signal|data|sheet1/i.test(n)) ??
    names[0];
  return pref ? sets[pref] : [];
}

/* ============================================================
   NEW: Needed-metrics API for your two-column UI
============================================================ */

/** Human-readable label → header key mapping (authoritative) */
const NEEDED_MAP: Record<Role, Array<[label: string, key: string]>> = {
  pitcher: [
    ["Pelvis Twist Velocity", "/Calc/Pelvis/Twist/Velocity_x"],
    ["Shoulder Twist Velocity", "/Calc/Shoulder/Twist/Velocity_x"],
    [
      "Elbow Flexion/Extension Velocity, dominant",
      "/Calc/Elbow/Dominant/FlexionExtension/Velocity_x",
    ],
    [
      "Shoulder Rotation Velocity, dominant",
      "/Calc/Shoulder/Dominant/Rotation/Velocity_x",
    ],
    ["Shoulder Horizontal, dominant", "/Calc/Shoulder/Dominant/Horizontal_x"],
    ["Shoulder Elevation, dominant", "/Calc/Shoulder/Dominant/Elevation_x"],
    ["Shoulder Rotation, dominant", "/Calc/Shoulder/Dominant/Rotation_x"],
    ["Elbow Flexion/Extension, dominant", "/Calc/Elbow/Dominant/FlexionExtension_x"],
    ["Trunk Tilt left/right", "/Calc/Trunk/Tilt/LeftRight_x"],
    ["Shoulder Twist", "/Calc/Shoulder/Twist_x"],
    ["Pelvis Twist", "/Calc/Pelvis/Twist_x"],
    ["Trunk Separation", "/Calc/Trunk/Separation_x"],
    ["Knee Flexion/Extension, lead", "/Calc/Knee/Lead/FlexionExtension_x"],
    ["Center of Gravity Velocity, Y", "/Calc/CenterOfGravity/VelocityY_x"],
  ],
  hitter: [
    ["Shoulder Twist Velocity", "/Calc/Shoulder/Twist/Velocity_x"],
    ["Pelvis Twist Velocity", "/Calc/Pelvis/Twist/Velocity_x"],
    [
      "Elbow Dominant Flexion/Extension Velocity",
      "/Calc/Elbow/Dominant/FlexionExtension/Velocity_x",
    ],
    ["Trunk Separation", "/Calc/Trunk/Separation_x"],
    ["Shoulder Other Horizontal", "/Calc/Shoulder/Other/Horizontal_x"],
    ["Trunk Tilt Left/Right", "/Calc/Trunk/Tilt/LeftRight_x"],
    ["Hip Plant Internal/External", "/Calc/Hip/Plant/InternalExternal_x"],
    ["Shoulder Twist", "/Calc/Shoulder/Twist_x"],
    ["Hip Lead Internal/External", "/Calc/Hip/Lead/InternalExternal_x"],
    ["Trunk Tilt Forwards/Backwards", "/Calc/Trunk/Tilt/ForwardsBackwards_x"],
    ["Shoulder Dominant Horizontal", "/Calc/Shoulder/Dominant/Horizontal_x"],
    ["Shoulder Dominant Rotation", "/Calc/Shoulder/Dominant/Rotation_x"],
    ["Pelvis Tilt Left/Right", "/Calc/Pelvis/Tilt/LeftRight_x"],
    ["Center of Gravity Velocity Z", "/Calc/CenterOfGravity/VelocityZ_x"],
  ],
};

/** Common header variants → canonical Excel key */
const HEADER_ALIASES = new Map<string, string>([
  // Flexion/Extension variants
  [
    "/Calc/Elbow/Dominant/Flexion/Extension/Velocity_x",
    "/Calc/Elbow/Dominant/FlexionExtension/Velocity_x",
  ],
  [
    "/Calc/Elbow/Dominant/Flexion/Extenstion/Velocity_x",
    "/Calc/Elbow/Dominant/FlexionExtension/Velocity_x",
  ],
  [
    "/Calc/Elbow/Dominant/Flexion/Extension_x",
    "/Calc/Elbow/Dominant/FlexionExtension_x",
  ],
  [
    "/Calc/Wrist/Dominant/Flexion/Extension/Velocity_x",
    "/Calc/Wrist/Dominant/FlexionExtension/Velocity_x",
  ],
  [
    "/Calc/Knee/Lead/Flexion/Extension_x",
    "/Calc/Knee/Lead/FlexionExtension_x",
  ],
]);

/** Parse an uploaded file and return ONLY the whitelisted metrics. */
export async function parseExcelToNeededMetrics(
  file: File
): Promise<NeededMetricsByRole> {
  const buf = await file.arrayBuffer();
  return parseWorkbookToNeededMetrics(buf);
}

/** Same as above but accepts an ArrayBuffer (if you already fetched it). */
export function parseWorkbookToNeededMetrics(
  buf: ArrayBuffer
): NeededMetricsByRole {
  const wb = read(buf, { type: "array" });

  // Prefer “Baseball Data” sheet
  const sheetName =
    wb.SheetNames.find((n) => /baseball.*data/i.test(n)) ?? wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error("Excel has no readable sheets.");

  // Raw 2D array
  const table: any[][] = utils.sheet_to_json(ws, {
    header: 1,
    defval: null,
    raw: true,
  });
  if (!table.length) throw new Error("Empty sheet.");

  // Find header row that contains 'Time'
  const { headerRowIndex, header } = findHeaderRowWithTime(table);
  const colIndex = new Map<string, number>();
  header.forEach((h, i) => {
    if (typeof h === "string") colIndex.set(h.trim(), i);
  });

  // Build reverse lookup that honors aliases
  function getColIndex(key: string): number {
    if (colIndex.has(key)) return colIndex.get(key)!;
    const aliased = HEADER_ALIASES.get(key);
    if (aliased && colIndex.has(aliased)) return colIndex.get(aliased)!;
    // try a few forgiving matches
    for (const [h, i] of colIndex.entries()) {
      if (h.replace(/\s+/g, "") === key.replace(/\s+/g, "")) return i;
    }
    return -1;
  }

  const timeIdx = getColIndex("Time");
  if (timeIdx < 0) throw new Error("Could not find 'Time' column.");

  const time: number[] = [];
  const buildRole = (role: Role): NeededMetrics => {
    const spec = NEEDED_MAP[role];
    const series: Series[] = spec.map(([label, key]) => ({
      label,
      key,
      values: [],
    }));

    for (let r = headerRowIndex + 1; r < table.length; r++) {
      const row = table[r] as any[];
      if (!row || !row.length) continue;

      if (role === "pitcher") {
        // Push time once; hitter will reuse same length check
        if (time.length < r - headerRowIndex) {
          const t = toNumber(row[timeIdx]);
          if (Number.isFinite(t)) time.push(t);
          else time.push(NaN);
        }
      }

      series.forEach((s) => {
        const idx = getColIndex(s.key);
        const v = idx >= 0 ? toNumber(row[idx]) : NaN;
        s.values.push(Number.isFinite(v) ? v : NaN);
      });
    }
    return { time, series };
  };

  // Build once; both roles share the same time array reference
  const pitcher = buildRole("pitcher");
  const hitter = buildRole("hitter");

  return { pitcher, hitter };
}

/* ============================================================
   Internals (header detection + normalization for legacy API)
============================================================ */

function detectHeaderAndExtract(table: any[][]): {
  headers: string[];
  dataRows: any[][];
} {
  const scanUpto = Math.min(table.length, 20);
  let bestIdx = 0;
  let bestScore = -1;

  for (let i = 0; i < scanUpto; i++) {
    const row = table[i] ?? [];
    const nonEmpty = row.filter(
      (c) => typeof c === "string" && c.trim() !== ""
    ).length;
    const hasTime = row.some(
      (c) => typeof c === "string" && /^(t|time|timestamp)$/i.test(c.trim())
    );
    const hasFrame = row.some(
      (c) => typeof c === "string" && /frame/i.test(c.trim())
    );
    let score = nonEmpty + (hasTime ? 3 : 0) + (hasFrame ? 2 : 0);
    if (nonEmpty <= 1) score -= 3; // penalize title-only rows
    if (score > bestScore) {
      bestScore = score;
      bestIdx = i;
    }
  }

  const rawHeaders = (table[bestIdx] ?? []).map((v) =>
    v == null ? "" : String(v)
  );
  const headers = dedupeHeaders(
    rawHeaders.map((h) => {
      const s = h.trim();
      if (/^frame(s)?$/i.test(s)) return "Frame";
      if (/^timestamp$/i.test(s)) return "Timestamp";
      if (/^(t|time)$/i.test(s)) return "Time";
      return s;
    })
  );

  const dataRows = table.slice(bestIdx + 1);
  return { headers, dataRows };
}

function toObjects(headers: string[], rows: any[][]): any[] {
  return rows.map((r) => {
    const obj: any = {};
    for (let i = 0; i < headers.length; i++) {
      const key = headers[i] || `Col${i}`;
      obj[key] = r[i];
    }
    return obj;
  });
}

function normalizeSheet(rowsRaw: any[], fpsGuess: number): Row[] {
  if (!rowsRaw.length) return [];

  const keys = Object.keys(rowsRaw[0] ?? {});
  const timeKey =
    keys.find((k) => /^(t|time)$/i.test(k)) ||
    keys.find((k) => /timestamp/i.test(k));
  const msKey = keys.find((k) => /(ms|millisecond)/i.test(k));
  const frameKey = keys.find((k) => /^frame(s)?$/i.test(k) || /frame ?index/i.test(k));

  const num = (v: any) => {
    if (typeof v === "number") return v;
    if (v == null) return NaN;
    const s = String(v).trim().replace(/,/g, "");
    const n = Number(s);
    return Number.isFinite(n) ? n : NaN;
  };

  // Timestamp base (if a date column exists)
  let ts0 = 0;
  if (!timeKey && keys.some((k) => /timestamp/i.test(k))) {
    const k = keys.find((k) => /timestamp/i.test(k))!;
    const first = rowsRaw.find((r) => r[k] != null)?.[k];
    const p = typeof first === "number" ? first : Date.parse(first);
    if (isFinite(p)) ts0 = p;
  }

  const out: Row[] = [];
  for (let i = 0; i < rowsRaw.length; i++) {
    const r = rowsRaw[i];
    const row: Row = {};
    let tSec = NaN;

    if (timeKey) {
      const v = r[timeKey];
      const n = num(v);
      const looksMs =
        !!msKey || (Number.isFinite(n) && n > 50 && averageDelta(rowsRaw, timeKey) > 10);
      tSec = looksMs ? n / 1000 : n;
    } else if (frameKey) {
      const f = num(r[frameKey]);
      const fps =
        fpsGuess && Number.isFinite(fpsGuess) ? fpsGuess : 120;
      tSec = Number.isFinite(f) ? f / fps : NaN;
    } else if (keys.some((k) => /timestamp/i.test(k))) {
      const k = keys.find((k) => /timestamp/i.test(k))!;
      const p =
        typeof r[k] === "number" ? (r[k] as number) : Date.parse(String(r[k]));
      tSec = Number.isFinite(p) ? (p - ts0) / 1000 : NaN;
    }

    row.t = tSec;

    for (const k of keys) {
      if (k === timeKey || k === msKey || k === frameKey) continue;
      const n = num(r[k]);
      if (Number.isFinite(n)) row[k] = n;
    }

    out.push(row);
  }

  return out;
}

function averageDelta(rows: any[], key: string): number {
  const vals: number[] = [];
  for (const r of rows) {
    const x = r[key];
    const n = typeof x === "number" ? x : Number(x);
    if (Number.isFinite(n)) vals.push(n);
  }
  if (vals.length < 2) return NaN;
  let sum = 0;
  for (let i = 1; i < vals.length; i++) sum += Math.abs(vals[i] - vals[i - 1]);
  return sum / (vals.length - 1);
}

/* ============================================================
   Small utilities used by the new API
============================================================ */

function toNumber(v: any): number {
  if (typeof v === "number") return v;
  if (v == null) return NaN;
  const s = String(v).trim().replace(/,/g, "");
  const n = Number(s);
  return Number.isFinite(n) ? n : NaN;
}

function findHeaderRowWithTime(table: any[][]): {
  headerRowIndex: number;
  header: string[];
} {
  const scanUpto = Math.min(table.length, 20);
  for (let i = 0; i < scanUpto; i++) {
    const row = table[i] ?? [];
    const hasTime = row.some(
      (c) => typeof c === "string" && /^time$/i.test(c.trim())
    );
    if (hasTime) {
      const header = (row as any[]).map((v) =>
        typeof v === "string" ? v.trim() : v
      );
      return { headerRowIndex: i, header };
    }
  }
  // Fallback to first row
  const header = (table[0] ?? []).map((v) =>
    typeof v === "string" ? v.trim() : v
  );
  return { headerRowIndex: 0, header };
}

function dedupeHeaders(headers: string[]): string[] {
  const seen = new Map<string, number>();
  return headers.map((h) => {
    if (!h) return h;
    const base = h;
    const count = (seen.get(base) ?? 0) + 1;
    seen.set(base, count);
    return count === 1 ? base : `${base} (${count})`;
  });
}
