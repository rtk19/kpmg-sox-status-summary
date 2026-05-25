// Faithful TypeScript port of main.py SOX summary builder.
// Uses ExcelJS to read the input workbook (preserving original sheets/styles)
// and append "status_N" summary sheets with formulas, fills, borders, etc.

import ExcelJS from "exceljs";

// ── Constants (mirrors main.py) ──────────────────────────────────────────────
const COLOR_HEADER_BG = "FF4F81BD";
const COLOR_HEADER_FG = "FFFFFFFF";
const COLOR_TOTAL_BG = "FFD9E1F2";
const COLOR_ROW_ALT = "FFEEF3FA";
const COLOR_ROW_NORMAL = "FFFFFFFF";

const EXCLUDED_STATUS_VALUES = new Set<string>([
  "", "לא מפתח", "not key", "IT", "it", "annual", "שנתי",
]);

const STATUS_HEADER_CANDIDATES = new Set<string>([
  "מסקנה", "בקרה", "סטטוס",
  "מסקנה סבב א'", "מסקנה סבב ב'",
  "מסקנה סבב א", "מסקנה סבב ב",
  "status", "conclusion", "control", "results",
]);

const ENGLISH_STATUS_HEADERS = new Set<string>([
  "status", "conclusion", "control", "results",
]);

// Smart matcher: anything that begins with "מסקנה"/"conclusion" (e.g. "מסקנה סבב ג",
// "מסקנה סבב 14", "conclusion round 3") is treated as a status/round column.
function isStatusHeader(text: string): boolean {
  if (!text) return false;
  if (STATUS_HEADER_CANDIDATES.has(text)) return true;
  if (/^מסקנה(\s|$)/.test(text)) return true;
  if (/^conclusion(\s|$)/.test(text)) return true;
  return false;
}

function isEnglishStatusHeader(text: string): boolean {
  if (ENGLISH_STATUS_HEADERS.has(text)) return true;
  if (/^conclusion(\s|$)/.test(text)) return true;
  return false;
}

function isHebrewStatusHeader(text: string): boolean {
  if (!text) return false;
  if (isEnglishStatusHeader(text)) return false;
  return isStatusHeader(text);
}

// Parse an explicit round number/letter from a status header
// (e.g. "מסקנה סבב ג" → 3, "conclusion round 14" → 14).
// Returns null if the header has no explicit round indicator.
const HEBREW_LETTER_VALUES: Record<string, number> = {
  "א": 1, "ב": 2, "ג": 3, "ד": 4, "ה": 5, "ו": 6, "ז": 7, "ח": 8, "ט": 9,
  "י": 10, "כ": 20, "ל": 30, "מ": 40, "נ": 50, "ס": 60, "ע": 70, "פ": 80, "צ": 90,
  "ק": 100, "ר": 200, "ש": 300, "ת": 400,
};
function hebrewLettersToNumber(letters: string): number | null {
  let total = 0;
  for (const ch of letters) {
    const v = HEBREW_LETTER_VALUES[ch];
    if (!v) return null;
    total += v;
  }
  return total > 0 ? total : null;
}
function parseRoundNumber(text: string): number | null {
  if (!text) return null;
  const digitMatch = text.match(/(\d+)/);
  if (digitMatch) return parseInt(digitMatch[1], 10);
  const hebMatch = text.match(/סבב\s+([\u05d0-\u05ea]+)'?/);
  if (hebMatch) {
    const n = hebrewLettersToNumber(hebMatch[1].replace(/'/g, ""));
    if (n) return n;
  }
  return null;
}

const KEY_PRIORITY_1 = new Set<string>(["מפתח", "key control"]);
const KEY_PRIORITY_2 = new Set<string>(["בקרת מפתח", "key"]);

const ID_HEADER_PRIORITY_EXACT = new Set<string>([
  "מס", "מס'", "מספר", "מס בקרה", "מס' בקרה", "מספר בקרה",
  "control number", "control no", "id", "no", "no.", ".no",
]);

const ID_HEADER_CONTAINS = [
  "מס", "מס'", "מספר", "control number", "control no", "id", "no", "no.", ".no",
];

const NON_ID_HEADERS_EXACT = new Set<string>([
  "מפתח", "בקרת מפתח", "מסקנה", "סטטוס", "בקרה", "key control",
]);

const PROCESS_HEADER_PRIMARY = new Set<string>(["process", "תהליך"]);
const PROCESS_HEADER_SECONDARY = new Set<string>(["process name", "שם התהליך"]);
const PROCESS_HEADER_EXCLUDE_TOKENS = ["sub", "sub-", "sub_", "משנה", "תת"];

type Lang = "he" | "en";

const LABELS: Record<Lang, Record<string, string>> = {
  he: {
    sheet_name: "גיליון",
    process_name: "שם התהליך",
    round: "סבב",
    total_controls: 'סה"כ בקרות',
    key_control: "בקרת מפתח",
    total_statuses: 'סה"כ סטטוסים',
    it_controls: "בקרות IT",
    annual: "שנתי",
    exec_pct: "אחוז ביצוע",
    total_row: 'סה"כ',
  },
  en: {
    sheet_name: "Sheet",
    process_name: "Process Name",
    round: "Round",
    total_controls: "Total Controls",
    key_control: "Key Control",
    total_statuses: "Total Statuses",
    it_controls: "IT Controls",
    annual: "Annual",
    exec_pct: "Execution %",
    total_row: "Total",
  },
};

// ── Errors ───────────────────────────────────────────────────────────────────
export class SoxBusinessStop extends Error {}
export class SoxClarificationRequired extends SoxBusinessStop {}
export class SoxStructureNotIdentified extends SoxBusinessStop {}

// ── Helpers ──────────────────────────────────────────────────────────────────
function colLetter(n: number): string {
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
function colIndex(letter: string): number {
  let n = 0;
  for (const ch of letter.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

function stripInvisible(s: string): string {
  // Remove NBSP, zero-width chars, BOM, then trim regular whitespace
  return s.replace(/[\u00a0\u200b\u200c\u200d\ufeff]/g, " ").trim();
}
function cleanText(v: unknown): string {
  if (v === null || v === undefined) return "";
  // ExcelJS may return { richText: [...] } or { formula, result }
  if (typeof v === "object") {
    const anyV = v as Record<string, unknown>;
    if ("richText" in anyV && Array.isArray(anyV.richText)) {
      return stripInvisible(
        (anyV.richText as Array<{ text?: string }>)
          .map((r) => r.text ?? "")
          .join("")
      );
    }
    if ("result" in anyV) return cleanText(anyV.result);
    if ("text" in anyV) return cleanText(anyV.text);
    if ("hyperlink" in anyV && "text" in anyV) return cleanText((anyV as { text?: unknown }).text);
  }
  return stripInvisible(String(v));
}
function cleanHeader(v: unknown): string {
  return cleanText(v).replace(/\s+/g, " ").toLowerCase();
}
function isEmpty(v: unknown): boolean { return cleanText(v) === ""; }
function escapeExcel(s: string): string { return s.replace(/"/g, '""'); }

function looksLikeHeaderCell(text: string): boolean {
  if (text === "") return false;
  // "digits_only" iff no unicode letter present
  const hasLetter = /\p{L}/u.test(text);
  return hasLetter;
}

function getMaxRow(ws: ExcelJS.Worksheet): number {
  return Math.max(ws.actualRowCount || 0, ws.rowCount || 0);
}
function getMaxCol(ws: ExcelJS.Worksheet): number {
  // actualColumnCount can under-report when later rows have wider data than the header.
  // Scan all rows to find the true max populated column.
  let maxCol = Math.max(ws.actualColumnCount || 0, ws.columnCount || 0);
  const rowCount = Math.max(ws.actualRowCount || 0, ws.rowCount || 0);
  for (let r = 1; r <= rowCount; r++) {
    const row = ws.getRow(r);
    const cellCount = row.cellCount || 0;
    if (cellCount > maxCol) maxCol = cellCount;
  }
  return maxCol;
}

function cellValue(ws: ExcelJS.Worksheet, row: number, col: number): unknown {
  return ws.getCell(row, col).value;
}
function rowValues(ws: ExcelJS.Worksheet, row: number, maxCol: number): string[] {
  const out: string[] = [];
  for (let c = 1; c <= maxCol; c++) out.push(cleanText(cellValue(ws, row, c)));
  return out;
}
function countNonEmptyInRow(ws: ExcelJS.Worksheet, row: number, maxCol: number): number {
  return rowValues(ws, row, maxCol).filter((v) => v !== "").length;
}

function normalizeStatusValues(values: string[]): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const v of values) {
    if (v === null || v === undefined) continue;
    const s = String(v).trim();
    if (EXCLUDED_STATUS_VALUES.has(s)) continue;
    if (!seen.has(s)) { seen.add(s); out.push(s); }
  }
  return out;
}

// ── Structure detection ─────────────────────────────────────────────────────
function scoreHeaderRow(ws: ExcelJS.Worksheet, row: number, maxCol: number): number {
  const values = rowValues(ws, row, maxCol);
  const nonEmpty = values.filter((v) => v !== "");
  if (nonEmpty.length < 2) return -1;

  const headerLike = nonEmpty.filter(looksLikeHeaderCell).length;
  const exactStatus = nonEmpty.filter((v) => isStatusHeader(cleanHeader(v))).length;
  const keyHits = nonEmpty.filter((v) => {
    const t = cleanHeader(v);
    return KEY_PRIORITY_1.has(t) || KEY_PRIORITY_2.has(t) ||
      (t.includes("מפתח") && !t.includes("לא מפתח"));
  }).length;
  const idHits = nonEmpty.filter((v) => {
    const t = cleanHeader(v);
    if (ID_HEADER_PRIORITY_EXACT.has(t)) return true;
    return ID_HEADER_CONTAINS.some((tok) => t.toLowerCase().includes(tok.toLowerCase()));
  }).length;

  const maxRow = getMaxRow(ws);
  const nextRowNonEmpty = row < maxRow ? countNonEmptyInRow(ws, row + 1, maxCol) : 0;

  return nonEmpty.length * 3 + headerLike * 2 + exactStatus * 6 + keyHits * 5 + idHits * 4 + Math.min(nextRowNonEmpty, 10);
}

function findHeaderRow(ws: ExcelJS.Worksheet): number {
  const maxCol = getMaxCol(ws);
  const maxRow = getMaxRow(ws);
  const scanUntil = Math.min(maxRow, 30);
  let bestRow: number | null = null;
  let bestScore = -1;
  for (let r = 1; r <= scanUntil; r++) {
    const s = scoreHeaderRow(ws, r, maxCol);
    if (s > bestScore) { bestScore = s; bestRow = r; }
  }
  if (bestRow === null || bestScore < 0) {
    throw new SoxStructureNotIdentified(
      `לא זוהתה שורת כותרת בגיליון '${ws.name}'.\n` +
      `איך לתקן: ודא שב-30 השורות הראשונות קיימת שורה אחת לפחות עם כותרות עמודות (לפחות שתי עמודות עם טקסט).\n` +
      `דוגמאות לכותרות מזוהות: מסקנה, סטטוס, בקרה, מסקנה סבב א', מפתח, בקרת מפתח, מס' בקרה, תהליך (או באנגלית: status, conclusion, control, key control, control number, process).`
    );
  }
  return bestRow;
}

function findTableBounds(ws: ExcelJS.Worksheet, headerRow: number): number {
  const maxRow = getMaxRow(ws);
  const firstDataRow = headerRow + 1;
  const noDataMsg =
    `לא נמצאו שורות נתונים מתחת לכותרת בגיליון '${ws.name}'.\n` +
    `איך לתקן: ודא שמתחת לשורת הכותרת קיימת לפחות שורה אחת עם ערך באחת מ-3 העמודות הראשונות (A/B/C). ` +
    `בלוקים של הערות (כמו "Control Objectives") שמופיעים רק מעמודה D ואילך לא נחשבים כשורות נתונים.`;
  if (firstDataRow > maxRow) {
    throw new SoxStructureNotIdentified(noDataMsg);
  }
  for (let r = firstDataRow; r <= maxRow; r++) {
    const aEmpty = isEmpty(cellValue(ws, r, 1));
    const bEmpty = isEmpty(cellValue(ws, r, 2));
    const cEmpty = isEmpty(cellValue(ws, r, 3));
    if (aEmpty && bEmpty && cEmpty) {
      if (r === firstDataRow) {
        throw new SoxStructureNotIdentified(noDataMsg);
      }
      return r - 1;
    }
  }
  return maxRow;
}

function findStatusColumns(
  ws: ExcelJS.Worksheet,
  headerRow: number,
): { letters: string[]; roundNumbers: (number | null)[] } {
  const letters: string[] = [];
  const roundNumbers: (number | null)[] = [];
  const maxCol = getMaxCol(ws);
  for (let c = 1; c <= maxCol; c++) {
    const text = cleanHeader(cellValue(ws, headerRow, c));
    if (isStatusHeader(text)) {
      letters.push(colLetter(c));
      roundNumbers.push(parseRoundNumber(text));
    }
  }
  // Natural left-to-right order; round numbers parsed from header text when present.
  return { letters, roundNumbers };
}

function findKeyColumn(ws: ExcelJS.Worksheet, headerRow: number, overrides: Record<string, string>): string | null {
  if (ws.name in overrides) return overrides[ws.name];
  const headers: Array<[string, string]> = [];
  const maxCol = getMaxCol(ws);
  for (let c = 1; c <= maxCol; c++) {
    const letter = colLetter(c);
    const text = cleanHeader(cellValue(ws, headerRow, c));
    if (text !== "") headers.push([letter, text]);
  }
  const p1 = headers.filter(([, t]) => KEY_PRIORITY_1.has(t));
  const p2 = headers.filter(([, t]) => KEY_PRIORITY_2.has(t));
  const p3 = headers.filter(([, t]) =>
    t.includes("מפתח") && !t.includes("לא מפתח") && !KEY_PRIORITY_1.has(t) && !KEY_PRIORITY_2.has(t)
  );
  for (const matches of [p1, p2, p3]) {
    if (matches.length === 1) return matches[0][0];
    if (matches.length > 1) {
      const lines = [`נמצאו מספר עמודות אפשריות לעמודת המפתח בגיליון '${ws.name}':`];
      for (const [letter, text] of matches) lines.push(`  • עמודה ${letter} — "${text}"`);
      lines.push("", "איך לתקן: השאר רק כותרת אחת שמכילה 'מפתח' / 'key control', או שנה את שמות שאר העמודות.");
      throw new SoxClarificationRequired(lines.join("\n"));
    }
  }
  return null;
}

function findIdColumn(
  ws: ExcelJS.Worksheet,
  headerRow: number,
  statusCols: Set<string>,
  keyCol: string | null,
): string | null {
  const headers: Array<{ c: number; letter: string; text: string }> = [];
  const maxCol = getMaxCol(ws);
  for (let c = 1; c <= maxCol; c++) {
    headers.push({ c, letter: colLetter(c), text: cleanHeader(cellValue(ws, headerRow, c)) });
  }
  const exactLower = new Set([...ID_HEADER_PRIORITY_EXACT].map((s) => s.toLowerCase()));
  const exact = headers.filter((h) => exactLower.has(h.text.toLowerCase()));
  if (exact.length >= 1) return exact[0].letter;

  const contains: typeof headers = [];
  for (const h of headers) {
    if (h.text === "") continue;
    if (statusCols.has(h.letter)) continue;
    if (keyCol && h.letter === keyCol) continue;
    if (NON_ID_HEADERS_EXACT.has(h.text)) continue;
    const lower = h.text.toLowerCase();
    if (ID_HEADER_CONTAINS.some((tok) => lower.includes(tok.toLowerCase()))) contains.push(h);
  }
  if (contains.length >= 1) return contains[0].letter;
  return null;
}

function findProcessColumn(ws: ExcelJS.Worksheet, headerRow: number): string | null {
  const maxCol = getMaxCol(ws);
  const primary: string[] = [];
  const secondary: string[] = [];
  for (let c = 1; c <= maxCol; c++) {
    const text = cleanHeader(cellValue(ws, headerRow, c));
    if (text === "") continue;
    // Reject any header containing "sub" / "משנה" / "תת" so a "Sub Process" column
    // is never picked up as the process column.
    if (PROCESS_HEADER_EXCLUDE_TOKENS.some((tok) => text.includes(tok))) continue;
    if (PROCESS_HEADER_PRIMARY.has(text)) primary.push(colLetter(c));
    else if (PROCESS_HEADER_SECONDARY.has(text)) secondary.push(colLetter(c));
  }
  if (primary.length) return primary[0];
  if (secondary.length) return secondary[0];
  return null;
}

interface SheetConfig {
  start_row: number;
  end_row: number;
  id_col: string | null;
  key_col: string | null;
  process_col: string | null;
  status_cols: string[];
  round_numbers: number[]; // 1-based round number per status column (parallel to status_cols)
}

function detectLanguage(wb: ExcelJS.Workbook, cfg: Map<string, SheetConfig>): Lang {
  for (const [name, c] of cfg) {
    if (!c.status_cols.length) continue;
    const ws = wb.getWorksheet(name);
    if (!ws) continue;
    const maxCol = getMaxCol(ws);
    for (let col = 1; col <= maxCol; col++) {
      const t = cleanHeader(cellValue(ws, c.start_row, col));
      if (isEnglishStatusHeader(t)) return "en";
      if (isHebrewStatusHeader(t)) return "he";
    }
  }
  return "he";
}

function analyzeWorkbook(
  wb: ExcelJS.Workbook,
  manualOverrides: Record<string, string>,
): { sheetConfig: Map<string, SheetConfig>; lang: Lang } {
  const visible = wb.worksheets.filter((ws) => (ws.state ?? "visible") === "visible");
  const emptyFileMsg =
    "לא נמצא בקובץ אף גיליון גלוי עם נתונים.\n" +
    "איך לתקן: ודא שהקובץ מכיל לפחות גיליון אחד שאינו מוסתר (Hidden) ושיש בו טבלת בקרות עם שורת כותרת וערכים.";
  if (!visible.length) throw new SoxStructureNotIdentified(emptyFileMsg);

  const sheetConfig = new Map<string, SheetConfig>();
  const skipped: Array<{ name: string; reason: string }> = [];
  for (const ws of visible) {
    const maxRow = getMaxRow(ws);
    const maxCol = getMaxCol(ws);
    if (maxRow <= 1 && maxCol <= 1 && isEmpty(cellValue(ws, 1, 1))) continue;

    try {
      const headerRow = findHeaderRow(ws);
      const { letters: statusCols, roundNumbers: parsedRoundNumbers } = findStatusColumns(ws, headerRow);
      if (statusCols.length === 0) {
        skipped.push({
          name: ws.name,
          reason: "לא נמצאה עמודת מסקנה/סטטוס בגיליון זה.",
        });
        continue;
      }
      const endRow = findTableBounds(ws, headerRow);
      const keyCol = findKeyColumn(ws, headerRow, manualOverrides);
      const idCol = findIdColumn(ws, headerRow, new Set(statusCols), keyCol);
      const processCol = findProcessColumn(ws, headerRow);

      sheetConfig.set(ws.name, {
        start_row: headerRow,
        end_row: endRow,
        id_col: idCol,
        key_col: keyCol,
        process_col: processCol,
        status_cols: statusCols,
        // Temporarily store parsed numbers (nulls allowed); finalized below once
        // language is known so positional fallbacks respect reading direction.
        round_numbers: parsedRoundNumbers as unknown as number[],
      });
    } catch (e) {
      // Skip sheets whose structure cannot be identified; keep processing the rest.
      // Clarification errors (multiple key columns) still propagate — they need user input.
      if (e instanceof SoxClarificationRequired) throw e;
      if (e instanceof SoxBusinessStop) {
        const firstLine = e.message.split("\n")[0];
        skipped.push({ name: ws.name, reason: firstLine });
        continue;
      }
      throw e;
    }
  }

  if (!sheetConfig.size) {
    const skippedList = skipped.length
      ? "\n\nגיליונות שנדלגו:\n" + skipped.map((s) => `  • '${s.name}' — ${s.reason}`).join("\n")
      : "";
    throw new SoxStructureNotIdentified(
      `לא נמצא בקובץ אף גיליון תקין לעיבוד.\n` +
      `איך לתקן: ודא שלפחות גיליון אחד מכיל שורת כותרת עם עמודת מסקנה/סטטוס ` +
      `(מסקנה, סטטוס, בקרה, status, conclusion, control, results) ושורות נתונים מתחתיה.` +
      skippedList
    );
  }

  if (skipped.length) {
    console.warn(
      "[SOX] גיליונות שנדלגו בעיבוד:\n" +
        skipped.map((s) => `  • '${s.name}' — ${s.reason}`).join("\n")
    );
  }

  const lang = detectLanguage(wb, sheetConfig);
  // Round numbering always follows column order (A, B, C, ...) so all rounds appear
  // in a stable, predictable sequence regardless of header text or language.
  for (const cfg of sheetConfig.values()) {
    cfg.round_numbers = cfg.status_cols.map((_, i) => i + 1);
  }
  return { sheetConfig, lang };
}

function extractAllDynamicStatusValues(
  wb: ExcelJS.Workbook,
  sheetConfig: Map<string, SheetConfig>,
): string[] {
  const values: string[] = [];
  for (const [name, cfg] of sheetConfig) {
    const ws = wb.getWorksheet(name);
    if (!ws) continue;
    for (const col of cfg.status_cols) {
      const colIdx = colIndex(col);
      for (let r = cfg.start_row + 1; r <= cfg.end_row; r++) {
        const text = cleanText(cellValue(ws, r, colIdx));
        if (text === "") continue;
        if (EXCLUDED_STATUS_VALUES.has(text)) continue;
        values.push(text);
      }
    }
  }
  return normalizeStatusValues(values);
}

function getProcessValues(ws: ExcelJS.Worksheet, cfg: SheetConfig): string[] {
  if (!cfg.process_col) return [];
  const colIdx = colIndex(cfg.process_col);
  const seen = new Set<string>();
  const out: string[] = [];
  for (let r = cfg.start_row + 1; r <= cfg.end_row; r++) {
    const text = cleanText(cellValue(ws, r, colIdx));
    if (text === "") continue;
    if (!seen.has(text)) { seen.add(text); out.push(text); }
  }
  return out;
}

function getMaxStatusCount(sheetConfig: Map<string, SheetConfig>): number {
  let m = 0;
  for (const cfg of sheetConfig.values()) m = Math.max(m, cfg.status_cols.length);
  return m;
}

// ── Styling helpers ─────────────────────────────────────────────────────────
const thin: Partial<ExcelJS.Border> = { style: "thin" };
const thinBorder: Partial<ExcelJS.Borders> = { top: thin, left: thin, right: thin, bottom: thin };

function styleHeaderCell(cell: ExcelJS.Cell) {
  cell.font = { bold: true, color: { argb: COLOR_HEADER_FG } };
  cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLOR_HEADER_BG } };
  cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  cell.border = { ...thinBorder };
}

function setFormula(cell: ExcelJS.Cell, formula: string) {
  // Strip newlines/extra spaces from multi-line python strings.
  const f = formula.replace(/\s+/g, " ").trim().replace(/^=/, "");
  cell.value = { formula: f } as ExcelJS.CellFormulaValue;
}

// ── Summary sheet builder ───────────────────────────────────────────────────
const SUMMARY_SHEET_NAME: Record<Lang, string> = { he: "סיכום", en: "Summary" };

interface SummaryRow {
  sheet: string;
  process: string | null;   // null when no process column on the sheet
  statusIndex: number;      // 0-based round index
  cfg: SheetConfig;
}

function buildSummaryRows(
  wb: ExcelJS.Workbook,
  sheetConfig: Map<string, SheetConfig>,
): SummaryRow[] {
  const rows: SummaryRow[] = [];
  for (const [name, cfg] of sheetConfig) {
    const ws = wb.getWorksheet(name);
    if (!ws) continue;
    const processes = getProcessValues(ws, cfg);
    const processColIdx = cfg.process_col ? colIndex(cfg.process_col) : 0;

    // Check whether a given status column has any non-empty value for the given
    // process (or for the whole sheet when there is no process column / process).
    const hasAnyStatusValue = (statusCol: string, process: string | null): boolean => {
      const colIdx = colIndex(statusCol);
      for (let r = cfg.start_row + 1; r <= cfg.end_row; r++) {
        if (process && processColIdx) {
          const procText = cleanText(cellValue(ws, r, processColIdx));
          if (procText !== process) continue;
        }
        const text = cleanText(cellValue(ws, r, colIdx));
        if (text !== "") return true;
      }
      return false;
    };

    for (let idx = 0; idx < cfg.status_cols.length; idx++) {
      const statusCol = cfg.status_cols[idx];
      if (processes.length) {
        for (const p of processes) {
          // Skip rounds with no data for this specific process — the control
          // was not performed in that round for that process.
          if (!hasAnyStatusValue(statusCol, p)) continue;
          rows.push({ sheet: name, process: p, statusIndex: idx, cfg });
        }
      } else {
        if (!hasAnyStatusValue(statusCol, null)) continue;
        rows.push({ sheet: name, process: null, statusIndex: idx, cfg });
      }
    }
  }
  return rows;
}

function createSummarySheet(
  wb: ExcelJS.Workbook,
  sheetConfig: Map<string, SheetConfig>,
  dynamicValues: string[],
  lang: Lang,
) {
  const labels = LABELS[lang];
  const filtered = normalizeStatusValues(dynamicValues);
  const summaryRows = buildSummaryRows(wb, sheetConfig);

  const maxRounds = getMaxStatusCount(sheetConfig);
  const showRound = maxRounds > 1;
  const showProcess = [...sheetConfig.values()].some((c) => c.process_col);

  const sheetName = SUMMARY_SHEET_NAME[lang];
  const existing = wb.getWorksheet(sheetName);
  if (existing) wb.removeWorksheet(existing.id);
  const ws = wb.addWorksheet(sheetName);

  ws.views = [{
    state: "frozen", xSplit: 0, ySplit: 1,
    rightToLeft: lang === "he",
  }];

  const headers: string[] = [labels.sheet_name];
  if (showProcess) headers.push(labels.process_name);
  if (showRound) headers.push(labels.round);
  headers.push(labels.total_controls, labels.key_control,
    ...filtered, labels.total_statuses, labels.it_controls,
    labels.annual, labels.exec_pct);

  const headerRow = ws.getRow(1);
  headers.forEach((h, i) => {
    const cell = headerRow.getCell(i + 1);
    cell.value = h;
    styleHeaderCell(cell);
  });
  headerRow.height = 30;
  headerRow.commit?.();

  // Column positions
  let col = 1;
  const colSheet = col++;
  const colProcess = showProcess ? col++ : 0;
  const colRound = showRound ? col++ : 0;
  const colTotal = col++;
  const colKey = col++;
  const colStatusStart = col;
  col += filtered.length;
  const colStatusEnd = filtered.length ? colStatusStart + filtered.length - 1 : colStatusStart - 1;
  const colRowSum = col++;
  const colIt = col++;
  const colAnnual = col++;
  const colExec = col++;

  const passIdx = filtered.findIndex((v) => {
    const x = v.trim().toLowerCase();
    return x === "עבר" || x === "pass";
  });
  const passColLetter = passIdx >= 0 ? colLetter(colStatusStart + passIdx) : null;
  const colKeyLetter = colLetter(colKey);
  const colItLetter = colLetter(colIt);
  const colAnnualLetter = colLetter(colAnnual);

  // Helper to build a process filter sub-expression (or empty if no process)
  const procFilter = (cfg: SheetConfig, sheet: string, process: string | null): string => {
    if (!process || !cfg.process_col) return "";
    const range = `'${sheet}'!${cfg.process_col}${cfg.start_row + 1}:${cfg.process_col}${cfg.end_row}`;
    return `(TRIM(${range})="${escapeExcel(process)}")*`;
  };

  let curr = 2;
  let wroteData = false;

  for (const r of summaryRows) {
    const { sheet, process, statusIndex, cfg } = r;
    const statusCol = cfg.status_cols[statusIndex];
    const start = cfg.start_row + 1;
    const end = cfg.end_row;
    const statusRange = `'${sheet}'!${statusCol}${start}:${statusCol}${end}`;
    const pf = procFilter(cfg, sheet, process);

    const row = ws.getRow(curr);
    row.getCell(colSheet).value = sheet;
    if (showProcess) row.getCell(colProcess).value = process ?? "";
    if (showRound) row.getCell(colRound).value = cfg.round_numbers[statusIndex] ?? statusIndex + 1;

    // Total controls
    if (cfg.id_col) {
      const idRange = `'${sheet}'!${cfg.id_col}${start}:${cfg.id_col}${end}`;
      setFormula(row.getCell(colTotal),
        `=SUMPRODUCT(${pf}(--ISNUMBER(--RIGHT(${idRange},1))))`);
    } else {
      row.getCell(colTotal).value = 0;
    }

    // Key control
    if (cfg.key_col) {
      const keyRange = `'${sheet}'!${cfg.key_col}${start}:${cfg.key_col}${end}`;
      if (pf) {
        setFormula(row.getCell(colKey),
          `=SUMPRODUCT(${pf}((${keyRange}="כן")+(${keyRange}="Yes")+(${keyRange}="Y / כן")))`);
      } else {
        setFormula(row.getCell(colKey),
          `=COUNTIF(${keyRange},"כן")+COUNTIF(${keyRange},"Yes")+COUNTIF(${keyRange},"Y / כן")`);
      }
    } else {
      row.getCell(colKey).value = 0;
    }

    // Status value counts
    filtered.forEach((val, i) => {
      setFormula(row.getCell(colStatusStart + i),
        `=SUMPRODUCT(${pf}(--(TRIM(${statusRange})="${escapeExcel(val)}")))`);
    });

    // Row sum of status values
    if (filtered.length) {
      const dStart = colLetter(colStatusStart);
      const dEnd = colLetter(colStatusEnd);
      setFormula(row.getCell(colRowSum), `=SUM(${dStart}${curr}:${dEnd}${curr})`);
    } else {
      row.getCell(colRowSum).value = 0;
    }

    // IT
    setFormula(row.getCell(colIt),
      `=SUMPRODUCT(${pf}(--ISNUMBER(SEARCH("IT",TRIM(${statusRange})))))`);

    // Annual
    setFormula(row.getCell(colAnnual),
      `=SUMPRODUCT(${pf}(--((TRIM(${statusRange})="שנתי")+(TRIM(${statusRange})="Annual"))))`);

    // Exec %
    if (passColLetter) {
      setFormula(row.getCell(colExec),
        `=IFERROR(${passColLetter}${curr}/MAX(${passColLetter}${curr},${colKeyLetter}${curr}-${colItLetter}${curr}-${colAnnualLetter}${curr}),0)`);
    } else {
      row.getCell(colExec).value = 0;
    }

    const fillColor = curr % 2 === 0 ? COLOR_ROW_ALT : COLOR_ROW_NORMAL;
    for (let c = 1; c <= colExec; c++) {
      const cell = row.getCell(c);
      cell.border = { ...thinBorder };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: fillColor } };
      cell.alignment = {
        horizontal: c <= (showProcess ? colProcess : colSheet) ? "right" : "center",
        vertical: "middle",
      };
    }

    wroteData = true;
    curr++;
  }

  // Total row
  const totalRow = ws.getRow(curr);
  const labelCell = totalRow.getCell(1);
  labelCell.value = labels.total_row;
  labelCell.font = { bold: true };
  labelCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLOR_TOTAL_BG } };
  labelCell.border = { ...thinBorder };
  labelCell.alignment = { horizontal: "right", vertical: "middle" };

  // Blank styled cells for sheet/process/round identifier columns
  for (let c = 2; c < colTotal; c++) {
    const cell = totalRow.getCell(c);
    cell.font = { bold: true };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLOR_TOTAL_BG } };
    cell.border = { ...thinBorder };
    cell.alignment = { horizontal: "center", vertical: "middle" };
  }

  for (let c = colTotal; c < colExec; c++) {
    const letter = colLetter(c);
    const cell = totalRow.getCell(c);
    setFormula(cell, `=SUM(${letter}2:${letter}${curr - 1})`);
    cell.font = { bold: true };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLOR_TOTAL_BG } };
    cell.border = { ...thinBorder };
    cell.alignment = { horizontal: "center", vertical: "middle" };
  }

  const finalExec = totalRow.getCell(colExec);
  if (passColLetter && wroteData) {
    setFormula(finalExec,
      `=IFERROR(${passColLetter}${curr}/MAX(${passColLetter}${curr},${colKeyLetter}${curr}-${colItLetter}${curr}-${colAnnualLetter}${curr}),0)`);
  } else {
    finalExec.value = 0;
  }
  finalExec.font = { bold: true };
  finalExec.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLOR_TOTAL_BG } };
  finalExec.border = { ...thinBorder };
  finalExec.alignment = { horizontal: "center", vertical: "middle" };
  finalExec.numFmt = "0%";

  for (let r = 2; r <= curr; r++) {
    ws.getCell(r, colExec).numFmt = "0%";
  }

  for (let c = 1; c <= colExec; c++) {
    let maxLen = 0;
    for (let r = 1; r <= curr; r++) {
      const v = ws.getCell(r, c).value;
      let s = "";
      if (v === null || v === undefined) s = "";
      else if (typeof v === "object" && v !== null && "formula" in (v as object)) {
        s = String((v as ExcelJS.CellFormulaValue).formula);
      } else s = String(v);
      if (s.length > maxLen) maxLen = s.length;
    }
    ws.getColumn(c).width = Math.min(maxLen + 4, 30);
  }
}

// ── Public entry point ──────────────────────────────────────────────────────
export interface ProcessResult {
  blob: Blob;
  filename: string;
}

function sanitizeSheetName(name: string): string {
  return name.replace(/[\[\]:*?/\\]/g, "_").slice(0, 31);
}

function uniqueSheetName(base: string, used: Set<string>): string {
  let candidate = sanitizeSheetName(base);
  let i = 2;
  while (used.has(candidate)) {
    const suffix = ` (${i++})`;
    candidate = sanitizeSheetName(base).slice(0, 31 - suffix.length) + suffix;
  }
  return candidate;
}

function copySheet(src: ExcelJS.Worksheet, destWb: ExcelJS.Workbook, name: string): ExcelJS.Worksheet {
  const dest = destWb.addWorksheet(name, {
    views: src.views ? JSON.parse(JSON.stringify(src.views)) : undefined,
  });
  src.columns?.forEach((c, i) => {
    if (c?.width) dest.getColumn(i + 1).width = c.width;
  });
  src.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    const destRow = dest.getRow(rowNumber);
    if (row.height) destRow.height = row.height;
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const destCell = destRow.getCell(colNumber);
      destCell.value = cell.value as ExcelJS.CellValue;
      if (cell.style) destCell.style = JSON.parse(JSON.stringify(cell.style));
      if (cell.numFmt) destCell.numFmt = cell.numFmt;
    });
    destRow.commit();
  });
  const merges = (src as unknown as { model?: { merges?: string[] } }).model?.merges;
  merges?.forEach((m) => {
    try { dest.mergeCells(m); } catch { /* ignore */ }
  });
  return dest;
}

async function loadWorkbook(file: File): Promise<ExcelJS.Workbook> {
  const buf = await file.arrayBuffer();
  const wb = new ExcelJS.Workbook();
  try {
    await wb.xlsx.load(buf);
  } catch (e) {
    const msg = String((e as Error)?.message ?? e);
    if (/old \.xls|BIFF/i.test(msg)) {
      throw new SoxBusinessStop(
        `הקובץ '${file.name}' הוא בפורמט xls הישן ואינו נתמך.\n` +
        `איך לתקן: פתחו אותו ב-Excel ובחרו "Save As" → סוג קובץ "Excel Workbook (.xlsx)", ואז העלו שוב.`
      );
    }
    throw new SoxBusinessStop(
      `הקובץ '${file.name}' אינו קובץ xlsx תקין.\n` +
      `איך לתקן: ודא שסיומת הקובץ היא .xlsx ושהוא אינו פגום או מוגן בסיסמה. פרטים: ${msg}`
    );
  }
  return wb;
}

export async function processSoxFile(
  file: File,
  manualOverrides: Record<string, string> = {},
): Promise<ProcessResult> {
  return processSoxFiles([file], manualOverrides);
}

export async function processSoxFiles(
  files: File[],
  manualOverrides: Record<string, string> = {},
): Promise<ProcessResult> {
  if (!files.length) throw new SoxBusinessStop("לא נבחרו קבצים");

  const combined = new ExcelJS.Workbook();
  const used = new Set<string>();
  const langs = new Set<Lang>();

  for (const file of files) {
    const wb = await loadWorkbook(file);

    // Detect language per source file (before merging)
    try {
      const { lang: fileLang } = analyzeWorkbook(wb, manualOverrides);
      langs.add(fileLang);
    } catch {
      // Ignore — will surface again when combined is analyzed
    }

    const stem = file.name.replace(/\.xlsx$/i, "");
    for (const ws of wb.worksheets) {
      if ((ws.state ?? "visible") !== "visible") continue;
      if (/^(status_\d+|סיכום|summary)$/i.test(ws.name)) continue;

      let base = ws.name;
      if (files.length > 1 && used.has(sanitizeSheetName(base))) {
        base = `${stem} - ${ws.name}`;
      }
      const name = uniqueSheetName(base, used);
      used.add(name);
      copySheet(ws, combined, name);
    }
  }

  if (!combined.worksheets.length) {
    throw new SoxStructureNotIdentified(
      "לא נמצא בקובץ אף גיליון גלוי עם נתונים.\n" +
      "איך לתקן: ודא שהקובץ מכיל לפחות גיליון אחד שאינו מוסתר ושיש בו טבלת בקרות."
    );
  }

  const { sheetConfig, lang: detectedLang } = analyzeWorkbook(combined, manualOverrides);
  // If sources mixed languages, default the summary to English
  const lang: Lang = langs.size > 1 ? "en" : detectedLang;

  for (const ws of [...combined.worksheets]) {
    if (/^status_\d+$/i.test(ws.name)) combined.removeWorksheet(ws.id);
  }

  const dyn = extractAllDynamicStatusValues(combined, sheetConfig);
  createSummarySheet(combined, sheetConfig, dyn, lang);

  const out = await combined.xlsx.writeBuffer();
  const filename = files.length === 1
    ? `${files[0].name.replace(/\.xlsx$/i, "").replace(/\s+/g, "_")}_with_status.xlsx`
    : `sox_summary_of_controls_combined_with_status.xlsx`;
  const blob = new Blob([out], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  return { blob, filename };
}


// ── Hebrew error classification ─────────────────────────────────────────────
export function classifyError(err: unknown): string {
  // Business errors already carry a detailed, user-friendly message.
  if (err instanceof SoxBusinessStop) return err.message;

  const msg = err instanceof Error ? err.message : String(err);
  if (msg.includes("No sheet") || msg.includes("Worksheet")) {
    return "הקובץ אינו מכיל גיליונות תקינים.\nאיך לתקן: ודא שהקובץ מכיל לפחות גיליון אחד גלוי עם נתונים.";
  }
  if (msg.includes("not an Excel file") || msg.toLowerCase().includes("zip")) {
    return "הקובץ אינו קובץ Excel תקין.\nאיך לתקן: ודא שהפורמט הוא .xlsx ושהקובץ אינו פגום.";
  }
  return `שגיאה לא צפויה בעיבוד הקובץ.\nפרטים טכניים: ${msg}`;
}
