# Codex Project Notes

Last reviewed: 2026-05-27

## What This App Does

This is a Hebrew RTL web tool for KPMG SOX status summaries. Users upload one or more `.xlsx` SOX control matrix files, and the browser generates a new `.xlsx` workbook that:

- preserves the original visible worksheets,
- skips existing generated summary/status sheets,
- detects each sheet's SOX table structure,
- appends a localized summary sheet named `סיכום` or `Summary`,
- downloads the generated workbook automatically.

The core business logic is client-side in `src/lib/soxProcessor.ts` using `exceljs`.

## Tech Stack

- Lovable-generated TanStack Start app.
- Vite dev/build through `@lovable.dev/vite-tanstack-config`.
- React 19, TypeScript, TanStack Router/Start, TanStack Query.
- Tailwind CSS v4 with shadcn-style UI components under `src/components/ui`.
- Excel processing via `exceljs`.
- Cloudflare-oriented config in `wrangler.jsonc`.
- Package manager appears to be Bun (`bun.lock`, `bunfig.toml`).

Useful commands:

```sh
bun install
bun run dev
bun run build
bun run lint
bun run format
```

`package.json` also supports `npm`-style scripts, but the lockfile is Bun.

## Important Files

- `src/routes/index.tsx`: only real app page. Handles upload UI, drag/drop, progress simulation, errors, download, and help dialog.
- `src/lib/soxProcessor.ts`: workbook analysis, sheet copying, summary generation, Hebrew error messages.
- `src/routes/__root.tsx`: root route, metadata, RTL HTML shell, QueryClient provider, route-level error/404 UI.
- `src/server.ts`: custom SSR wrapper that catches/normalizes catastrophic server errors into branded HTML.
- `src/start.ts`: TanStack Start middleware that catches SSR request errors.
- `src/styles.css`: Tailwind v4 setup, Heebo font faces, KPMG-ish color tokens, RTL base styles.
- `vite.config.ts`: Lovable config wrapper. Do not manually add plugins already supplied by `@lovable.dev/vite-tanstack-config`.
- `.lovable/plan.md`: previous Lovable implementation plan for the help-dialog tabs.

## UI Flow

`src/routes/index.tsx` has three views:

- `upload`: drag/drop or file picker, validates `.xlsx`, shows selected files, submit button.
- `loading`: fake progress ticker while `processSoxFiles(files)` runs.
- `success`: completion state with redownload and reset actions.

Files are deduplicated by `name_size`. Only `.xlsx` extension is accepted. On success, the app creates an object URL, triggers an immediate download, and stores the URL/name so the user can redownload. Object URLs are revoked on reset/unmount.

The help dialog explains accepted input patterns. It uses shadcn `Dialog` and `Tabs`.

## SOX Processing Flow

Public entry points:

- `processSoxFile(file, manualOverrides?)`
- `processSoxFiles(files, manualOverrides?)`
- `classifyError(err)`

High-level algorithm in `processSoxFiles`:

1. Validate at least one file.
2. Load each workbook with `ExcelJS.Workbook().xlsx.load`.
3. Analyze each source workbook to detect language.
4. Copy all visible non-generated worksheets into a combined workbook.
5. If multiple files have duplicate sheet names, prefix with the file stem.
6. Analyze the combined workbook.
7. Remove existing generated `status_N` sheets if any.
8. Extract all dynamic status values from detected status columns.
9. Create one summary sheet.
10. Write the workbook to a Blob and return a generated filename.

Generated filenames:

- Single file: `<original_stem>_with_status.xlsx`
- Multiple files: `sox_summary_of_controls_combined_with_status.xlsx`

## Sheet Detection Rules

Visible sheets only are processed. Hidden sheets are ignored.

Generated/input summary sheets are skipped while copying when their names match:

- `status_N`
- `סיכום`
- `summary`

Header row detection scans the first 30 rows and scores rows by:

- number of non-empty cells,
- header-like text cells,
- status header matches,
- key column matches,
- id/control number matches,
- non-empty count in the following row.

Table end detection starts below the header and stops at the first row where columns A, B, and C are all empty. Rows with content only from D onward are not considered data rows.

Status columns are required. Recognized status/conclusion headers include:

- Hebrew: `מסקנה`, `בקרה`, `סטטוס`, `מסקנה סבב א`, `מסקנה סבב א'`, and any header beginning with `מסקנה`.
- English: `status`, `conclusion`, `control`, `results`, and any header beginning with `conclusion`.

Key column is optional. Recognized key headers are prioritized:

1. `מפתח`, `key control`
2. `בקרת מפתח`, `key`
3. Other headers containing `מפתח` but not `לא מפתח`

If multiple candidates exist at the same priority, processing stops with a Hebrew clarification error.

ID column is optional. It prefers exact matches like `מס`, `מס'`, `מספר`, `מס בקרה`, `control number`, `control no`, `id`, `no`, `no.`. It avoids status/key columns and known non-ID headers.

Process column is optional. Recognized as `process`, `תהליך`, then secondary names like `process name`, `שם התהליך`. Headers containing `sub`, `sub-`, `sub_`, `משנה`, or `תת` are rejected so sub-process columns are not used.

Round numbering currently follows status column order only: first status column is round 1, second is round 2, etc. Parsed round numbers from header text are not used in the final config.

## Summary Sheet Behavior

Language detection checks status headers. If uploaded source files have mixed languages, summary language defaults to English. Otherwise it uses the detected language, defaulting to Hebrew.

The summary sheet is named:

- Hebrew: `סיכום`
- English: `Summary`

Columns are dynamic:

- Always starts with sheet name.
- Adds process column if any processed sheet has a process column.
- Adds round column only if any sheet has more than one status column.
- Adds total controls, key controls, one column per dynamic status value, total statuses, IT controls, annual, execution percent.

Rows are created per `(sheet x process x status round)` when process values exist, or per `(sheet x status round)` otherwise. Rounds with no status values for that process/sheet are skipped.

The summary uses Excel formulas (`SUMPRODUCT`, `COUNTIF`, etc.) so Excel recalculates when opened. It does not hard-code all calculated values.

Excluded status values are not added as dynamic status columns:

- empty string
- `לא מפתח`
- `not key`
- `IT`
- `it`
- `annual`
- `שנתי`

Execution percent depends on a dynamic status value exactly equal to `עבר` or `pass` after lowercase/trim. If no such status exists, execution percent is `0`.

## Error Handling

Business errors extend `SoxBusinessStop` and are displayed directly in Hebrew:

- `SoxClarificationRequired`
- `SoxStructureNotIdentified`

`classifyError` maps generic Excel/worksheet/zip errors to friendlier Hebrew messages and includes technical details for unexpected errors.

Server-side rendering has extra protection:

- `src/start.ts` catches middleware errors.
- `src/server.ts` wraps TanStack Start server entry and replaces swallowed h3 JSON 500 responses with a branded HTML error page.
- `src/lib/error-capture.ts` stores the last global error/rejection for up to 5 seconds so `server.ts` can log the real error.

## Styling Notes

The app is RTL and Hebrew-first:

- Root HTML is `lang="he"` and `dir="rtl"`.
- Heebo is loaded locally from `public/fonts` and also referenced via Google Fonts.
- `src/styles.css` defines custom Tailwind v4 theme tokens for KPMG-like blue/teal/red/gray colors.

Most UI components in `src/components/ui` are standard shadcn/Radix-style scaffolding. Prefer using the existing components and `lucide-react` icons.

## Lovable/TanStack Gotchas

`vite.config.ts` intentionally uses:

```ts
import { defineConfig } from "@lovable.dev/vite-tanstack-config";
```

Do not manually add plugins already provided by this config package. The file notes that Lovable already includes TanStack Start, React, Tailwind, tsconfig paths, Cloudflare build plugin, component tagger, `VITE_*` env injection, aliases, dedupe, error logger plugins, and sandbox detection.

If changing the server entry, keep `tanstackStart.server.entry = "server"` unless there is a specific reason to remove the custom SSR wrapper.

## Current Gaps / Watch Points

- There are no local tests in the repo. Any processor change should ideally add fixture-based tests or at least be verified with representative `.xlsx` files.
- Summary formulas may not be recalculated by ExcelJS itself; behavior depends on Excel/Sheets opening and recalculating.
- Sheet copying preserves values, styles, widths, row heights, views, and merges, but not every possible Excel workbook feature.
- `manualOverrides` exists for key-column overrides but no UI currently exposes it.
- The app is client-side heavy for Excel processing; very large workbooks may affect browser memory/performance.
