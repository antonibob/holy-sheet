/**
 * Tool execution.
 *
 * The model chooses a tool; this runs it. Everything host-specific sits behind
 * the Host interface so the dispatch logic below is unit-testable without an
 * Excel instance — CI has no Office host.
 *
 * One rule is enforced here rather than trusted to the prompt: nothing in this
 * file writes to the worksheet. propose_write stages a diff for the user to
 * accept. If the model hallucinates that it saved something, the sheet is still
 * untouched.
 */

import { findDuplicates, parseAmount } from "../engine/index.js";
import { reconcileTwoColumns } from "../taskpane/reconcile.js";
import type { ToolCall, ToolResult } from "../model/tools.js";

export type CellValue = string | number | boolean | null;

export interface ColumnInfo {
  /** Spreadsheet letter for this column, e.g. "C". */
  letter: string;
  /** Zero-based offset within the range that was inspected. */
  index: number;
  header: string;
  inferredType: "amount" | "date" | "text" | "number" | "empty";
  sample: string[];
}

export interface RawSheetContext {
  sheetName: string;
  /** Null when the sheet is empty. */
  usedRangeAddress: string | null;
  rowCount: number;
  columnCount: number;
  selectionAddress: string | null;
  /** Header + a bounded sample of data rows from the used range. */
  sample: CellValue[][];
  /** Formula view of the same rows, when the host can provide it. */
  sampleFormulas?: string[][];
  tables: Array<{ name: string; address: string }>;
}

export interface ReadResult {
  address: string;
  values: CellValue[][];
  truncated: boolean;
  totalRows: number;
}

export interface FormulaReadResult {
  address: string;
  /**
   * The formula view of each cell: "=SUM(A1:A9)" for live formulas, the plain
   * literal for constant cells. This is what distinguishes a computed cell
   * from a pasted number.
   */
  formulas: string[][];
  truncated: boolean;
  totalRows: number;
}

/** Everything the executor needs from the spreadsheet host. */
export interface Host {
  getSheetContext(sampleRows: number): Promise<RawSheetContext>;
  getSheetNames(): Promise<string[]>;
  readRange(address: string, maxRows: number): Promise<ReadResult>;
  readFormulas(address: string, maxRows: number): Promise<FormulaReadResult>;
  createTable(address: string, hasHeaders: boolean): Promise<{ address: string }>;
}

/** A change awaiting the user's yes/no. Never applied by the executor. */
export interface PendingWrite {
  id: string;
  address: string;
  note: string;
  values: CellValue[][];
  /** Current cell contents, so the UI can show a before/after diff. */
  before: CellValue[][];
}

export interface ExecutorDeps {
  host: Host;
  /** Shows the user a diff; returns what the model is told. */
  stageWrite: (write: PendingWrite) => void;
}

const DEFAULT_SAMPLE_ROWS = 8;
const DEFAULT_MAX_ROWS = 200;
/** Hard ceiling regardless of what the model asks for; keeps the bridge sane. */
const ABSOLUTE_MAX_ROWS = 5000;

/** Zero-based column index to spreadsheet letter: 0 -> A, 26 -> AA. */
export function columnLetter(index: number): string {
  let n = index;
  let out = "";
  do {
    out = String.fromCharCode(65 + (n % 26)) + out;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return out;
}

/** Leading column index of an A1 address, so range offsets map to real letters. */
export function startColumnOf(address: string): number {
  const cell = address.includes("!") ? address.slice(address.indexOf("!") + 1) : address;
  const m = /^\$?([A-Z]+)/i.exec(cell.trim());
  if (!m?.[1]) return 0;
  return m[1]
    .toUpperCase()
    .split("")
    .reduce((acc, ch) => acc * 26 + (ch.charCodeAt(0) - 64), 0) - 1;
}

/** Leading row index (zero-based) of an A1 address. */
export function startRowOf(address: string): number {
  const cell = address.includes("!") ? address.slice(address.indexOf("!") + 1) : address;
  const m = /^\$?[A-Z]+\$?(\d+)/i.exec(cell.trim());
  return m?.[1] ? Number(m[1]) - 1 : 0;
}

function looksLikeDate(s: string): boolean {
  if (!/\d/.test(s)) return false;
  return /^\d{4}-\d{2}-\d{2}/.test(s) || /^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}$/.test(s);
}

function looksLikeAmount(s: string): boolean {
  // Accounting shapes: 1,234.56  (123.45)  $45.00  45.00 CR
  return /^[($\s]*-?[\d,]+(\.\d{1,2})?[)\s]*(DR|CR)?$/i.test(s.trim()) && /\d/.test(s);
}

function inferType(samples: string[]): ColumnInfo["inferredType"] {
  const nonEmpty = samples.filter((s) => s.trim() !== "");
  if (nonEmpty.length === 0) return "empty";
  const ratio = (pred: (s: string) => boolean) =>
    nonEmpty.filter(pred).length / nonEmpty.length;
  if (ratio(looksLikeDate) > 0.6) return "date";
  if (ratio(looksLikeAmount) > 0.6) return "amount";
  if (ratio((s) => !Number.isNaN(Number(s))) > 0.6) return "number";
  return "text";
}

/**
 * Describe the columns of a sampled range so the model can pick indices without
 * guessing. This is what makes hardcoding "columns 0 and 1" unnecessary.
 */
export function describeColumns(sample: CellValue[][], baseAddress: string | null): ColumnInfo[] {
  if (sample.length === 0) return [];
  const startCol = baseAddress ? startColumnOf(baseAddress) : 0;
  const width = Math.max(...sample.map((r) => r.length));
  const headerRow = sample[0] ?? [];
  const body = sample.slice(1);
  return Array.from({ length: width }, (_, i) => {
    const samples = body.map((r) => String(r[i] ?? ""));
    return {
      letter: columnLetter(startCol + i),
      index: i,
      header: String(headerRow[i] ?? "").trim() || `(column ${columnLetter(startCol + i)})`,
      inferredType: inferType(samples),
      sample: samples.slice(0, 3),
    };
  });
}

/**
 * Sort a range's rows by one column, keeping the first row (header) fixed and
 * every row intact. Pure: returns new arrays, never mutates the input.
 * Numeric-looking cells compare as numbers, everything else as text; blank
 * cells sink to the bottom regardless of direction.
 */
export function sortRangeValues(
  values: CellValue[][],
  column: number,
  direction: "asc" | "desc",
): CellValue[][] {
  const header = values[0] ?? [];
  const body = values.slice(1).map((row) => [...row]);
  const isEmpty = (v: CellValue) => v === null || v === undefined || String(v).trim() === "";
  const numeric = (v: CellValue): number | null => {
    if (typeof v === "number") return v;
    const s = String(v).trim();
    if (s === "" || Number.isNaN(Number(s))) return null;
    return Number(s);
  };
  const dir = direction === "desc" ? -1 : 1;
  body.sort((ra, rb) => {
    const a = ra[column] ?? null;
    const b = rb[column] ?? null;
    const ae = isEmpty(a);
    const be = isEmpty(b);
    if (ae && be) return 0;
    if (ae) return 1; // blanks last, whichever way we sort
    if (be) return -1;
    const an = numeric(a);
    const bn = numeric(b);
    if (an !== null && bn !== null) return (an - bn) * dir;
    return String(a).localeCompare(String(b)) * dir;
  });
  return [header, ...body];
}

function ok(call: ToolCall, content: unknown): ToolResult {
  return { toolCallId: call.id, name: call.name, content };
}

function fail(call: ToolCall, message: string): ToolResult {
  return { toolCallId: call.id, name: call.name, content: { error: message }, isError: true };
}

/** Read a required argument, reporting clearly instead of coercing nonsense. */
function str(args: Record<string, unknown>, key: string): string | null {
  const v = args[key];
  return typeof v === "string" && v.trim() !== "" ? v.trim() : null;
}

function int(args: Record<string, unknown>, key: string): number | null {
  const v = args[key];
  if (typeof v === "number" && Number.isInteger(v)) return v;
  // Models routinely send "2" where the schema says integer.
  if (typeof v === "string" && /^-?\d+$/.test(v.trim())) return Number(v.trim());
  return null;
}

/** One group in an aggregate_range result. Money is exact integer cents. */
export interface AggregateGroup {
  key: string;
  count: number;
  sumCents: number;
  minCents: number | null;
  maxCents: number | null;
  /** Row offsets (within the read range) of members, for click-through. */
  rows: number[];
  /** Members whose amount actually parsed — the denominator for avg. */
  parsedCount: number;
}

export interface AggregateOptions {
  aggregations?: string[];
  maxGroups?: number;
}

export interface AggregateResult {
  groups: AggregateGroup[];
  totalGroups: number;
  grandSumCents: number;
  rowCount: number;
  /** Non-empty cells in valueColumn that could not be read as amounts. */
  parseFailures: Array<{ row: number; raw: string }>;
}

/**
 * Host-side group-by/summary over one text/group column and one amount column.
 *
 * Exists so the model never has to pull thousands of rows into context to add
 * them up. All money arithmetic is exact integer cents via the shared amount
 * parser (parentheses negatives, DR/CR, currency symbols), and rows whose
 * amount cannot be parsed are reported back rather than dropped silently.
 */
export function aggregateRange(
  values: CellValue[][],
  groupBy: number,
  valueColumn: number,
  opts: AggregateOptions = {},
): AggregateResult {
  const body = values.slice(1); // first row is the header, matching read_range/reconcile conventions
  const groups = new Map<string, AggregateGroup>();
  const parseFailures: Array<{ row: number; raw: string }> = [];
  let grandSumCents = 0;

  body.forEach((row, offset) => {
    const raw = row[valueColumn];
    const key = String(row[groupBy] ?? "").trim() || "(blank)";
    let g = groups.get(key);
    if (!g) {
      g = { key, count: 0, sumCents: 0, minCents: null, maxCents: null, rows: [], parsedCount: 0 };
      groups.set(key, g);
    }
    g.count += 1;
    g.rows.push(offset);

    if (raw === null || raw === undefined || String(raw).trim() === "") return; // absence, not an error
    const parsed = parseAmount(raw as string | number);
    if (!parsed.ok) {
      parseFailures.push({ row: offset + 2, raw: String(raw) }); // +2: skip header, 1-based
      return;
    }
    g.minCents = g.minCents === null ? parsed.cents : Math.min(g.minCents, parsed.cents);
    g.maxCents = g.maxCents === null ? parsed.cents : Math.max(g.maxCents, parsed.cents);
    g.sumCents += parsed.cents;
    g.parsedCount += 1;
    grandSumCents += parsed.cents;
  });

  const maxGroups = opts.maxGroups ?? 200;
  const sorted = [...groups.values()].sort((a, b) => b.sumCents - a.sumCents || a.key.localeCompare(b.key));
  return {
    groups: sorted.slice(0, maxGroups),
    totalGroups: groups.size,
    grandSumCents,
    rowCount: body.length,
    parseFailures,
  };
}

let writeSeq = 0;

export function createExecutor(deps: ExecutorDeps) {
  const { host, stageWrite } = deps;

  return async function execute(call: ToolCall): Promise<ToolResult> {
    switch (call.name) {
      case "get_sheet_context": {
        const ctx = await host.getSheetContext(DEFAULT_SAMPLE_ROWS);
        return ok(call, {
          sheetName: ctx.sheetName,
          usedRange: ctx.usedRangeAddress,
          rows: ctx.rowCount,
          columns: ctx.columnCount,
          selection: ctx.selectionAddress,
          tables: ctx.tables,
          columnDetail: describeColumns(ctx.sample, ctx.usedRangeAddress),
          sampleRows: ctx.sample,
          sampleFormulas: ctx.sampleFormulas,
          note:
            ctx.rowCount === 0
              ? "The sheet is empty."
              : "Column indices in columnDetail are offsets within usedRange, which is what the reconcile and duplicate tools expect.",
        });
      }
      case "list_sheets": {
        const names = await host.getSheetNames();
        return ok(call, { names });
      }

      case "read_range": {
        const address = str(call.args, "address");
        if (!address) return fail(call, "read_range needs an `address`, e.g. \"A1:D200\".");
        const requested = int(call.args, "maxRows") ?? DEFAULT_MAX_ROWS;
        const maxRows = Math.min(Math.max(requested, 1), ABSOLUTE_MAX_ROWS);
        const res = await host.readRange(address, maxRows);
        return ok(call, {
          address: res.address,
          values: res.values,
          totalRows: res.totalRows,
          truncated: res.truncated,
          note: res.truncated
            ? `Showing the first ${res.values.length} of ${res.totalRows} rows. Totals must come from a tool, not from counting these.`
            : undefined,
        });
      }

      case "read_formulas": {
        const address = str(call.args, "address");
        if (!address) return fail(call, "read_formulas needs an `address`, e.g. \"A1:D200\".");
        const requested = int(call.args, "maxRows") ?? DEFAULT_MAX_ROWS;
        const maxRows = Math.min(Math.max(requested, 1), ABSOLUTE_MAX_ROWS);
        const res = await host.readFormulas(address, maxRows);
        return ok(call, {
          address: res.address,
          formulas: res.formulas,
          totalRows: res.totalRows,
          truncated: res.truncated,
          note: res.truncated
            ? `Showing the formula view of the first ${res.formulas.length} of ${res.totalRows} rows.`
            : "Cells shown as literals are constants; cells shown with a leading = are live formulas the sheet recalculates.",
        });
      }

      case "reconcile_columns": {
        const address = str(call.args, "address");
        const colA = int(call.args, "columnA");
        const colB = int(call.args, "columnB");
        if (!address) return fail(call, "reconcile_columns needs an `address`.");
        if (colA === null || colB === null) {
          return fail(call, "reconcile_columns needs integer `columnA` and `columnB` offsets within the address.");
        }
        if (colA === colB) return fail(call, "columnA and columnB are the same column.");
        const res = await host.readRange(address, ABSOLUTE_MAX_ROWS);
        if (res.truncated) {
          return fail(
            call,
            `That range has ${res.totalRows} rows, beyond the ${ABSOLUTE_MAX_ROWS}-row limit for a single reconciliation. Narrow the address.`,
          );
        }
        const width = Math.max(0, ...res.values.map((r) => r.length));
        if (colA >= width || colB >= width) {
          return fail(call, `The range is ${width} column(s) wide; columnA=${colA} columnB=${colB} is out of bounds.`);
        }
        const summary = reconcileTwoColumns(
          res.values as Array<Array<string | number | null | undefined>>,
          colA,
          colB,
          { toleranceCents: int(call.args, "toleranceCents") ?? 0 },
        );
        // Offsets become worksheet row numbers the reviewer can actually click.
        const base = startRowOf(res.address);
        const toSheetRow = (r: number) => base + r + 1;
        return ok(call, {
          matchedPairs: summary.matchedPairs,
          glTotal: (summary.sumACents / 100).toFixed(2),
          bankTotal: (summary.sumBCents / 100).toFixed(2),
          residual: (summary.residualCents / 100).toFixed(2),
          balanced: summary.residualCents === 0,
          likelyTransposition: summary.likelyTransposition,
          unmatchedGlRows: summary.unmatchedA.map(toSheetRow),
          unmatchedBankRows: summary.unmatchedB.map(toSheetRow),
          sampleMatches: summary.matches.slice(0, 10).map((m) => ({
            glRow: toSheetRow(m.aRow),
            bankRow: toSheetRow(m.bRow),
            amount: (m.amountCents / 100).toFixed(2),
          })),
          unreadableCells: summary.parseFailures.map((f) => ({
            row: toSheetRow(f.row),
            side: f.side === "A" ? "GL" : "bank",
            value: f.raw,
          })),
          summary: summary.text,
        });
      }

      case "aggregate_range": {
        const address = str(call.args, "address");
        const groupBy = int(call.args, "groupBy");
        const valueColumn = int(call.args, "valueColumn");
        if (!address) return fail(call, "aggregate_range needs an `address`.");
        if (groupBy === null || valueColumn === null) {
          return fail(call, "aggregate_range needs integer `groupBy` and `valueColumn` offsets within the address.");
        }
        if (groupBy === valueColumn) return fail(call, "groupBy and valueColumn are the same column.");
        const res = await host.readRange(address, ABSOLUTE_MAX_ROWS);
        if (res.truncated) {
          return fail(
            call,
            `That range has ${res.totalRows} rows, beyond the ${ABSOLUTE_MAX_ROWS}-row limit for one aggregation. Narrow the address.`,
          );
        }
        const width = Math.max(0, ...res.values.map((r) => r.length));
        if (groupBy >= width || valueColumn >= width) {
          return fail(call, `The range is ${width} column(s) wide; groupBy=${groupBy} valueColumn=${valueColumn} is out of bounds.`);
        }
        const agg = aggregateRange(res.values, groupBy, valueColumn);
        const money = (c: number | null) => (c === null ? null : (c / 100).toFixed(2));
        return ok(call, {
          address: res.address,
          rowsGrouped: agg.rowCount,
          totalGroups: agg.totalGroups,
          grandSum: (agg.grandSumCents / 100).toFixed(2),
          groups: agg.groups.map((g) => ({
            key: g.key,
            count: g.count,
            sum: money(g.sumCents),
            min: money(g.minCents),
            max: money(g.maxCents),
            avg: g.parsedCount > 0 ? money(Math.round(g.sumCents / g.parsedCount)) : null,
          })),
          truncated: agg.totalGroups > agg.groups.length,
          unreadableCells: agg.parseFailures.slice(0, 20),
          unreadableCount: agg.parseFailures.length,
          summary:
            `${agg.totalGroups} group(s) over ${agg.rowCount} row(s); grand total ${(agg.grandSumCents / 100).toFixed(2)}.` +
            (agg.parseFailures.length > 0
              ? ` ${agg.parseFailures.length} amount cell(s) could not be parsed and are excluded from the sums (rows ${agg.parseFailures
                  .slice(0, 5)
                  .map((f) => f.row)
                  .join(", ")}${agg.parseFailures.length > 5 ? ", …" : ""}).`
              : ""),
        });
      }

      case "find_duplicates": {
        const address = str(call.args, "address");
        const vendorCol = int(call.args, "vendorColumn");
        const amountCol = int(call.args, "amountColumn");
        const dateCol = int(call.args, "dateColumn");
        if (!address) return fail(call, "find_duplicates needs an `address`.");
        if (vendorCol === null || amountCol === null || dateCol === null) {
          return fail(call, "find_duplicates needs integer `vendorColumn`, `amountColumn` and `dateColumn`.");
        }
        const res = await host.readRange(address, ABSOLUTE_MAX_ROWS);
        const base = startRowOf(res.address);
        const rows = res.values.slice(1).map((r) => ({
          vendor: String(r[vendorCol] ?? ""),
          amountCents: Math.round(Number(String(r[amountCol] ?? "0").replace(/[^0-9.-]/g, "")) * 100),
          date: String(r[dateCol] ?? ""),
        }));
        const dups = findDuplicates(rows, { dateWindowDays: int(call.args, "dateWindowDays") ?? 3 });
        return ok(call, {
          count: dups.length,
          // +2: skip the header row, then convert to 1-based sheet numbering.
          duplicates: dups.slice(0, 50).map((d) => ({
            rowA: base + d.a + 2,
            rowB: base + d.b + 2,
            vendor: rows[d.a]?.vendor,
            amount: ((rows[d.a]?.amountCents ?? 0) / 100).toFixed(2),
            reason: d.reason,
          })),
          truncated: dups.length > 50,
        });
      }

      case "propose_write": {
        const address = str(call.args, "address");
        const note = str(call.args, "note") ?? "Proposed change";
        const raw = call.args.values;
        if (!address) return fail(call, "propose_write needs an `address`.");
        if (!Array.isArray(raw) || raw.length === 0) {
          return fail(call, "propose_write needs a non-empty `values` array of rows.");
        }
        if (!raw.every((r) => Array.isArray(r))) {
          return fail(call, "`values` must be an array of arrays, one inner array per row.");
        }
        const values = raw as CellValue[][];
        const width = values[0]?.length ?? 0;
        if (!values.every((r) => r.length === width)) {
          return fail(call, "Every row in `values` must have the same number of cells.");
        }
        // Formula cells must arrive as strings; a number can never be a formula.
        for (const row of values) {
          for (const v of row) {
            if (typeof v === "number" && String(v).startsWith("=")) {
              return fail(call, "A numeric value cannot start with '='; send formulas as strings.");
            }
          }
        }
        // The staged write is applied through the host's `formulas` channel, so
        // any string beginning with "=" lands as a live, recalculating formula
        // and every other value lands as a literal — one assignment, no clobber.
        const formulaCells = values.reduce(
          (n, row) => n + row.filter((v) => typeof v === "string" && v.trim().startsWith("=")).length,
          0,
        );
        let before: CellValue[][] = [];
        try {
          before = (await host.readRange(address, values.length)).values;
        } catch {
          before = []; // a diff without the "before" half is still worth showing
        }
        const write: PendingWrite = {
          id: `w${++writeSeq}`,
          address,
          note,
          values,
          before,
        };
        stageWrite(write);
        return ok(call, {
          staged: true,
          address,
          rows: values.length,
          columns: width,
          formulaCells: formulaCells > 0 ? formulaCells : undefined,
          message:
            "Shown to the user as a diff awaiting their approval. The sheet is unchanged until they accept. Tell the user what you have proposed and why — do not claim it has been applied." +
            (formulaCells > 0
              ? ` ${formulaCells} cell(s) begin with '=' and will be written as live formulas the sheet recalculates.`
              : ""),
        });
      }

      case "sort_range": {
        const address = str(call.args, "address");
        const column = int(call.args, "column");
        const dirRaw = str(call.args, "direction") ?? "asc";
        if (!address) return fail(call, "sort_range needs an `address`, e.g. \"A1:D200\".");
        if (column === null || column < 0) {
          return fail(call, "sort_range needs integer `column` — the zero-based offset within the address.");
        }
        if (dirRaw !== "asc" && dirRaw !== "desc") {
          return fail(call, "`direction` must be \"asc\" or \"desc\".");
        }
        const res = await host.readRange(address, ABSOLUTE_MAX_ROWS);
        if (res.truncated) {
          return fail(call, `Range has ${res.totalRows} rows, past the ${ABSOLUTE_MAX_ROWS}-row ceiling. Narrow the address.`);
        }
        const width = Math.max(0, ...res.values.map((r) => r.length));
        if (column >= width) {
          return fail(call, `Column offset ${column} is outside ${address} (${width} columns). Use describeColumns-style offsets from the range's left edge.`);
        }
        if (res.values.length <= 1) {
          return fail(call, "Nothing to sort: the range has no data rows below its header.");
        }
        const sorted = sortRangeValues(res.values, column, dirRaw);
        const note = str(call.args, "note") ?? `Sort ${address} by column ${column + 1} (${dirRaw})`;
        const write: PendingWrite = {
          id: `w${++writeSeq}`,
          address,
          note,
          values: sorted,
          before: res.values,
        };
        stageWrite(write);
        return ok(call, {
          staged: true,
          address,
          rows: sorted.length - 1,
          columns: width,
          sortBy: column,
          direction: dirRaw,
          message:
            "Shown to the user as a before/after diff awaiting their approval. The sheet is unchanged until they accept. Tell the user what you have proposed and why — do not claim it has been applied.",
        });
      }

      case "create_table": {
        const address = str(call.args, "address");
        if (!address) return fail(call, "create_table needs an `address`.");
        const hasHeaders = call.args.hasHeaders !== false;
        const res = await host.createTable(address, hasHeaders);
        return ok(call, { created: true, address: res.address, hasHeaders });
      }

      default:
        return fail(call, `Unknown tool "${call.name}".`);
    }
  };
}
