/**
 * The live Excel implementation of Host.
 *
 * Structural types again, deliberately: CI has no Office host, and the ambient
 * @microsoft/office-js globals do not type-check without one. The real
 * Excel.OfficeExtension objects satisfy these shapes at runtime.
 *
 * Two rules this file exists to enforce:
 *   1. Every read is intersected with the used range. A user clicking a column
 *      header selects 1,048,576 rows, and marshalling that across the Office.js
 *      bridge hangs the task pane.
 *   2. Writes go through `formulas` alone, never `values` and `formulas` in the
 *      same sync. Assigning both applies the second over the whole rectangle,
 *      blanking every cell the first one filled.
 */

import type {
  CellValue,
  ConditionalSpec,
  FormulaReadResult,
  Host,
  PendingFormat,
  PendingWrite,
  RawSheetContext,
  ReadResult,
} from "./executor.js";

interface Loadable {
  load(props: string): void;
}

interface RangeFormatLike {
  autofitColumns(): void;
  fill: { color: string };
  font: { color: string; bold: boolean; italic: boolean };
  numberFormat: string[][];
}

interface ConditionalFormatLike {
  id: string;
  delete(): void;
}

interface RangeLike extends Loadable {
  address: string;
  rowIndex: number;
  columnIndex: number;
  rowCount: number;
  columnCount: number;
  values: CellValue[][];
  formulas: Array<Array<string | number | null>>;
  isNullObject?: boolean;
  getIntersectionOrNullObject(other: RangeLike): RangeLike;
  format: RangeFormatLike;
  conditionalFormats: {
    add(...args: unknown[]): ConditionalFormatLike;
    getItem(id: string): ConditionalFormatLike;
  };
}

interface SheetLike extends Loadable {
  name: string;
  getUsedRangeOrNullObject(): RangeLike;
  getRange(address: string): RangeLike;
  getRangeByIndexes(row: number, col: number, rows: number, cols: number): RangeLike;
  tables: { load(p: string): void; items: Array<{ name: string; getRange(): RangeLike }> };
}

interface ContextLike {
  workbook: {
    worksheets: {
      getActiveWorksheet(): SheetLike;
      getItem(name: string): SheetLike;
      load(props: string): void;
      items: SheetLike[];
    };
    getSelectedRange(): RangeLike;
    tables: { add(range: RangeLike, hasHeaders: boolean): { name: string; getRange(): RangeLike } };
  };
  sync(): Promise<unknown>;
}

interface ExcelLike {
  run<T>(cb: (ctx: ContextLike) => Promise<T>): Promise<T>;
}

/** The Excel namespace, or null when running as a plain web page. */
export function getExcel(): ExcelLike | null {
  const g = globalThis as { Excel?: ExcelLike };
  return g.Excel ?? null;
}

export function isInExcel(): boolean {
  return getExcel() !== null;
}

function excelOrThrow(): ExcelLike {
  const e = getExcel();
  if (!e) {
    throw new Error("Not running inside Excel, so the worksheet is unavailable.");
  }
  return e;
}

/**
 * Split an optional sheet-qualified address ("Sheet2!A1:C10") into its sheet
 * name and plain range. A bare range resolves to the active sheet (null).
 */
export function splitQualifiedAddress(address: string): { sheetName: string | null; range: string } {
  const pos = address.indexOf("!");
  if (pos < 0) return { sheetName: null, range: address };
  return { sheetName: address.substring(0, pos), range: address.substring(pos + 1) };
}

/** Bounded read: clip to the used range, then cap rows. */
function clip(
  sheet: SheetLike,
  target: RangeLike,
  used: RangeLike,
  maxRows: number,
): { range: RangeLike; totalRows: number; truncated: boolean } | null {
  if (used.isNullObject) return null;
  const top = Math.max(target.rowIndex, used.rowIndex);
  const left = Math.max(target.columnIndex, used.columnIndex);
  const bottom = Math.min(target.rowIndex + target.rowCount, used.rowIndex + used.rowCount);
  const right = Math.min(target.columnIndex + target.columnCount, used.columnIndex + used.columnCount);
  const totalRows = bottom - top;
  const cols = right - left;
  if (totalRows <= 0 || cols <= 0) return null;
  const rows = Math.min(totalRows, maxRows);
  return {
    range: sheet.getRangeByIndexes(top, left, rows, cols),
    totalRows,
    truncated: rows < totalRows,
  };
}

const USED_PROPS = "address,rowIndex,columnIndex,rowCount,columnCount,isNullObject";

/**
 * Resolve an optional sheet-qualified address to its worksheet. A bare range
 * ("A1:C10") targets the active sheet; "Sheet2!A1:C10" targets Sheet2.
 * Returns the resolved sheet (name loaded) and the plain range address.
 */
async function resolveSheet(
  ctx: ContextLike,
  address: string,
): Promise<{ sheet: SheetLike; rangeAddress: string }> {
  const { sheetName, range } = splitQualifiedAddress(address);
  if (sheetName === null) {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");
    await ctx.sync();
    return { sheet, rangeAddress: range };
  }
  const worksheets = ctx.workbook.worksheets;
  worksheets.load("items/name");
  await ctx.sync();
  const sheet = worksheets.getItem(sheetName);
  sheet.load("name");
  await ctx.sync();
  return { sheet, rangeAddress: range };
}

/** Prefix a range address with the resolved sheet name, for unambiguous replies.
 *  Office.js range.address already carries the sheet ("Sheet1!A1:C3"), so strip
 *  any existing prefix before re-qualifying with the sheet we actually resolved. */
function qualify(sheet: SheetLike, address: string): string {
  const pos = address.indexOf("!");
  const bare = pos >= 0 ? address.substring(pos + 1) : address;
  return `${sheet.name}!${bare}`;
}

export const excelHost: Host = {
  async getSheetContext(sampleRows: number): Promise<RawSheetContext> {
    return excelOrThrow().run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      const used = sheet.getUsedRangeOrNullObject();
      used.load(USED_PROPS);
      const selection = ctx.workbook.getSelectedRange();
      selection.load("address");
      sheet.tables.load("items/name");
      await ctx.sync();

      if (used.isNullObject) {
        return {
          sheetName: sheet.name,
          usedRangeAddress: null,
          rowCount: 0,
          columnCount: 0,
          selectionAddress: selection.address ?? null,
          sample: [],
          tables: [],
        };
      }

      const rows = Math.min(used.rowCount, Math.max(sampleRows, 1));
      const sample = sheet.getRangeByIndexes(used.rowIndex, used.columnIndex, rows, used.columnCount);
      sample.load("values,formulas");
      const tableRanges = sheet.tables.items.map((t) => {
        const r = t.getRange();
        r.load("address");
        return { name: t.name, range: r };
      });
      await ctx.sync();

      return {
        sheetName: sheet.name,
        usedRangeAddress: used.address,
        rowCount: used.rowCount,
        columnCount: used.columnCount,
        selectionAddress: selection.address ?? null,
        sample: sample.values ?? [],
        sampleFormulas: (sample.formulas ?? []).map((row) =>
          row.map((c) => (c === null || c === undefined ? "" : String(c))),
        ),
        tables: tableRanges.map((t) => ({ name: t.name, address: t.range.address })),
      };
    });
  },

  async readRange(address: string, maxRows: number): Promise<ReadResult> {
    return excelOrThrow().run(async (ctx) => {
      const { sheet, rangeAddress } = await resolveSheet(ctx, address);
      const target = sheet.getRange(rangeAddress);
      target.load("address,rowIndex,columnIndex,rowCount,columnCount");
      const used = sheet.getUsedRangeOrNullObject();
      used.load(USED_PROPS);
      await ctx.sync();

      const clipped = clip(sheet, target, used, maxRows);
      if (!clipped) {
        return { address: qualify(sheet, target.address), values: [], truncated: false, totalRows: 0 };
      }
      clipped.range.load("address,values");
      await ctx.sync();
      return {
        address: qualify(sheet, clipped.range.address),
        values: clipped.range.values ?? [],
        truncated: clipped.truncated,
        totalRows: clipped.totalRows,
      };
    });
  },
  async readFormulas(address: string, maxRows: number): Promise<FormulaReadResult> {
    return excelOrThrow().run(async (ctx) => {
      const { sheet, rangeAddress } = await resolveSheet(ctx, address);
      const target = sheet.getRange(rangeAddress);
      target.load("address,rowIndex,columnIndex,rowCount,columnCount");
      const used = sheet.getUsedRangeOrNullObject();
      used.load(USED_PROPS);
      await ctx.sync();

      const clipped = clip(sheet, target, used, maxRows);
      if (!clipped) {
        return { address: qualify(sheet, target.address), formulas: [], truncated: false, totalRows: 0 };
      }
      clipped.range.load("address,formulas");
      await ctx.sync();
      return {
        address: qualify(sheet, clipped.range.address),
        formulas: (clipped.range.formulas ?? []).map((row) =>
          row.map((c) => (c === null || c === undefined ? "" : String(c))),
        ),
        truncated: clipped.truncated,
        totalRows: clipped.totalRows,
      };
    });
  },

  async createTable(address: string, hasHeaders: boolean): Promise<{ address: string }> {
    return excelOrThrow().run(async (ctx) => {
      const { sheet, rangeAddress } = await resolveSheet(ctx, address);
      const range = sheet.getRange(rangeAddress);
      const table = ctx.workbook.tables.add(range, hasHeaders);
      const created = table.getRange();
      created.load("address");
      await ctx.sync();
      return { address: created.address };
    });
  },
  async getSheetNames(): Promise<string[]> {
    return excelOrThrow().run(async (ctx) => {
      const sheets = ctx.workbook.worksheets;
      sheets.load("items/name");
      await ctx.sync();
      return sheets.items.map((s) => s.name);
    });
  }
};

/** What the user can put back if they change their mind. */
export interface AppliedWrite {
  address: string;
  /** Formulas as they were before we overwrote them. */
  priorFormulas: Array<Array<string | number | null>>;
}

/**
 * Apply a staged write, after the user accepted it.
 *
 * Values and formulas go through a single `formulas` assignment: Office treats
 * a plain literal in that array as a literal, and a leading "=" as a live
 * formula, so one array covers both without the two-assignment clobber.
 */
export async function applyWrite(write: PendingWrite): Promise<AppliedWrite> {
  return excelOrThrow().run(async (ctx) => {
    const { sheet, rangeAddress } = await resolveSheet(ctx, write.address);
    const height = write.values.length;
    const width = write.values[0]?.length ?? 0;
    const anchor = sheet.getRange(rangeAddress);
    anchor.load("rowIndex,columnIndex");
    await ctx.sync();

    const dest = sheet.getRangeByIndexes(anchor.rowIndex, anchor.columnIndex, height, width);
    dest.load("address,formulas");
    await ctx.sync();
    const priorFormulas = dest.formulas ?? [];

    dest.formulas = write.values.map((row) =>
      Array.from({ length: width }, (_, i) => {
        const v = row[i];
        return v === null || v === undefined ? "" : typeof v === "boolean" ? String(v) : v;
      }),
    );
    dest.format.autofitColumns();
    await ctx.sync();

    return { address: dest.address, priorFormulas };
  });
}

/** Restore the formulas captured before a write. */
export async function revertWrite(applied: AppliedWrite): Promise<void> {
  await excelOrThrow().run(async (ctx) => {
    const { sheet, rangeAddress } = await resolveSheet(ctx, applied.address);
    sheet.getRange(rangeAddress).formulas = applied.priorFormulas;
    await ctx.sync();
  });
}

/** Everything needed to put a range's formatting back. */
export interface AppliedFormat {
  address: string;
  prior: { fill: string | null; fontColor: string | null; bold: boolean | null; italic: boolean | null; numberFormat: string[][] | null };
  /** Ids of conditional formats this apply added, so undo can delete exactly those. */
  addedConditionalIds: string[];
}

/** Map a ConditionalSpec to Office.js conditionalFormats.add arguments. */
function conditionalAddArgs(when: ConditionalSpec["when"]): unknown[] {
  switch (when) {
    case "positive":
      return ["CellIs", null, "GreaterThan", "0"];
    case "negative":
      return ["CellIs", null, "LessThan", "0"];
    case "blank":
      return ["CellIs", null, "EqualTo", "="];
    case "top10":
      return ["TopBottom", null, null, null]; // rank defaults to top 10
  }
}

/**
 * Apply a staged formatting proposal, after the user accepted it. Only the
 * properties present on the spec are touched; the rest keep whatever the user
 * had. Conditional rules are added as real Excel conditional formats, so they
 * keep tracking the data.
 */
export async function applyFormat(pending: PendingFormat): Promise<AppliedFormat> {
  return excelOrThrow().run(async (ctx) => {
    const { sheet, rangeAddress } = await resolveSheet(ctx, pending.address);
    const range = sheet.getRange(rangeAddress);
    range.load("address,format/fill/color,format/font/color,format/font/bold,format/font/italic,format/numberFormat");
    await ctx.sync();

    const format = range.format;
    const prior = {
      fill: format.fill.color ?? null,
      fontColor: format.font.color ?? null,
      bold: format.font.bold ?? null,
      italic: format.font.italic ?? null,
      numberFormat: format.numberFormat ?? null,
    };

    const { format: spec, conditional } = pending;
    if (spec.fill !== undefined) format.fill.color = spec.fill;
    if (spec.fontColor !== undefined) format.font.color = spec.fontColor;
    if (spec.bold !== undefined) format.font.bold = spec.bold;
    if (spec.italic !== undefined) format.font.italic = spec.italic;
    if (spec.numberFormat !== undefined) {
      range.load("rowCount,columnCount");
      await ctx.sync();
      const nf: string[][] = [];
      for (let r = 0; r < range.rowCount; r++) nf.push(Array(range.columnCount).fill(spec.numberFormat));
      format.numberFormat = nf;
    }

    const addedConditionalIds: string[] = [];
    if (conditional) {
      const style = conditional.style;
      const args = conditionalAddArgs(conditional.when);
      const added = range.conditionalFormats.add(...args) as ConditionalFormatLike & {
        format: { fill: { color: string }; font: { color: string; bold: boolean } };
      };
      added.format.fill.color = style.fill ?? "";
      added.format.font.color = style.fontColor ?? "";
      if (style.bold !== undefined) added.format.font.bold = style.bold;
      await ctx.sync();
      addedConditionalIds.push(added.id);
    }

    await ctx.sync();
    return { address: range.address, prior, addedConditionalIds };
  });
}

/** Undo an accepted format, restoring the captured prior state. */
export async function revertFormat(applied: AppliedFormat): Promise<void> {
  await excelOrThrow().run(async (ctx) => {
    const { sheet, rangeAddress } = await resolveSheet(ctx, applied.address);
    const range = sheet.getRange(rangeAddress);
    const { prior, addedConditionalIds } = applied;
    for (const id of addedConditionalIds) {
      range.conditionalFormats.getItem(id).delete();
    }
    if (prior.fill !== null) range.format.fill.color = prior.fill;
    if (prior.fontColor !== null) range.format.font.color = prior.fontColor;
    if (prior.bold !== null) range.format.font.bold = prior.bold;
    if (prior.italic !== null) range.format.font.italic = prior.italic;
    if (prior.numberFormat !== null) range.format.numberFormat = prior.numberFormat;
    await ctx.sync();
  });
}
