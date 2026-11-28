import { describe, it, expect } from "vitest";
import {
  columnLetter,
  createExecutor,
  aggregateRange,
  describeColumns,
  startColumnOf,
  startRowOf,
  type FormulaReadResult,
  type Host,
  type PendingWrite,
  type ReadResult,
} from "../src/office/executor.js";
import type { ToolCall } from "../src/model/tools.js";

function fakeHost(
  values: Array<Array<string | number | null>>,
  address = "A1:C9",
  formulas?: string[][],
): Host {
  const formulaView = formulas ?? values.map((row) => row.map((v) => String(v ?? "")));
  return {
    async getSheetContext() {
      return {
        sheetName: "Aug",
        usedRangeAddress: address,
        rowCount: values.length,
        columnCount: Math.max(0, ...values.map((r) => r.length)),
        selectionAddress: address,
        sample: values.slice(0, 8),
        sampleFormulas: formulaView.slice(0, 8),
        tables: [],
      };
    },
    async readRange(addr: string, maxRows: number): Promise<ReadResult> {
      return {
        address: addr,
        values: values.slice(0, maxRows),
        truncated: values.length > maxRows,
        totalRows: values.length,
      };
    },
    async readFormulas(addr: string, maxRows: number): Promise<FormulaReadResult> {
      return {
        address: addr,
        formulas: formulaView.slice(0, maxRows),
        truncated: formulaView.length > maxRows,
        totalRows: formulaView.length,
      };
    },
    async createTable(addr: string) {
      return { address: addr };
    },
    async getSheetNames() {
      return ["Aug"];
    },
  };
}

const call = (name: string, args: Record<string, unknown> = {}): ToolCall => ({ id: "c1", name, args });

describe("address helpers", () => {
  it("converts column indices to letters past Z", () => {
    expect(columnLetter(0)).toBe("A");
    expect(columnLetter(25)).toBe("Z");
    expect(columnLetter(26)).toBe("AA");
    expect(columnLetter(27)).toBe("AB");
    expect(columnLetter(51)).toBe("AZ");
    expect(columnLetter(52)).toBe("BA");
  });

  it("reads the origin out of an A1 address, with or without a sheet", () => {
    expect(startColumnOf("A1:D9")).toBe(0);
    expect(startColumnOf("C5:D9")).toBe(2);
    expect(startColumnOf("Sheet2!AA3")).toBe(26);
    expect(startRowOf("A1:D9")).toBe(0);
    expect(startRowOf("C5:D9")).toBe(4);
    expect(startRowOf("Sheet2!$B$12")).toBe(11);
  });
});

describe("describeColumns", () => {
  it("labels columns with real letters and infers types", () => {
    const cols = describeColumns(
      [
        ["Date", "Vendor", "Amount"],
        ["2026-08-01", "Acme", "1,250.00"],
        ["2026-08-02", "Beta", "(75.00)"],
      ],
      "C5:E40",
    );
    expect(cols.map((c) => c.letter)).toEqual(["C", "D", "E"]);
    expect(cols[0]?.inferredType).toBe("date");
    expect(cols[1]?.inferredType).toBe("text");
    expect(cols[2]?.inferredType).toBe("amount");
    expect(cols[2]?.header).toBe("Amount");
  });

  it("names unheaded columns by letter rather than leaving them blank", () => {
    const cols = describeColumns([["", "x"], ["1", "2"]], "A1:B2");
    expect(cols[0]?.header).toBe("(column A)");
  });
});

describe("executor dispatch", () => {
  const stageNoop = () => {};

  it("rejects an unknown tool", async () => {
    const exec = createExecutor({ host: fakeHost([]), stageWrite: stageNoop });
    const res = await exec(call("nope"));
    expect(res.isError).toBe(true);
  });

  it("requires an address rather than guessing one", async () => {
    const exec = createExecutor({ host: fakeHost([]), stageWrite: stageNoop });
    expect((await exec(call("read_range"))).isError).toBe(true);
    expect((await exec(call("reconcile_columns", { columnA: 0, columnB: 1 }))).isError).toBe(true);
  });

  it("accepts integer-like strings for column arguments", async () => {
    const exec = createExecutor({
      host: fakeHost([["GL", "Bank"], [100, 100]]),
      stageWrite: stageNoop,
    });
    const res = await exec(call("reconcile_columns", { address: "A1:B2", columnA: "0", columnB: "1" }));
    expect(res.isError).toBeUndefined();
  });

  it("refuses to reconcile a column against itself", async () => {
    const exec = createExecutor({ host: fakeHost([["a", "b"]]), stageWrite: stageNoop });
    const res = await exec(call("reconcile_columns", { address: "A1:B1", columnA: 1, columnB: 1 }));
    expect(res.isError).toBe(true);
  });

  it("rejects out-of-bounds column indices instead of reconciling blanks", async () => {
    const exec = createExecutor({ host: fakeHost([["GL", "Bank"], [1, 1]]), stageWrite: stageNoop });
    const res = await exec(call("reconcile_columns", { address: "A1:B2", columnA: 0, columnB: 7 }));
    expect(res.isError).toBe(true);
    expect(String((res.content as { error: string }).error)).toMatch(/out of bounds/);
  });

  it("maps reconciliation results to real worksheet row numbers", async () => {
    // Range starts at row 5, so the header is sheet row 5 and data starts at 6.
    const host = fakeHost(
      [
        ["GL", "Bank"],
        [100, 100],
        [250, 999],
      ],
      "A5:B7",
    );
    const exec = createExecutor({ host, stageWrite: stageNoop });
    const res = await exec(call("reconcile_columns", { address: "A5:B7", columnA: 0, columnB: 1 }));
    const c = res.content as {
      unmatchedGlRows: number[];
      unmatchedBankRows: number[];
      sampleMatches: Array<{ glRow: number; bankRow: number }>;
      balanced: boolean;
    };
    expect(c.sampleMatches[0]).toMatchObject({ glRow: 6, bankRow: 6 });
    expect(c.unmatchedGlRows).toEqual([7]);
    expect(c.unmatchedBankRows).toEqual([7]);
    expect(c.balanced).toBe(false);
  });

  it("stages a write without touching the sheet", async () => {
    const staged: PendingWrite[] = [];
    const exec = createExecutor({
      host: fakeHost([["old"]]),
      stageWrite: (w) => staged.push(w),
    });
    const res = await exec(call("propose_write", { address: "A1", values: [["new"]], note: "set it" }));
    expect(staged).toHaveLength(1);
    expect(staged[0]?.values).toEqual([["new"]]);
    expect(staged[0]?.before).toEqual([["old"]]);
    expect((res.content as { staged: boolean }).staged).toBe(true);
  });

  it("rejects a ragged write rather than writing a misaligned rectangle", async () => {
    const exec = createExecutor({ host: fakeHost([[]]), stageWrite: stageNoop });
    const res = await exec(call("propose_write", { address: "A1", values: [["a", "b"], ["c"]], note: "x" }));
    expect(res.isError).toBe(true);
  });

  it("rejects values that are not an array of rows", async () => {
    const exec = createExecutor({ host: fakeHost([[]]), stageWrite: stageNoop });
    expect((await exec(call("propose_write", { address: "A1", values: "nope", note: "x" }))).isError).toBe(true);
    expect((await exec(call("propose_write", { address: "A1", values: ["a"], note: "x" }))).isError).toBe(true);
  });

  it("caps read_range rather than marshalling an unbounded selection", async () => {
    const big = Array.from({ length: 9000 }, (_, i) => [i]);
    const exec = createExecutor({ host: fakeHost(big), stageWrite: stageNoop });
    const res = await exec(call("read_range", { address: "A:A", maxRows: 999999 }));
    const c = res.content as { values: unknown[]; truncated: boolean };
    expect(c.values.length).toBe(5000);
    expect(c.truncated).toBe(true);
  });

  it("refuses to reconcile a range past the row ceiling instead of silently truncating", async () => {
    const big = Array.from({ length: 9000 }, () => [1, 1]);
    const exec = createExecutor({ host: fakeHost(big), stageWrite: stageNoop });
    const res = await exec(call("reconcile_columns", { address: "A1:B9000", columnA: 0, columnB: 1 }));
    expect(res.isError).toBe(true);
    expect(String((res.content as { error: string }).error)).toMatch(/Narrow the address/);
  });
});

describe("aggregateRange (host-side group-by)", () => {
  const sheet = [
    ["Category", "Vendor", "Amount"],
    ["Travel", "AirCo", "1,250.00"],
    ["Meals", "Cafe", "(75.50)"],
    ["Travel", "Taxi", "$45.00"],
    ["Meals", "Deli", "n/a?"], // unparseable on purpose
    ["Travel", "Rail", ""], // blank = absence, not a failure
  ];

  it("groups by column and sums exact integer cents", () => {
    const agg = aggregateRange(sheet, 0, 2);
    const travel = agg.groups.find((g) => g.key === "Travel");
    const meals = agg.groups.find((g) => g.key === "Meals");
    expect(agg.totalGroups).toBe(2);
    expect(travel).toMatchObject({ count: 3, sumCents: 129500, minCents: 4500, maxCents: 125000, parsedCount: 2 });
    expect(meals).toMatchObject({ count: 2, sumCents: -7550 });
    expect(agg.grandSumCents).toBe(121950);
  });

  it("reports unparseable amounts with their 1-based sheet rows instead of dropping them silently", () => {
    const agg = aggregateRange(sheet, 0, 2);
    expect(agg.parseFailures).toEqual([{ row: 5, raw: "n/a?" }]);
  });

  it("groups blank keys as (blank) rather than collapsing them into another group", () => {
    const agg = aggregateRange([["K", "V"], ["", "1"], ["A", "2"], [null, "3"]], 0, 1);
    expect(agg.groups.find((g) => g.key === "(blank)")?.sumCents).toBe(400);
    expect(agg.totalGroups).toBe(2);
  });

  it("sorts groups by sum descending and truncates the group list past 200", () => {
    const big = [["K", "V"], ...Array.from({ length: 250 }, (_, i) => [`g${i}`, i])];
    const agg = aggregateRange(big, 0, 1);
    expect(agg.groups.length).toBe(200);
    expect(agg.totalGroups).toBe(250);
    expect(agg.groups[0]?.sumCents ?? 0).toBeGreaterThanOrEqual(agg.groups[1]?.sumCents ?? 0);
  });
});

describe("aggregate_range dispatch", () => {
  const stageNoop = () => {};

  it("validates arguments and bounds before aggregating", async () => {
    const exec = createExecutor({ host: fakeHost([["A", "B", "C"], [1, 2, 3]]), stageWrite: stageNoop });
    expect((await exec(call("aggregate_range", { groupBy: 0, valueColumn: 2 }))).isError).toBe(true);
    expect((await exec(call("aggregate_range", { address: "A1:C2" }))).isError).toBe(true);
    expect((await exec(call("aggregate_range", { address: "A1:C2", groupBy: 2, valueColumn: 2 }))).isError).toBe(true);
    const oob = await exec(call("aggregate_range", { address: "A1:C2", groupBy: 0, valueColumn: 9 }));
    expect(String((oob.content as { error: string }).error)).toMatch(/out of bounds/);
  });

  it("returns per-group totals the model can quote without seeing raw rows", async () => {
    const sheet = [
      ["Category", "Vendor", "Amount"],
      ["Travel", "AirCo", "1,250.00"],
      ["Meals", "Cafe", "(75.50)"],
      ["Travel", "Taxi", "$45.00"],
      ["Meals", "Deli", "n/a?"],
      ["Travel", "Rail", ""],
    ];
    const exec = createExecutor({ host: fakeHost(sheet), stageWrite: stageNoop });
    const res = await exec(call("aggregate_range", { address: "A1:C6", groupBy: 0, valueColumn: 2 }));
    const c = res.content as {
      grandSum: string;
      groups: Array<{ key: string; count: number; sum: string; min: string | null; max: string | null; avg: string | null }>;
      unreadableCount: number;
      summary: string;
    };
    expect(c.grandSum).toBe("1219.50");
    expect(c.groups[0]).toMatchObject({ key: "Travel", count: 3, sum: "1295.00", min: "45.00", max: "1250.00", avg: "647.50" });
    expect(c.unreadableCount).toBe(1);
    expect(c.summary).toMatch(/could not be parsed/);
  });

  it("refuses to aggregate a range past the row ceiling", async () => {
    const big = Array.from({ length: 9000 }, () => ["a", 1]);
    const exec = createExecutor({ host: fakeHost(big), stageWrite: stageNoop });
    const res = await exec(call("aggregate_range", { address: "A1:B9000", groupBy: 0, valueColumn: 1 }));
    expect(res.isError).toBe(true);
    expect(String((res.content as { error: string }).error)).toMatch(/Narrow the address/);
});
});

describe("formula parity", () => {
  const stageNoop = () => {};

  it("read_formulas returns the formula view and distinguishes live cells from constants", async () => {
    const exec = createExecutor({
      host: fakeHost(
        [["Qty", "Total"], [2, 300], [3, 450]],
        "A1:B3",
        [["Qty", "Total"], ["2", "=A2*150"], ["3", "=A3*150"]],
      ),
      stageWrite: stageNoop,
    });
    const res = await exec(call("read_formulas", { address: "A1:B3" }));
    expect(res.isError).toBeUndefined();
    const c = res.content as { formulas: string[][]; totalRows: number; note: string };
    expect(c.formulas[1]?.[1]).toBe("=A2*150");
    expect(c.formulas[1]?.[0]).toBe("2"); // constant comes back as a literal
    expect(c.totalRows).toBe(3);
    expect(c.note).toMatch(/live formulas/);
  });

  it("read_formulas requires an address and caps rows", async () => {
    const exec = createExecutor({ host: fakeHost([["x"]]), stageWrite: stageNoop });
    expect((await exec(call("read_formulas"))).isError).toBe(true);
    const res = await exec(call("read_formulas", { address: "A1", maxRows: 999999 }));
    expect((res.content as { formulas: unknown[] }).formulas.length).toBe(1);
  });

  it("get_sheet_context surfaces the formula sample so the model sees live cells", async () => {
    const exec = createExecutor({
      host: fakeHost([["Total"], [300]], "A1:B2", [["Total"], ["=SUM(C2:C2)"]]),
      stageWrite: stageNoop,
    });
    const res = await exec(call("get_sheet_context"));
    const c = res.content as { sampleFormulas?: string[][] };
    expect(c.sampleFormulas?.[1]?.[0]).toBe("=SUM(C2:C2)");
  });

  it("propose_write stages '=' strings as live formulas and says how many", async () => {
    const staged: PendingWrite[] = [];
    const exec = createExecutor({ host: fakeHost([["old"]]), stageWrite: (w) => staged.push(w) });
    const res = await exec(
      call("propose_write", { address: "A1:B1", values: [["=SUM(C:C)", 42]], note: "add total" }),
    );
    expect(res.isError).toBeUndefined();
    expect(staged[0]?.values[0]?.[0]).toBe("=SUM(C:C)"); // passes through untouched
    const c = res.content as { formulaCells?: number; message: string };
    expect(c.formulaCells).toBe(1);
    expect(c.message).toMatch(/live formulas/);
  });

  it("propose_write omits the formula count when there are no formulas", async () => {
    const exec = createExecutor({ host: fakeHost([["old"]]), stageWrite: () => {} });
    const res = await exec(call("propose_write", { address: "A1", values: [["plain"]], note: "x" }));
    const c = res.content as { formulaCells?: number };
    expect(c.formulaCells).toBeUndefined();
  });
});

describe("sort_range", () => {
  const stageNoop = () => {};
  const sheet = [
    ["Vendor", "Amount", "Note"],
    ["AirCo", 30, "x"],
    ["Cafe", 10, "y"],
    ["Deli", 20, "z"],
  ];

  it("sorts ascending by a column offset, keeping the header and rows intact", async () => {
    const staged: PendingWrite[] = [];
    const exec = createExecutor({ host: fakeHost(sheet, "A1:C4"), stageWrite: (w) => staged.push(w) });
    const res = await exec(call("sort_range", { address: "A1:C4", column: 1 }));
    expect(res.isError).toBeUndefined();
    const c = res.content as { staged: boolean; rows: number; message: string };
    expect(c.staged).toBe(true);
    expect(c.rows).toBe(3);
    expect(c.message).toMatch(/awaiting their approval/);
    // Header untouched, rows whole, ordered 10/20/30.
    expect(staged[0]?.values[0]).toEqual(["Vendor", "Amount", "Note"]);
    expect(staged[0]?.values.slice(1).map((r) => r[1])).toEqual([10, 20, 30]);
    expect(staged[0]?.values[1]).toEqual(["Cafe", 10, "y"]);
    // The sheet is not mutated until approval: before keeps the original order.
    expect(staged[0]?.before[1]).toEqual(["AirCo", 30, "x"]);
  });

  it("sorts descending with blanks sinking to the bottom either way", async () => {
    const blanked = [
      ["Name", "Score"],
      ["Ann", 5],
      ["Bo", null],
      ["Cy", 9],
    ];
    const staged: PendingWrite[] = [];
    const exec = createExecutor({ host: fakeHost(blanked, "A1:B4"), stageWrite: (w) => staged.push(w) });
    await exec(call("sort_range", { address: "A1:B4", column: 1, direction: "desc" }));
    const col = staged[0]?.values.slice(1).map((r) => r[1]);
    expect(col).toEqual([9, 5, null]); // blank last even descending
  });

  it("compares numeric-looking strings as numbers, not text", async () => {
    const textual = [
      ["ID", "Amt"],
      ["a", "100"],
      ["b", "20"],
      ["c", "3"],
    ];
    const staged: PendingWrite[] = [];
    const exec = createExecutor({ host: fakeHost(textual, "A1:B4"), stageWrite: (w) => staged.push(w) });
    await exec(call("sort_range", { address: "A1:B4", column: 1 }));
    expect(staged[0]?.values.slice(1).map((r) => r[1])).toEqual(["3", "20", "100"]);
  });

  it("validates the column offset and direction", async () => {
    const exec = createExecutor({ host: fakeHost(sheet, "A1:C4"), stageWrite: stageNoop });
    const badCol = await exec(call("sort_range", { address: "A1:C4", column: 7 }));
    expect(badCol.isError).toBe(true);
    expect(String((badCol.content as { error: string }).error)).toMatch(/outside/);
    const negCol = await exec(call("sort_range", { address: "A1:C4", column: -1 }));
    expect(negCol.isError).toBe(true);
    const badDir = await exec(call("sort_range", { address: "A1:C4", column: 0, direction: "up" }));
    expect(badDir.isError).toBe(true);
  });

  it("needs an address and refuses a header-only range", async () => {
    const exec = createExecutor({ host: fakeHost(sheet, "A1:C4"), stageWrite: stageNoop });
    expect((await exec(call("sort_range", { column: 0 }))).isError).toBe(true);
    const single = createExecutor({ host: fakeHost([["H"]], "A1:C1"), stageWrite: stageNoop });
    expect((await single(call("sort_range", { address: "A1:C1", column: 0 }))).isError).toBe(true);
  });
});
