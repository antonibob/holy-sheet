import { describe, it, expect } from "vitest";
import { createExecutor, type Host } from "../src/office/executor.js";
import { splitQualifiedAddress } from "../src/office/excelHost.js";
import type { ToolCall } from "../src/model/tools.js";

describe("splitQualifiedAddress", () => {
  it("returns null sheet for a bare range", () => {
    expect(splitQualifiedAddress("A1:C10")).toEqual({ sheetName: null, range: "A1:C10" });
  });

  it("splits a sheet-qualified range", () => {
    expect(splitQualifiedAddress("Sheet2!A1:C10")).toEqual({ sheetName: "Sheet2", range: "A1:C10" });
  });

  it("handles names with spaces and dollar signs", () => {
    expect(splitQualifiedAddress("P&L 2026!$B$2:$D$9")).toEqual({
      sheetName: "P&L 2026",
      range: "$B$2:$D$9",
    });
  });

  it("only splits on the first bang", () => {
    expect(splitQualifiedAddress("Weird!Name!A1")).toEqual({ sheetName: "Weird", range: "Name!A1" });
  });
});

function hostWithSheets(names: string[]): Host {
  return {
    async getSheetContext() {
      return {
        sheetName: names[0] ?? "Sheet1",
        usedRangeAddress: null,
        rowCount: 0,
        columnCount: 0,
        selectionAddress: null,
        sample: [],
        tables: [],
      };
    },
    async getSheetNames() {
      return names;
    },
    async readRange(address: string) {
      return { address, values: [], truncated: false, totalRows: 0 };
    },
    async readFormulas(address: string) {
      return { address, formulas: [], truncated: false, totalRows: 0 };
    },
    async createTable(address: string) {
      return { address };
    },
  };
}

const call = (name: string, args: Record<string, unknown> = {}): ToolCall => ({ id: "c1", name, args });

describe("list_sheets tool", () => {
  it("returns every worksheet name", async () => {
    const stageNoop = async (w: unknown) => w;
    const ex = createExecutor({ host: hostWithSheets(["Aug", "Sep", "Summary"]), stageWrite: stageNoop });
    const res = await ex(call("list_sheets"));
    expect(res.content).toEqual({ names: ["Aug", "Sep", "Summary"] });
  });

  it("reports a single-sheet workbook", async () => {
    const stageNoop = async (w: unknown) => w;
    const ex = createExecutor({ host: hostWithSheets(["Sheet1"]), stageWrite: stageNoop });
    const res = await ex(call("list_sheets"));
    expect(res.content).toEqual({ names: ["Sheet1"] });
  });
});
