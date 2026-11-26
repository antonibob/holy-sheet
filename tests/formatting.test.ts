import { describe, it, expect } from "vitest";
import {
  createExecutor,
  formatProblems,
  isValidHexColor,
  type PendingFormat,
  type PendingWrite,
} from "../src/office/executor.js";
import { TOOL_NAMES } from "../src/model/tools.js";
import type { ToolCall } from "../src/model/tools.js";

const call = (name: string, args: Record<string, unknown> = {}): ToolCall => ({ id: "c1", name, args });

/** Minimal Host — propose_format never touches the sheet, so nothing is read. */
function fakeHost() {
  return {
    async getSheetContext() {
      throw new Error("not used by propose_format");
    },
    async getSheetNames() {
      return ["Aug"];
    },
    async readRange() {
      throw new Error("not used by propose_format");
    },
    async readFormulas() {
      throw new Error("not used by propose_format");
    },
    async createTable(addr: string) {
      return { address: addr };
    },
  };
}

function makeExecutor() {
  const staged: Array<PendingWrite | PendingFormat> = [];
  const execute = createExecutor({
    host: fakeHost(),
    stageWrite: (w) => staged.push(w),
    stageFormat: (f) => staged.push(f),
  });
  return { execute, staged };
}

describe("isValidHexColor", () => {
  it("accepts 3- and 6-digit hex", () => {
    expect(isValidHexColor("#B00020")).toBe(true);
    expect(isValidHexColor("#F00")).toBe(true);
    expect(isValidHexColor("#f00")).toBe(true);
  });
  it("rejects named colours and malformed hex", () => {
    expect(isValidHexColor("red")).toBe(false);
    expect(isValidHexColor("#12")).toBe(false);
    expect(isValidHexColor("#GGGGGG")).toBe(false);
    expect(isValidHexColor(42)).toBe(false);
  });
});

describe("formatProblems", () => {
  it("is empty for a valid static format", () => {
    expect(formatProblems({ fill: "#FFF3CD", bold: true, numberFormat: "#,##0.00" })).toEqual([]);
  });
  it("rejects an empty spec", () => {
    expect(formatProblems({})).toHaveLength(1);
  });
  it("rejects non-hex colours", () => {
    expect(formatProblems({ fill: "yellow" })).toHaveLength(1);
  });
  it("rejects a blank number format", () => {
    expect(formatProblems({ numberFormat: "  " })).toHaveLength(1);
  });
  it("rejects a conditional rule with no visible style", () => {
    expect(formatProblems({ bold: true }, { when: "negative", style: { bold: true } })).toHaveLength(1);
  });
  it("accepts a conditional rule with a fill", () => {
    expect(formatProblems({ bold: true }, { when: "negative", style: { fill: "#FCC" } })).toEqual([]);
  });
  it("rejects an unknown conditional trigger", () => {
    expect(
      formatProblems({ bold: true }, { when: "weird" as never, style: { fill: "#FCC" } }),
    ).toHaveLength(1);
  });
});

describe("propose_format tool", () => {
  it("is registered in the tool list", () => {
    expect(TOOL_NAMES).toContain("propose_format");
  });

  it("stages a static format without touching the sheet", async () => {
    const { execute, staged } = makeExecutor();
    const res = await execute(
      call("propose_format", {
        address: "C2:C99",
        note: "Two-decimal money",
        format: { numberFormat: "#,##0.00", bold: true },
      }),
    );
    expect(res.isError).toBeUndefined();
    expect((res.content as { staged: boolean }).staged).toBe(true);
    expect(staged).toHaveLength(1);
    const pending = staged[0] as PendingFormat;
    expect(pending.address).toBe("C2:C99");
    expect(pending.format).toEqual({ numberFormat: "#,##0.00", bold: true });
    expect(pending.id).toMatch(/^f\d+$/);
  });

  it("stages a conditional rule alongside static formatting", async () => {
    const { execute, staged } = makeExecutor();
    const res = await execute(
      call("propose_format", {
        address: "Sheet2!B2:B99",
        note: "Red negatives",
        format: { numberFormat: "#,##0.00" },
        conditional: { when: "negative", style: { fontColor: "#B00020" } },
      }),
    );
    expect(res.isError).toBeUndefined();
    const pending = staged[0] as PendingFormat;
    expect(pending.address).toBe("Sheet2!B2:B99");
    expect(pending.conditional?.when).toBe("negative");
    expect(pending.conditional?.style.fontColor).toBe("#B00020");
  });

  it("fails without an address", async () => {
    const { execute, staged } = makeExecutor();
    const res = await execute(call("propose_format", { format: { bold: true }, note: "x" }));
    expect(res.isError).toBe(true);
    expect(staged).toHaveLength(0);
  });

  it("fails when no formatting property is set", async () => {
    const { execute, staged } = makeExecutor();
    const res = await execute(call("propose_format", { address: "A1:B2", note: "x", format: {} }));
    expect(res.isError).toBe(true);
    expect((res.content as { error: string }).error).toMatch(/at least one/i);
    expect(staged).toHaveLength(0);
  });

  it("fails on a non-hex colour", async () => {
    const { execute, staged } = makeExecutor();
    const res = await execute(
      call("propose_format", { address: "A1", note: "x", format: { fill: "yellow" } }),
    );
    expect(res.isError).toBe(true);
    expect((res.content as { error: string }).error).toMatch(/hex/i);
    expect(staged).toHaveLength(0);
  });

  it("fails a conditional rule with no visible style", async () => {
    const { execute, staged } = makeExecutor();
    const res = await execute(
      call("propose_format", {
        address: "A1",
        note: "x",
        format: { bold: true },
        conditional: { when: "negative", style: { bold: true } },
      }),
    );
    expect(res.isError).toBe(true);
    expect(staged).toHaveLength(0);
  });

  it("does not stage writes when propose_format rejects", async () => {
    const { execute, staged } = makeExecutor();
    await execute(call("propose_format", { address: "A1", note: "x", format: { fill: "nope" } }));
    expect(staged).toHaveLength(0);
  });
});
