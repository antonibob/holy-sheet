import { describe, expect, it } from "vitest";
import type { ChatMessage } from "../src/model/client.js";
import {
  restorableTurns,
  restoreConversation,
  serializeConversation,
} from "../src/taskpane/history.js";

const user = (content: string): ChatMessage => ({ role: "user", content });
const assistant = (content: string): ChatMessage => ({ role: "assistant", content });

describe("restorableTurns", () => {
  it("keeps user and assistant text turns in order", () => {
    const turns = restorableTurns([user("hi"), assistant("hello"), user("bye")]);
    expect(turns).toEqual([user("hi"), assistant("hello"), user("bye")]);
  });

  it("strips the system prompt, tool replies and tool-calling assistant turns", () => {
    const messages: ChatMessage[] = [
      { role: "system", content: "sys" },
      user("sum it"),
      { role: "assistant", content: "", toolCalls: [{ id: "t1", name: "read_range", args: {} }] },
      { role: "tool", toolCallId: "t1", content: "{}" },
      assistant("total is 5"),
    ];
    expect(restorableTurns(messages)).toEqual([user("sum it"), assistant("total is 5")]);
  });

  it("drops blank assistant turns", () => {
    expect(restorableTurns([assistant("   ")])).toEqual([]);
  });
});

describe("serializeConversation / restoreConversation", () => {
  it("round-trips user and assistant turns", () => {
    const messages: ChatMessage[] = [
      { role: "system", content: "sys" },
      user("what is in A1?"),
      assistant("The value 42."),
    ];
    const raw = serializeConversation(messages);
    expect(raw).not.toBeNull();
    expect(restoreConversation(raw)).toEqual([user("what is in A1?"), assistant("The value 42.")]);
  });

  it("returns null when there is nothing restorable", () => {
    expect(serializeConversation([{ role: "system", content: "sys" }])).toBeNull();
  });

  it("returns null for junk input instead of throwing", () => {
    expect(restoreConversation(null)).toBeNull();
    expect(restoreConversation(undefined)).toBeNull();
    expect(restoreConversation("")).toBeNull();
    expect(restoreConversation("not json {")).toBeNull();
    expect(restoreConversation("42")).toBeNull();
    expect(restoreConversation('{"role":"user"}')).toBeNull();
  });

  it("rejects arrays containing invalid entries — corrupt storage degrades to fresh", () => {
    expect(restoreConversation('[{"role":"tool","content":"{}"}]')).toBeNull();
    expect(restoreConversation('[{"role":"user","content":5}]')).toBeNull();
    expect(restoreConversation('["hi"]')).toBeNull();
    expect(restoreConversation('[{"role":"user","content":""}]')).toBeNull(); // blank = skipped, array left empty
    expect(restoreConversation('[{"role":"user","content":""},{"role":"assistant","content":"hi"}]'))
      .toEqual([assistant("hi")]);
  });

  it("returns null for an empty restored array", () => {
    expect(restoreConversation("[]")).toBeNull();
  });
});
