/**
 * Conversation persistence for the taskpane.
 *
 * The pane reloads whenever Excel reopens the workbook or the user switches
 * sheets with "always keep pane open" off — and used to lose the whole
 * conversation with it. These helpers serialise the model-facing transcript to
 * localStorage and restore a sane subset on load.
 *
 * Only plain user/assistant text turns survive a restore. Tool plumbing
 * (tool_call ids, tool replies) is dropped: a restored transcript with
 * assistant tool_calls but no matching tool messages is a malformed request
 * OpenRouter rejects, and replaying stale tool results about sheet state that
 * has since changed would mislead the model anyway.
 */

import type { ChatMessage } from "../model/client.js";

const ROLES = new Set(["user", "assistant"]);

/**
 * The restorable subset of a transcript: user and assistant text turns in
 * order, with any tool plumbing stripped.
 */
export function restorableTurns(messages: ChatMessage[]): ChatMessage[] {
  return messages.filter(
    (m) => ROLES.has(m.role) && m.content.trim() !== "" && !m.toolCalls?.length,
  );
}

/** Serialise a transcript for storage. Returns compact JSON or null when empty. */
export function serializeConversation(messages: ChatMessage[]): string | null {
  const turns = restorableTurns(messages);
  if (turns.length === 0) return null;
  return JSON.stringify(turns);
}

/**
 * Parse stored transcript JSON back into messages. Returns null for anything
 * that is not clearly a stored conversation — corrupt storage must degrade to
 * a fresh chat, never crash the pane.
 */
export function restoreConversation(raw: string | null | undefined): ChatMessage[] | null {
  if (!raw) return null;
  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch {
    return null;
  }
  if (!Array.isArray(parsed)) return null;
  const turns: ChatMessage[] = [];
  for (const item of parsed) {
    if (typeof item !== "object" || item === null) return null;
    const { role, content } = item as { role?: unknown; content?: unknown };
    if (typeof role !== "string" || !ROLES.has(role)) return null;
    if (typeof content !== "string") return null;
    // A blank turn is junk, not corruption — skip it like restorableTurns does.
    if (content.trim() === "") continue;
    turns.push({ role: role as ChatMessage["role"], content });
  }
  return turns.length > 0 ? turns : null;
}
