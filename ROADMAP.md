# Holy Sheet — Dev Roadmap (drives the autonomous dev loop)

**Holy Sheet** (repo `antonibob/holy-sheet`, formerly aireconciler) is a
general-purpose AI assistant in Excel: chat-first, does whatever Claude in
Excel does — formulas, aggregation, formatting, sorting, charts, multi-sheet
work. Domain-specific reconciliation is PARKED (code stays, no new work) until
general capabilities hit parity; a domain vertical can branch off after that.

Baseline: 87/87 vitest tests passing.

## Rules for the dev loop (non-negotiable)
1. Work on a branch named `auto/<short-topic>` — NEVER commit directly to `main`.
2. All existing tests must pass before opening a PR (`npx vitest run`).
3. Every PR: description states what changed, why, and test/benchmark evidence.
4. One roadmap item per cycle. If blocked, open an issue instead of forcing it.
5. Match existing code style. No rewrites of working modules.
6. General capability first: nothing new that only helps reconciliation.

## Priority order (pick topmost not yet done)
1. **Rename branding** (this cycle): package.json, taskpane title/manifest,
   README — product name is "Holy Sheet".
2. **Formula parity** (issue #1): read cell formulas; extend propose_write to
   emit live `=FORMULA(...)` cells. The sheet must stay recalculable.
3. **Aggregation** (issue #2): group-by/summary computed host-side so the model
   never ships thousands of rows through context.
4. **Chat UX parity**: taskpane feels like a chat assistant — full history,
   streaming steps, follow-up questions about the sheet.
5. **Multi-sheet** (issue #3): list sheets, address ranges across sheets.
6. **Formatting** (issue #4): staged conditional + cell formatting proposals.
7. **Sort/filter** (issue #6): staged sort_range.
8. **Charts** (issue #5): staged chart definitions.
9. **Benchmark harness**: `bench/` with general Excel tasks (not recon
   scenarios), scored on accuracy + wall time; run in every PR.
10. **Error recovery**: agent loop retries malformed tool calls with repair
    prompts instead of failing.
11. **Cost/latency**: prompt caching + smaller-model routing for tool calls.

## Parked (do NOT work on these)
- reconcile_columns, find_duplicates, engine/* reconciliation logic,
  examples/datasets recon scenarios. No new recon tools or tests. Existing
  recon tests stay green — do not delete them.

## Done log
- 2026-11-25 Multi-sheet support (roadmap #5 / issue #3): list_sheets tool,
  sheet-qualified ranges in read/formulas/aggregate/propose/create_table,
  105/105 tests (PR #14, branch auto/multi-sheet).
- 2026-11-23 Parity audit vs Claude-in-Excel (PR #7, issues #1–#6).
  Already at parity: undo-safe staged edits, describeColumns-guided picking,
  bounded reads with 5000-row ceiling.
- 2026-11-23 Pivoted to general "Claude-in-Excel" assistant; recon parked.
