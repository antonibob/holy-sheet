# aireconciler — Dev Roadmap (drives the autonomous dev loop)

Goal: perform as well as Claude in Excel at reconciliation/coding/flux work.
Baseline at roadmap creation: 87/87 vitest tests passing, ~2.3k LOC TS.

## Rules for the dev loop (non-negotiable)
1. Work on a branch named `auto/<short-topic>` — NEVER commit directly to `main`.
2. All existing tests must pass before opening a PR (`npx vitest run`).
3. Every PR: description states what changed, why, and test/benchmark evidence.
4. One roadmap item per cycle. If blocked, open an issue instead of forcing it.
5. Match existing code style. No rewrites of working modules.

## Priority order (pick topmost not yet done)
1. **Formula/tool parity audit**: enumerate what Claude can do in Excel
   (formulas, pivot-style aggregation, conditional formatting, chart data,
   multi-sheet ops, undo-safe edits) and diff against `src/office/executor.ts`
   capabilities. File the gap list as issues, then close them one per cycle.
2. **Determinism & provenance**: every cell write must be traceable to a
   source row/citation (extend `reconcile_provenance.test.ts` coverage).
3. **Error recovery**: agent loop (`src/taskpane/agent.ts`) should retry
   malformed model tool calls with repair prompts instead of failing.
4. **Benchmark harness**: `bench/` with 3 real recon scenarios (AP 3-way,
   CC recon, flux) scored on accuracy + wall time; run in every PR.
5. **Streaming UX**: show model progress/steps live in the taskpane.
6. **Cost/latency**: prompt caching + smaller-model routing for tool calls.

## Done log
- (empty — loop fills this in per merged PR)
