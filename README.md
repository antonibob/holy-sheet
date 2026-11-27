# Holy Sheet — an AI assistant in Excel

A chat-first, general-purpose AI copilot for Excel. Ask it things about your
workbook, and it does the work: reads ranges and formulas, aggregates, sorts,
formats, builds tables, works across sheets — staging every change as a
**reviewable diff** before anything touches a cell. The model is a small, cheap
**BYO-key** (OpenRouter) model; the deterministic work happens host-side.

The add-in **never writes to the sheet directly.** Every edit is staged as a
diff the user confirms with Apply. That discipline is the product: it's the
difference between a tool that breaks your workbook and one you can trust.

Built as an **Office JavaScript Add-in** (React + TypeScript), with the core
logic extracted as pure, unit-tested modules that run without Office.

> **Legacy note:** this repo formerly shipped as a reconciliation copilot
> ("AI Closer"). The reconciliation engine (`engine/`, `reconcile_columns`,
> `find_duplicates`, `src/taskpane/reconcile.ts`) is **parked** — kept green
> and working, but no longer the product focus. Holy Sheet is a general
> assistant; the domain vertical can branch off later.

---

## What it can do today

- **Read & inspect** the sheet: `get_sheet_context`, `read_range`,
  `read_formulas` (live `=` formulas), bounded to 5000 rows.
- **Aggregate** host-side: `aggregate_range` group-by/summarize so thousands of
  rows never enter the model's context.
- **Write** values: `propose_write` stages a diff — Apply/Revert in the UI.
- **Format**: `propose_format` stages static (fill, font, bold, numberFormat)
  and conditional (positive/negative/top-10) formatting — staged and undo-safe.
- **Tables**: `create_table` over an address, staged.
- **Multi-sheet**: `list_sheets` and sheet-qualified ranges (`Sheet2!A1:C10`).
- **Chat UX**: persistent conversation across taskpane reloads, streaming steps,
  tool cards, New Chat.
- **Formulas read**: see exactly what's already computed before proposing
  anything.

## Why this exists

Spreadsheet copilots that dump the whole sheet into the model are slow, wrong
at arithmetic, and unsafe. Holy Sheet flips the division of labor:

| Concern | Owner | Example |
|---|---|---|
| Math, sums, aggregations | **Host** (deterministic, tested) | `aggregate_range` sums server-side |
| Reading the sheet | **Host** | bounded reads, formula extraction |
| Writing / formatting | **User**, via staged diffs | nothing reaches a cell without Apply |
| Judgment, picking tools | **LLM** (BYO-key) | "which column is the vendor?" |

The model reports numbers returned by tools; it never computes totals in its
head. And because the key is yours, confidential workbook data never rides a
shared subscription.

## Current status

- ✅ **Core** — pure TS modules, **113/113 unit tests green**, `tsc` clean
- ✅ **Model layer** — OpenRouter `chat/completions` with tool calling + streaming
- ✅ **Taskpane** — React UI, runs live in a plain browser on a demo dataset
- ✅ **Bundle** — esbuild produces `dist/` (js + css) for sideload
- 🔶 **Office sideload packaging** — manifest present; serve over https + icons TODO

### Verified

```sh
npx vitest run    # 113 passing — core, executor, model, taskpane modules
npm run typecheck # 0 errors (strict)
npm run build     # bundles dist/taskpane.{js,css}
```

The `src/office` adapters and manifest are scaffolded but not exercised in CI
(no Office host on the runner). The core modules they feed are fully tested.

## Quick start — run the live demo (no Office needed)

```sh
npm install
npm run build
npx serve dist        # or: python -m http.server 8137 --directory dist
# open http://localhost:8137/taskpane.html
```

## Architecture

```
manifest.xml              Office Add-in manifest (sideload)
src/
  model/
    client.ts             BYO-key OpenRouter client: tool calling + streaming
    tools.ts              tool schemas + system prompt handed to the model
    openrouter.ts         JSON drafts
  office/
    executor.ts           runs a tool call; host-agnostic, unit-tested
    excelHost.ts          the only code that touches Excel (structural types)
  taskpane/
    agent.ts              the tool-calling loop
    App.tsx               React UI — streaming, tool cards, diff approval
    markdown.tsx          renders replies as React elements (no innerHTML)
    reconcile.ts          (parked legacy) two-column reconcile
    demo.ts               demo dataset
    index.{tsx,css} App.css
  engine/                 (parked legacy) deterministic recon engine, fully tested
  commands/               manifest FunctionFile stub
tests/                    113 tests
scripts/
  build.mjs               esbuild bundle of taskpane + commands entries
```

### How a turn works

The add-in does not guess intent from keywords. The model is given the tools in
`model/tools.ts` and picks; `taskpane/agent.ts` executes what it asked for, feeds
the results back, and repeats until the model answers.

Two invariants hold regardless of what the model says:

- **It cannot do the arithmetic.** Aggregation and any deterministic math run
  host-side, exposed as tools. The model reports returned numbers; it never
  computes them.
- **It cannot write to the sheet.** `propose_write` / `propose_format` stage a
  diff. Nothing reaches a cell until the user clicks Apply.

## Install & run in Excel

Everything needed is wired up. On this machine the system CA already trusts the
localhost cert and the server runs on `https://localhost:3000`.

```sh
npm run build     # typecheck + icons + bundle dist/
npm run serve     # HTTPS server on https://localhost:3000 with the localhost cert
```

Then sideload the add-in into Excel. The shared-folder catalog is the most
reliable path (works regardless of whether the newer "Apps" store shows an
Upload button in your build):

```sh
# one-time: point Excel at a folder that lists your add-in manifests
scripts/setup-shared-folder.bat   # or run the reg add it performs
mkdir -p manifests && cp manifest.xml manifests/   # folder already has it
```

Then in Excel:
1. **File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs**
2. Under "Categories/Trusted catalogs", check **Catalog URL** that points to the
   `manifests` folder, and enable **"Show in Menu"**
3. Restart Excel → **Home → Add-ins → My Add-ins** → the **Holy Sheet** taskpane
   is listed → click it.

The add-in fetches `https://localhost:3000/taskpane.html` (trusted cert) and
mounts the taskpane on `Office.onReady`.

> The add-in also runs as a standalone web demo in any browser at
> `https://localhost:3000/taskpane.html` (no Office required) — Office.js is
> loaded from the CDN but the dual-mode bootstrap skips `Office.onReady` when
> no host is present.

**First-time cert setup** (already done here, but for a fresh clone):

```sh
npm i -D office-addin-dev-certs
npx office-addin-dev-certs install   # generates + trusts localhost cert
npm run serve                        # reuses ~/.office-addin-dev-certs/
```

> Note: `office-addin-dev-certs` needs to write the CA cert into the system
> trust store (admin elevation on Windows) the first time.

## Models & cost

BYO key, pick a cheap model. Default is `deepseek/deepseek-v3-0724`; a normal
session costs pennies because bulk row data is aggregated host-side, never
shipped through the model's context.

## Prior art

Two earlier sketches of the OpenRouter-in-Excel plumbing live in `references/`
(kept as reference, not part of the build):
- `references/OpenRouterExcel.bas` — VBA UDFs, `=OPENROUTER_CHAT(...)` via WinHTTP
- `references/OpenRouterOfficeScript.ts` — Office Script, built-in `fetch`, free Nemotron

This repo is the productized version: those proved the pipe; this adds the
staged-edit core, the tools, and the shell.

## License

MIT
