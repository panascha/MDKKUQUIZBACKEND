---
name: add-gemini-model
description: Onboard a newly-released Gemini model into the AI_Models registry safely. Use whenever the user says a new Gemini model exists/is available/was released and asks to "add it", "recognize it", or "update configs" for it — e.g. "Google released gemini-X.Y-flash, add it". Also use to explain why a model is stuck Disabled or missing RPD.
---

# Add Gemini Model

The registry (`AI_Models` sheet + `ai-gemini.gs`) already auto-discovers, auto-ranks, and
auto-enables new Gemini models via daily/weekly triggers (`runGeminiModelSync` /
`runGeminiToolProbe`). **Never hand-write a model ID into any sheet or config file** —
Google's live `models.list` is the only source of truth (a past grilling session killed this
exact anti-pattern: fabricated IDs like `gemini-3-flash`/`gemini-3.1-pro` were seeded from an
LLM's guess and had to be purged). Your job is almost always to *verify* the automation already
did the work, not to do it by hand.

## Step 1 — Confirm the model is real (never trust the claim alone)

```
curl -s "<GAS_EXEC_URL>?action=discoverGeminiModels" -L
```

Public, read-only, zero quota cost. Grep the `models[].id` list for the claimed ID. If it's not
there, stop — the model doesn't exist yet on this Google Cloud project, regardless of what a
release announcement or the user says.

`<GAS_EXEC_URL>` = `gas.url` in `claude-kkuintelsphere-router/config.local.json`.

## Step 2 — Check whether it's already onboarded

```
curl -s "<GAS_EXEC_URL>?action=getAIModels" -L
```

Also public/read-only. The daily sync (`runGeminiModelSync` → `reconcileGeminiModels`) and
weekly probe (`runGeminiToolProbe`) run on their own schedule — by the time anyone asks about a
new model, it has usually already been appended, auto-ranked, and probed. Read the row for the
model in question:

- **Not present at all** → reconcile hasn't run since release. It's editor/trigger-only (not
  exposed on `doGet` — the deployment is public no-auth and this action mutates + costs quota),
  so either wait for the daily trigger or run `reconcileGeminiModels()` from the Apps Script
  editor. Do not add rows by hand.
- **Present, `status: "Disabled"`, notes has `auto-discovered <date>`** → correctly gated, no
  tool-probe pass yet. Wait for the weekly probe or re-run `runGeminiToolProbe()` from the
  editor. Do not flip it Active by hand.
- **Present, `status: "Disabled"`, notes starts `tool-incapable:`** → it **failed** the
  mandatory tool-call smoke test (forced `tool_choice`, asserts a `tool_calls` block). This is
  fail-closed by design (agentic callers need tool use) — **but check the reason string before
  trusting it.** `reason: "no tool_calls (finish=length)"` means the model hit the smoke test's
  `max_tokens` cap before emitting the tool call — that's a **false negative** on thinking-by-
  default models (their reasoning tokens eat the budget first), not proof of incapability. A
  genuine capability failure looks like an HTTP 4xx/400 or `finish=stop` with no tool_calls.
  (Found 2026-07-24: `geminiToolSmokeTest_`'s `max_tokens` was 64, too low for `gemini-3.6-flash`
  — raised to 1024 and fixed. If this pattern recurs on a future thinking model, check that
  constant again before writing the model off.) Only leave a model Disabled on a genuine
  capability failure — don't force it live either way without re-probing (editor-only,
  `enableGeminiModel`/`runGeminiToolProbe`, real quota + sheet mutation, needs owner sign-off).
- **Present, `status: "Active"`, `rpd: 0` or blank** → passed the capability gate but nobody has
  confirmed its real free-tier RPD yet. This is the one gate that's deliberately human-only
  (the API exposes no quota info). It is dormant (`getAIModelRegistry_` requires
  `Active && limit>0`) until RPD is set — **ask the owner for the number**, don't guess a quota.
  Set it via `setModelRpd` (POST, owner/admin-authed) once known.
- **Present, `status: "Active"`, `rpd > 0`** → fully live already. Nothing to do.

## Step 3 — The two ranking dimensions, if you do touch the ranker code

Only relevant if you're editing `assignDiscoveredPriority_`/`_parseGeminiRank_` themselves
(rare — this is stable, already-built logic, not something to redo per model):

- **Monotonic structure**: new model rows are *appended*, new `AI_Config` quota columns are
  appended at `getLastColumn()+1`. Existing rows/columns are never inserted, reordered, or
  deleted (concurrent readers rely on stable column indices). Existing hand-set `Priority`
  values are never rewritten — a new model finds its slot via
  `min(existing)-1` / `max(existing)+1` / midpoint-with-integer-gap among the current rows.
- **Version-derived priority**: `^gemini-(\d+)\.(\d+)-(flash-lite|flash|pro)$` (check
  `flash-lite` before `flash` or it never matches) → sort key `(tierRank asc, major desc, minor
  desc)`, `tierRank: flash=0, flash-lite=1, pro=2`. Anything that doesn't cleanly parse
  (`-preview`, `-latest`, date-stamped, `-customtools`, unknown suffix) is a **fail-safe**:
  worst priority + `Disabled` + `needs-manual-priority` note. Never guess a rank for a curveball
  ID — that note is a deliberate stop sign for a human, not a bug.

## Step 4 — Static config that does NOT auto-update

Two places still need a manual touch, and only for models that already passed Step 2's gates:

- **`CONVERTER_FALLBACK_MODELS`** (`ai-gemini.gs`, student PDF/MCQ converter) — an
  emergency-only backup chain used *solely* if the `AI_Models` sheet can't be read; the normal
  path (`apiKeyInfo.fallbackModels`) is already built dynamically from Active registry rows, so
  this constant is cosmetic 99% of the time. Still worth appending a newly-Active, RPD-set model
  to it for the degraded-mode case. The converter doesn't call tools, so a `tool-incapable`
  model is still fine to add *here specifically* (but not anywhere agentic).
- **`AGENT_QUERY_MODEL_OVERRIDE.Gemini`** (`intelsphere.gs`) — the single static model used by
  the *separate* IntelSphere-donor-key Gemini tier (not the per-model `AI_Models`/`AI_Config`
  registry, and not the personal-Gemini terminal fallback tier). This one **requires tool
  support** (agentic callers can't work without it) — only point it at a model that already
  passed the Step 2 tool-probe gate. Changing which model is the flagship here is a deliberate
  owner choice, not something to do reflexively just because a newer ID exists.

Router side (`claude-kkuintelsphere-router/src/`) never needs a change for a new model — the
dashboard reads `AI_Models`/`AI_Config` live via `aiConfigStatus`/`agentPoolStatus` and renders
whatever columns exist.

## Step 5 — Deploy

```
cd "<repo>/MDKKU52QUIZ CODEBASE/MDKKUQUIZBACKEND"
clasp push --force
clasp version "<description>"
clasp deploy -i <live-deployment-id> -V <new-version-number>   # never bare `clasp deploy`
```

Get `<live-deployment-id>` from `clasp deployments` (the one matching `gas.url` in the router's
`config.local.json`) or from CLAUDE.md. Verify afterward with the same `getAIModels`/
`aiConfigStatus` GET call from Step 1/2 — both are public read-only, safe to hit post-deploy.

## What NOT to do

- Don't hand-write a model ID into `AI_MODELS_DEFAULTS`, the sheet, or any static list without
  first confirming it via live `discoverGeminiModels` — that array only seeds a *blank* sheet
  and drifts from live state by design once real data exists; editing it has no effect on
  production and isn't the source of truth either way.
- Don't call `enableGeminiModel`/`reconcileGeminiModels`/`runGeminiToolProbe`/
  `verifyToolSmokeTest`/`verifyRpmCooldown` over HTTP — they're intentionally not exposed on the
  public `doGet` (some mutate, some burst the shared student pool). Editor or trigger only.
- Don't fabricate an RPD number. If it's missing, that's the owner's irreducible manual step —
  ask, don't estimate.
- Don't flip a `tool-incapable` model to Active by hand to "fix" a missing feature — and don't
  take the note at face value either. Read the `reason` string first: `finish=length` is
  probably the smoke test's `max_tokens` being too small for a thinking model (false negative,
  fixable by re-probing), while a 4xx/`finish=stop` is a genuine failure that should stay
  Disabled. Either way, only an actual re-probe (`enableGeminiModel`, editor-only, real quota +
  sheet mutation) resolves it — never hand-write `Status=Active`.
