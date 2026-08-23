---
name: add-intelsphere-provider
description: Wire a brand-new AI provider — exposed through KKU IntelSphere's shared OpenAI-compatible gateway (gen.ai.kku.ac.th) — into the student chatbot rotation and/or the agentQuery (Claude Code proxy) priority chain. Use when IntelSphere starts fronting a new company/provider and its model IDs show up in the live catalog, or when GAS logs show "[IntelSphere] Unclassified model IDs: ...". NOT for a new model from an already-known provider (that's automatic, see the catalog). NOT for a provider integrated *outside* IntelSphere's gateway with its own key pool/endpoint/quota tracking (like Gemini's AI_Config/AI_Models system) — that's a separate, much bigger subsystem, not a config-only add; see add-gemini-model instead.
---

# Add IntelSphere Provider

IntelSphere (`gen.ai.kku.ac.th`) is a single OpenAI-compatible gateway that fronts many
upstream AI companies behind one endpoint and one API key format. Adding a provider here means
teaching this backend's registries about a new *company name* the gateway already routes to —
it is entirely config/constants work in `maintenance.gs` + `intelsphere.gs`, never a new
integration. Confirm the model IDs are real via IntelSphere's own catalog before writing
anything — never hand-write a model ID from a release announcement or the user's claim.

## Step 0 — Confirm this is actually the "new provider" case

Two things must both be true, or you're in the wrong skill:

- The new company's models are served **through IntelSphere** (same endpoint, same donor-key
  auth) — not a standalone integration with its own key pool, its own endpoint, or its own
  per-model quota sheet. Gemini already has that heavier shape (`ai-gemini.gs`, `AI_Config`/
  `AI_Models` sheets, its own RPD tracking) — that pattern is **not** what you're building for a
  new IntelSphere-fronted provider; don't create a parallel version of it reflexively.
- The provider genuinely isn't recognized yet. `inferProviderFromModel` (`intelsphere.gs:2-16`)
  classifies every model ID by regex prefix; if the new company's IDs fail every branch it
  returns `null` and `getIntelSphereModelCatalog` (`intelsphere.gs:35-77`) logs
  `"[IntelSphere] Unclassified model IDs: ..."` (line 66) instead of dropping them silently —
  check GAS logs / Stackdriver for that string first.

## Step 1 — Confirm the models are real, live, from IntelSphere itself

```
POST <GAS_EXEC_URL>  body: { "action": "listModels" }
```

Public, lock-free (`router-doPost.gs:174-180`), zero quota cost. Returns
`{ catalog: getIntelSphereModelCatalog(), donors: ... }`. `catalog` is `{ providerName:
[modelId, ...] }`, built live from IntelSphere's own `GET /api/v1/models` (`intelsphere.gs:47`,
cached 6h). If the new provider's models aren't classified into a proper key yet, they will
either be missing entirely (still `null` from Step 0) or bucketed under whatever the fallback
constant currently says (`PROVIDER_MODELS_FALLBACK` only kicks in if the live fetch itself
fails, `intelsphere.gs:70-73`) — don't confuse a classification gap with a fetch failure.

## Step 2 — Classify the new provider's model IDs

Add one regex branch to `inferProviderFromModel` (`intelsphere.gs:2-16`), matching the existing
style — anchor on the model ID prefix pattern IntelSphere actually returns (verified in Step 1),
not a guessed convention:

```js
if (/^newprovider-/i.test(modelId))  return "NewProviderDisplayName";
```

Order matters only where prefixes could collide (e.g. `mistral-`/`codestral-`/`devstral-` are
grouped into one `Mistral` branch already — check for prefix overlap with an existing branch
before adding a new one).

## Step 3 — Add the provider to the registry constants (`maintenance.gs:850-900` area)

All four must be updated together, or the provider silently falls out of rotation
(`getActiveIntelSphereKey`, `intelsphere.gs:206-209`, skips any provider whose
`{Provider}_Remaining` column can't be resolved — no error, just quietly excluded):

- **`INTELSPHERE_LIMITS`** (`maintenance.gs:852-857`) — daily token quota estimate for the new
  provider's free/donor tier. This is a **human-confirmed number**, same rule as Gemini RPD —
  IntelSphere doesn't expose quota introspection; ask the owner or check the provider's own
  published free-tier limit, don't guess.
- **`INTELSPHERE_PROVIDER_PRIORITY`** (`maintenance.gs:859-861`) — the student-chatbot rotation
  order. This list is **cost/quota-first** (cheapest-feeling providers earlier), separate from
  and unrelated to the agent's priority list in Step 5. Append, don't reorder existing entries
  reflexively — order here reflects a deliberate cost ladder the owner already tuned.
- **`PROVIDER_MODEL_MAP`** (`maintenance.gs:865-870`) — one flagship model ID for the new
  provider. Used only as the fallback model when rotation lands on this provider but the
  student didn't request one of its models by name (`executeChatbotQuery`, `intelsphere.gs:262`
  — "rotation moved provider, we don't know which of *this* provider's models they wanted, so
  use flagship").
- **`PROVIDER_MODELS_FALLBACK`** (`maintenance.gs:873+`) — the new provider's model ID list,
  used **only** when IntelSphere's live `GET /models` fetch itself fails
  (`intelsphere.gs:70-73`). Populate with real IDs confirmed in Step 1 — this array is a
  degraded-mode safety net, not the source of truth, but a stale/fabricated list here still
  actively serves wrong model IDs during an outage.

## Step 4 — Add the sheet column (manual, on the *live* spreadsheet)

`setupIntelSphereSheet` (`intelsphere.gs:652-694`) only writes its hardcoded header list
(`intelsphere.gs:662-667`) when row 1 is completely empty (`if (!firstCell)`, line 672) — on
the real production sheet, which already has data, calling it again does **not** add a missing
column. It does, however, return a live diagnostic:

```js
setupIntelSphereSheet()   // run from the Apps Script editor
// → { ..., missingProviderColumns: [...] }
```

Any provider in `INTELSPHERE_PROVIDER_PRIORITY` without a matching `{Provider}_Remaining`
header shows up here. **Manually add that column header to the live `IntelSphere_Keys` sheet**
(append at the end, past `Notes` or wherever the current last column is) before any of the
runtime code below can read or write it — `getActiveIntelSphereKey`
(`intelsphere.gs:206-209`), `seedIntelSphereKey`'s `setCol` (`intelsphere.gs:736`, no-ops
silently if the column doesn't exist), and the daily reset in `getActiveIntelSphereKey`
(`intelsphere.gs:191-197`) all depend on the header already being there.

Once the column exists, re-seed or wait for the next daily reset — new/existing donor key rows
pick up the new `{Provider}_Remaining` = `INTELSPHERE_LIMITS[provider]` automatically on their
next reset pass (`intelsphere.gs:191-197`), no per-row manual edit needed.

## Step 5 — Optional: wire into the agentQuery (Claude Code proxy) chain

Steps 1-4 alone make the new provider available to the **student chatbot** only
(`executeChatbotQuery`). The owner's `agentQuery` proxy chain is a deliberately separate,
capability-first priority list — adding a provider to the student chain does **not**
automatically add it here, and shouldn't be done reflexively:

- **`AGENT_QUERY_PROVIDER_PRIORITY`** (`intelsphere.gs:362`) — append only if the owner wants
  Claude Code agentic traffic to actually use this provider.
- **`AGENT_PROVIDER_CONTEXT`** (`intelsphere.gs:365`) — a conservative context-window estimate
  (tokens) for the new provider's flagship, used by `orderAgentProviders`
  (`intelsphere.gs:445-467`) to filter out providers whose window can't fit the request before
  ranking the rest by quota. Missing from this map falls back to a conservative 128k
  (`intelsphere.gs:451`) rather than erroring — fine as a stopgap, but confirm the real number.
- **`AGENT_QUERY_OVERFLOW_PROVIDERS`** (`intelsphere.gs:368`) — set `true` only if this
  provider should be a large-quota *buffer* used after the primary workhorses
  (Deepseek/Qwen/OpenAI), not ranked to the front purely because it has the most quota left
  (mirrors how Gemini/xAI are already flagged).
- **`AGENT_QUERY_MODEL_OVERRIDE`** (`intelsphere.gs:373`) — only needed if the agent-tier model
  should differ from the student flagship in `PROVIDER_MODEL_MAP` (e.g. the student flagship
  lacks tool-use support — agentic callers need `tool_calls`, verify this before overriding or
  leaving it un-set).
- **`getIntelSphereQuotaTotals`** (`intelsphere.gs:386-439`) already iterates
  `AGENT_QUERY_PROVIDER_PRIORITY` generically — no code change needed there once the constant
  above is updated.

## Step 6 — Error parsing: usually nothing to add

Both `executeChatbotQuery` (`intelsphere.gs:291-321`) and `executeAgentQuery`
(`intelsphere.gs:539-562`) classify IntelSphere's HTTP 401 body by **substring match on the
error text**, not by provider name — `"reached daily limit"` → quota exhausted (zero the
column, rotate), `"Invalid model"` → catalog drift (bust cache, retry), anything else → treat
the donor key itself as revoked. This is IntelSphere's own gateway normalizing errors across all
upstream providers uniformly, so a new provider inherits this for free. Also note the
200-wrapped-error guard (`executeAgentQuery`, `intelsphere.gs:518-529`, checking
`body.choices`) — IntelSphere sometimes proxies an upstream provider's error inside an HTTP 200
envelope; this guard is already provider-agnostic too.

Only add provider-specific parsing if you confirm (via a real failing call, not assumption)
that this new provider's IntelSphere wrapper returns genuinely different error text for the
same conditions — don't add a branch speculatively.

## Step 7 — Deploy and verify

```bash
cd "<repo>/MDKKU52QUIZ CODEBASE/MDKKUQUIZBACKEND"
clasp push --force
clasp version "<description>"
clasp deploy -i <live-deployment-id> -V <new-version-number>   # never bare `clasp deploy`
```

Verify with the same `listModels` POST from Step 1 (new provider key should now appear in
`catalog`), and check the router's dashboard (`claude-kkuintelsphere-router`'s `/dashboard`,
backed by `getAgentPoolStatus`, `intelsphere.gs:624-649`) — it already renders **every**
`{Provider}_Remaining` column it finds dynamically, so a correctly-added column shows up with
zero router-side changes needed.

## What NOT to do

- Don't hand-write or guess a model ID for the new provider anywhere — confirm against
  IntelSphere's own live catalog (`listModels` / `getIntelSphereModelCatalog`) first, same rule
  as Gemini onboarding.
- Don't expect `setupIntelSphereSheet()` to add a missing column on the live production sheet —
  it only writes headers to a completely empty sheet. Use its `missingProviderColumns` output
  as a diagnostic, then add the header manually.
- Don't build a parallel per-model registry (like Gemini's `AI_Models`/`AI_Config`) for a
  provider that's actually fronted by IntelSphere — that's solving a problem this provider
  doesn't have. IntelSphere's own per-provider, non-per-model `{Provider}_Remaining` column is
  the right granularity here.
- Don't add the new provider to `AGENT_QUERY_PROVIDER_PRIORITY` just because it's in
  `INTELSPHERE_PROVIDER_PRIORITY` — the two lists serve different goals (student cost-ladder vs.
  owner capability-first) and are populated independently, by deliberate choice each time.
- Don't fabricate an `INTELSPHERE_LIMITS` quota number — it's human-confirmed only, same as
  Gemini RPD; IntelSphere exposes no quota-introspection endpoint.
- Don't add custom error-parsing branches speculatively — the existing substring classification
  in `executeChatbotQuery`/`executeAgentQuery` is already provider-agnostic; only extend it
  after observing a real response that doesn't fit.
