# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

MDKKUQUIZBACKEND is the Google Apps Script (GAS) backend for MDKKUQUIZ, a medical exam quiz platform for KKU medical students (batch 52). Remote: github.com/panascha/MDKKUQUIZBACKEND. Deployed to Google Sheets as a web app, managed via **clasp** (GAS CLI).

This repo is one of three sibling sub-repos in a parent monorepo (`MDKKUQUIZREAL` student quiz app, `MDKKUQUIZDATABASE` admin dashboard). See the parent folder's `CLAUDE.md` for cross-repo context and the `/deploy`, `/issuelist` slash commands.

## Deployment

Prerequisites: `npm install -g @google/clasp` and `clasp login`.

```bash
clasp push        # push Code.js to Google Apps Script
clasp deploy      # creates a NEW deployment with a NEW URL — avoid unless intentionally changing the URL
```

**To update the live backend without changing the URL** (the normal case — avoids editing `APPSCRIPT_URL` in every frontend), snapshot a version and point the existing deployment at it. `clasp redeploy` is **not a real clasp subcommand** — it errors ("Read-only deployments may not be modified"); use `clasp deploy -i` instead:

```bash
clasp version "<desc>"                                          # snapshot HEAD as an immutable numbered version
clasp deploy -i <deploymentId> -V <versionNumber> -d "<desc>"    # repoint the live deployment at it
```

Get `<deploymentId>` from `clasp deployments` (it's the ID embedded in `window.APPSCRIPT_URL` in each frontend's `js/config.js`). The GAS script ID is in `.clasp.json`.

## Architecture

Source is in `Code.js` (~3570 lines). `appsscript.json` sets `access: ANYONE_ANONYMOUS` — all GET requests are intentionally unauthenticated.

Key constants at the top of `Code.js`:
- `SHEET_ID` — Google Sheets spreadsheet ID (the database)
- `DRIVE_FOLDER_ID` — Google Drive folder for uploaded images
- `VOTE_THRESHOLD_CONFIRM` — vote count to auto-confirm a category change; currently `2` (accepted hijack-risk tradeoff — was briefly raised to `20` then lowered back per explicit user request)
- `REPORT_VOTE_THRESHOLD` — **separate constant**, gates report-vote auto-apply (correction to wrong-answer reports); currently `5`, unrelated to `VOTE_THRESHOLD_CONFIRM`

**Caching**: `CacheService` with a chunked large-value workaround (90KB chunks, since GAS has a 100KB per-key cache limit). Cache TTL is 30 minutes for question/structure data. Cache is invalidated via a version key `v` stored in `PropertiesService`.

**`doGet(e)`** (~line 375) handles read-only, unauthenticated actions: `getStructure` (subjects/categories tree), `getQuestions&subject=X`, `getPendingVotes`, `getAllData` (used by the admin dashboard — currently unauthenticated, see Known Issues).

**`doPost(e)`** (~line 896) uses a **3-tier lock model**, not one global lock:
1. **Lock-free group** — `verifySession`, `askAIExpert`. No `LockService` call at all.
2. **Localized lock group** — `submitVote`, `submitReport`, `voteOnReport`, `deleteSession`, `batchLog`. 15s `tryLock`.
3. **Admin lock group** — question/category/subject/announcement CRUD (`editQuestion`, `deleteQuestion`, `addCategory`, `adminImport`, `updateReportStatus`, `deleteCategory`, `updateCategory`, `deleteGroup`, `updateAccordionGroup`, `addSubject`, `updateSubject`, `deleteSubject`, `addAnnouncement`, `editAnnouncement`, `deleteAnnouncement`). 25s `tryLock`, requires a valid session/admin auth.

**When adding a new POST action, place it in the correct tier deliberately**: lock-free for pure reads, localized for high-frequency small writes, admin for schema/data-mutating writes that need auth. Putting a write in the lock-free group risks races under concurrent load; putting a read in a locked group adds needless contention.

Other key actions: `checkGoogleAuth` (Google SSO login → issues `mdkku_session_token`), `batchLog` (analytics, batched — the old per-event `logUserActivity` endpoint no longer exists).

Auth uses KKU Google accounts (`@kkumail.com` or `@kku.ac.th`). GAS checks an admin whitelist sheet and issues its own session tokens stored in a "Sessions" Google Sheet. `generateSessionToken` still uses `Math.random()`, not `Utilities.getUuid()`.

## Data Format

Questions use `///` as a multi-value separator (`img`, `choices` columns); `category` is stored as a JSON array. Never use commas or pipes inside these fields — frontends split on `///` literally.

## Key Invariants

- **`doPost` lock tiers are load-bearing** — don't add a new write action without picking lock-free/localized/admin deliberately (see Architecture above).
- **`VOTE_THRESHOLD_CONFIRM` (2) and `REPORT_VOTE_THRESHOLD` (5) are independent constants** — don't conflate them when touching vote or report logic.
- **`///` is the data delimiter** for multi-value fields.
- Community report-vote auto-apply flow (`processReports()` → `applyReportCorrection()` → Gemini-generated explanation) strips raw Drive image URLs / `<svg` choices to `[รูปภาพ]` before sending to Gemini — don't remove that sanitization, it prevents leaking binary data into the prompt.

## Known Open Issues (this repo)

- 🔴 `doGet`'s `getAllData` is served unauthenticated — no session check before `getAllDataForAdminCached()`.
- 🟡 `generateSessionToken` uses `Math.random()` not `Utilities.getUuid()`.
- Not independently re-verified: `getGeminiConfig()` double `SpreadsheetApp.openById` call and `Usage_Count` write-before-success ordering — check directly if touching Gemini quota logic.

Full cross-repo issue list: parent `Idea/code-review-2026-06-14.md` or `/issuelist`.
