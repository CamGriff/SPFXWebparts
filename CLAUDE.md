# PnP Multilingual Governance — Build Context

## What this is

An SPFx solution that provides translation governance for multilingual SharePoint Online sites. It detects translation drift between source language pages and their translations, surfaces the results in an interactive dashboard, and sends nudge notifications to translators via Microsoft Graph API.

The goal is a single `.sppkg` file that a tenant administrator can deploy to the App Catalog. No Azure dependency — everything runs inside the customer's M365 tenant.

## Product name

Working name: `pnp-multilingual-governance`. Not yet public — build first, decide on naming and licensing later.

---

## Tech stack

- **SPFx** with React and TypeScript — Heft build system (not gulp)
- **PnP JS** (`@pnp/sp`, `@pnp/logging`, `@pnp/core`) for all SharePoint list and page operations
- **Microsoft Graph API** called directly from SPFx for notifications (`sendMail`, Teams channel messages)
- **No Azure Functions, no Power Automate, no external services**
- Build command: `npm run build`
- Dev server: `npm start`

---

## Architecture

Four SPFx components in one package:

### 1. Application Customiser (`src/extensions/`)
- Fires on every page load on sites where it is deployed
- On first load, checks whether the two required lists exist and creates them if not
- Uses `ListProvisioner.ts` service for idempotent list provisioning
- Should be lightweight — fast check, no UI

### 2. Dashboard Web Part (`src/webparts/`)
- Main user-facing component
- Reads from the `TranslationDrift` list via PnP JS
- Shows summary tiles (Total, In Sync, Stale, Missing, Orphaned, Abandoned)
- Sortable, filterable table of pages with drift status
- Per-row nudge button that calls `NotificationService` to email the translator
- Bulk nudge option — notify all stale/missing for a site
- Cross-site view — aggregate across all registered sites, drill down per site
- Refresh button to reload data

### 3. Config Web Part (`src/webparts/` or property pane of dashboard)
- Allows site owners to register sites to monitor
- Set language per site (default `fr-fr`, extensible to others)
- Set stale threshold in days per site
- Assign translator name and email per site per language
- Reads/writes `GovernanceConfig` list via `ConfigService`

### 4. Drift Detection Engine
- Not a separate web part — a TypeScript service (`DriftService.ts`)
- Called when user clicks "Run scan" button in the dashboard
- Scans Site Pages library across all registered sites
- Matches English source pages to translations via `_SPTranslationSourceItemId`
- Calculates drift status for each page
- Writes results to `TranslationDrift` list (upsert — update existing rows, create new ones)

---

## File structure

```
src/
├── webparts/
│   └── multilingualGovernanceDashboard/
│       ├── MultilingualGovernanceDashboardWebPart.ts
│       ├── components/
│       │   ├── IMultilingualGovernanceDashboardProps.ts
│       │   ├── MultilingualGovernanceDashboard.tsx
│       │   ├── dashboard/
│       │   │   ├── SummaryTiles.tsx
│       │   │   ├── DriftTable.tsx
│       │   │   ├── StatusPill.tsx
│       │   │   └── NudgeButton.tsx
│       │   └── config/
│       │       └── ConfigPanel.tsx
├── extensions/
│   └── multilingualGovernanceCustomiser/
│       ├── MultilingualGovernanceCustomiserApplicationCustomizer.ts
│       └── provisioning/
│           └── ListProvisioner.ts
├── services/
│   ├── DriftService.ts
│   ├── NotificationService.ts
│   └── ConfigService.ts
└── models/
    ├── IDriftItem.ts
    ├── IConfigItem.ts
    └── DriftStatus.ts
```

---

## Data model

### TranslationDrift list

Populated by `DriftService`. One row per source page per language.

| Column | Type | Notes |
|---|---|---|
| Title | Text | English page title |
| DefaultPageTitle | Text | |
| DefaultPageUrl | URL | Link to English source page |
| DefaultPageModified | DateTime | |
| TranslationLanguage | Text | e.g. `fr-fr` |
| TranslationPageUrl | URL | Link to French translation |
| TranslationModified | DateTime | |
| DaysDrift | Number | English modified - French modified in days |
| DriftStatus | Choice | In Sync / Stale / Missing / Orphaned / Abandoned |
| TranslatorName | Text | From GovernanceConfig |
| TranslatorEmail | Text | From GovernanceConfig |
| SiteUrl | Text | Source site URL |
| LastChecked | DateTime | When the drift scan last ran |
| PageGuid | Text | UniqueId of English source page — composite key with TranslationLanguage |
| NudgeSent | Boolean | Whether a nudge notification has been sent |
| NudgeDate | DateTime | When last nudge was sent |

### GovernanceConfig list

One row per site per language.

| Column | Type | Notes |
|---|---|---|
| Title | Text | Site display name |
| SiteUrl | Text | SharePoint site URL |
| Language | Text | e.g. `fr-fr` |
| StaleDays | Number | Days before Stale status — default 7 |
| TranslatorName | Text | |
| TranslatorEmail | Text | |
| IsActive | Boolean | Whether this site is included in scans |

---

## Key technical details

### GUID vs UniqueId — critical

`_SPTranslationSourceItemId` on a French page contains the **UniqueId** of the English source page — NOT the `GUID` field that `Get-PnPListItem` returns by default. These are different values.

In PnP JS, request `UniqueId` explicitly:
```typescript
await sp.web.lists.getByTitle('Site Pages').items
  .select('Title', 'FileRef', 'UniqueId', '_SPTranslationSourceItemId', '_SPIsTranslation', '_SPTranslationLanguage')()
```

Build the French pages lookup keyed on `_SPTranslationSourceItemId.toString().toLowerCase()` and match against English page `UniqueId.toString().toLowerCase()`.

### Translation fields on Site Pages

| Field | Type | Meaning |
|---|---|---|
| `_SPIsTranslation` | Boolean | True on translated pages, False/null on source pages |
| `_SPTranslationLanguage` | Text | Locale code on translated pages e.g. `fr-fr` (lowercase) |
| `_SPTranslationSourceItemId` | Guid | UniqueId of the English source page |
| `_SPTranslatedLanguages` | Text | On source pages — which languages have translations |

### Drift status logic

- **In Sync** — French exists, DaysDrift ≤ StaleDays threshold. If French is newer than English, clamp DaysDrift to 0 (treat as In Sync — English is always source of truth).
- **Stale** — French exists, DaysDrift > StaleDays threshold
- **Missing** — No French page linked to this English source
- **Abandoned** — French page exists (`_SPIsTranslation = true`) but was never published (draft state)
- **Orphaned** — French page exists but its `_SPTranslationSourceItemId` points to an English page that no longer exists

### Filtering rules for source pages

Exclude from drift calculation:
- Folders (items where `FileRef` does not end in `.aspx`)
- Items where `_SPIsTranslation = true` (these are translations, not sources)
- Items in `/fr/` paths
- Items with no Title
- Items where content approval status is not Approved (field: `_ModerationStatus = 0`)

---

## Permissions required

Graph API permissions (admin consent required):
- `Sites.Read.All` — to scan Site Pages across registered sites
- `Mail.Send` — to send nudge emails from the web part
- `ChannelMessage.Send` — for Teams notifications (optional v1)

SharePoint permissions:
- Site owner or SharePoint admin to deploy
- Site owner to register sites and run scans

---

## Build notes

- Heft build system — use `npm run build`, not `gulp`
- `npm start` for local dev server (SharePoint Workbench)
- ESLint is strict — avoid `any` types, use `void` operator on floating promises
- `null` is flagged — use `undefined` instead throughout
- SPFx version: check `package.json` for exact version

---

## What is already built (in a separate project)

A working proof of concept exists as two standalone pieces:

1. **`Get-TranslationDrift.ps1`** — PnP PowerShell script that does the drift calculation and populates the TranslationDrift list. The TypeScript `DriftService.ts` should replicate this logic exactly.

2. **Translation Drift Dashboard SPFx web part** — a simpler single-file React web part that reads the TranslationDrift list and renders the dashboard. The new dashboard component should extend this with cross-site support, additional status types, and the nudge button.

The proof of concept confirmed:
- The UniqueId matching approach works
- The list schema is correct
- The dashboard layout and status pill design works
- `fr-fr` locale code is lowercase in SharePoint

---

## Build order recommendation

1. **Models first** — `IDriftItem.ts`, `IConfigItem.ts`, `DriftStatus.ts`
2. **ListProvisioner** — creates both lists, called by the Application Customiser
3. **Application Customiser** — wires up ListProvisioner on first load
4. **ConfigService** — reads/writes GovernanceConfig list
5. **DriftService** — core drift detection logic, port from PowerShell proof of concept
6. **NotificationService** — Graph API sendMail
7. **Dashboard components** — StatusPill, SummaryTiles, DriftTable, NudgeButton
8. **Main dashboard web part** — wires everything together
9. **Config panel** — site registration and settings UI
