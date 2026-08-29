# B2B Intermediary-Referral KPI Report Generator

Pulls deal, meeting and contact data from HubSpot and writes a styled,
two-sheet Excel workbook (`reports/reports.xlsx`) tracking five
year-to-date stages (Retained Clients, Referred Clients, New
Intermediaries, Presentations, Calls/Meetings) for the Institutional
Relations BDMs (João Pacheco Gonçalves and Rohan Harris), with a
click-to-expand Month → Week → Day drill-down, plus a static reference
sheet of annual KPI targets.

Runs automatically every week via GitHub Actions, committing the
regenerated report back to the repo - see
[Automation](#automation-github-actions) below. It also runs fine
standalone from a local checkout for testing or one-off use.

## Setup

1. **Create a virtual environment and install dependencies:**

   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   pip install -r requirements.txt
   ```

2. **Get a HubSpot token.**

   Preferred: a **Service Key** - Settings → Integrations → Service Keys
   (or Development → Keys → Service Keys, currently in public beta). This
   is HubSpot's recommended credential for a scheduled, data-only script
   with no webhooks or UI extensions.

   Fallback: a legacy-style **Private App** access token (Settings →
   Integrations → Private Apps) works as a drop-in replacement with zero
   code changes - both are used identically as an `Authorization: Bearer`
   header.

   Grant exactly these scopes, no more:

   | Scope | Used for |
   |---|---|
   | `crm.objects.contacts.read` | Contact search/associations, meetings retrieval |
   | `crm.objects.deals.read` | Deal search, batch-read, pipelines/stages, meeting↔deal associations |
   | `crm.objects.owners.read` | Owner list |
   | `crm.schemas.contacts.read` | Contact property metadata |
   | `crm.schemas.deals.read` | Deal property metadata |

3. **Configure the token:**

   ```bash
   cp .env.example .env
   # edit .env and paste the token as HUBSPOT_ACCESS_TOKEN=...
   ```

   `.env` is gitignored and must never be committed.

## Running it

```bash
python generate_report.py
```

The script is idempotent - safe to re-run any time. It overwrites
`reports/reports.xlsx` cleanly on every run and leaves no state behind
between runs. It prints, in order:

1. **Step 0** - every ID/value it resolved from the portal (association
   labels, the lead-source property, pipeline/stage IDs, the
   `lifecyclestage` "Customer" value, and the two BDMs' owner IDs) so you
   can sanity-check them against your portal before trusting the numbers
   below.
2. Each stage's own diagnostics as it's computed - counts, and any
   anomalies it found along the way (referred contacts with no recorded
   introducer, contacts with multiple qualifying deals, meetings already
   booked for a future date, etc.), so a data-quality gap is visible
   immediately rather than buried in a final total.
3. A final **Year-to-Date Summary** - one line per stage with its current
   grand total.

On any failure (missing token, a HubSpot API error, rate-limiting
exhausted after retries) the script exits non-zero with a clear message
on stderr.

### Reading the numbers

Every stage is scoped to **this calendar year, through today** - a
meeting already booked for next week, or a deal signed last year, is
excluded from the current total (the script tells you when it excludes
something like that, and why). Because of this, the totals move day to
day, and don't chase a single fixed "expected" figure the way an
all-time report would. Two things worth knowing if a number looks lower
than you expected:

- **Retained Clients** in particular can look small this early in the
  year - it only counts deals actually signed within the current year, so
  a deal signed last year for a client referred last year won't show up
  even though that client is genuinely retained.
- Any KPI-population gap the script finds (e.g. a Partner-Referral
  contact with no recorded introducer association) is logged with the
  specific contact/deal IDs responsible, not silently dropped.

## What's in the workbook

One tab per Stage, in this order, plus a static reference tab:
**Retained Clients**, **Referred Clients**, **New Intermediaries**,
**Presentations**, **Calls-Meetings** (Excel doesn't allow `/` in a tab
name, so this one tab is named without it; every in-sheet label still
reads "Calls/Meetings"), then **KPI Targets**.

Every Stage tab shares the same layout:

- **Columns**: Month → Week → Day, using Excel's native **column outline
  grouping** - only the Month-total columns are visible by default; click
  the **+** control above a Month to reveal its Week-total columns, and
  click **+** again above a Week to reveal its individual Days. A **YTD
  Total** column follows the last month, always visible. This works in
  Excel, LibreOffice Calc, and Google Sheets (via File → Import or opening
  the .xlsx directly) - look for small **+ / −** buttons in the grey bar
  just above the column headers.
- **Rows**: the two BDMs plus a Grand Total row. Every Week/Month/YTD
  total is a real `SUM()` formula over its own Day/Week/Month cells, never
  a hardcoded number.

The **Retained Clients** tab additionally has:

- **A row-level drill-down per BDM**: each BDM's row expands (via Excel's
  row-level **+/−** outline control, to the left of the row numbers) into
  nested rows broken down by the deal's **Country and Program of
  Interest**, collapsed by default - the same outline-grouping mechanism
  as the column axis, just on rows instead. A BDM's own row is itself a
  `SUM()` of their breakdown rows, not a separately-tracked number.
- **Two annual target rows** below Grand Total, one per BDM, showing a
  literal progressive fraction - "1/12" in January, "2/12" in February,
  and so on through "12/12" in December - against the 12-retained-
  clients-per-year minimum each BDM is individually held to. This is
  static display text, not a formula.

The **Calls-Meetings** tab additionally has a **"% of Target" column**
right after every Month Total column, a real formula (that month's total
÷ 440, the annual per-BDM target) - e.g. `=AL3/440`. The Grand Total
row's version divides by 880 (the sum of both BDMs' individual targets).

**KPI Targets** is static reference content (the minimum annual KPI
targets and the detailed rationale behind them, ordered to match the
Stage tab order) - no API calls.

An in-progress month or week is never padded with future placeholder
columns - it simply has fewer Day columns than a finished one, which is
what makes its total read as "month-to-date" / "week-to-date" without any
special-case logic.

Formulas are recalculated before the file is saved so it opens with real
totals, not formula placeholders that need an app to compute them. The
script prefers a headless LibreOffice round-trip for this and falls back
automatically to writing the already-known values as cached formula
results if LibreOffice isn't available in the running environment -
either way, the formula itself is always the real `SUM()`, never a
hardcoded value.

### Terminology note

HubSpot's underlying property is genuinely called **Deal Owner**
(`hubspot_owner_id`) and the code, API calls and variable names all use
that name throughout, since that's what the field actually is in the CRM.
The generated Excel output is the one place this is renamed: every
user-facing label there says **"BDM"** instead - a display-only
substitution.

## Automation (GitHub Actions)

`.github/workflows/weekly-report.yml` runs the whole pipeline on a
schedule:

- **Schedule**: every Monday, `0 7 * * 1` (07:00 UTC) - plus
  `workflow_dispatch` for an on-demand manual run from the Actions tab.
  GitHub Actions cron is UTC-only and doesn't shift for DST, so 07:00 UTC
  approximates 08:00 Europe/Lisbon during WEST/DST (roughly late
  Mar-late Oct) and lands at 07:00 Lisbon the rest of the year. Adjust the
  cron expression in the workflow file if a different time or day is
  needed.
- **What it does**: installs dependencies, runs `generate_report.py`, and
  commits the regenerated `reports/reports.xlsx` back to the repository
  (skipped if nothing changed that week). No email delivery - check the
  Actions tab (or watch the repository) if you want to know a run
  happened or failed.
- **No LibreOffice install step**: GitHub-hosted `ubuntu-latest` runners
  don't ship LibreOffice Calc by default, and installing it would add
  real minutes to every run. `generate_report.py`'s cached-value fallback
  (see above) handles this automatically - the workflow doesn't need to
  do anything special.

### Required repository secrets

Add these under Settings → Secrets and variables → Actions:

| Secret | Purpose |
|---|---|
| `HUBSPOT_ACCESS_TOKEN` | The same Service Key / private app token used locally |

### Triggering a manual run

Actions tab → "Weekly B2B IR KPI Report" → Run workflow. Useful for
testing the secrets/schedule without waiting for Monday.

## Web dashboard

Alongside the Excel workbook, `generate_report.py` also writes
`dashboard/data/kpi-data.json` - the same computed numbers, reshaped for a
live web dashboard published on Vercel. This is a completely separate,
additive delivery channel: it doesn't touch `reports/reports.xlsx` in any
way.

The dashboard follows the same pattern as the sibling
`gcs-hubspot-funnel-reporting` project: no framework, no server. A small
Python build step inlines the JSON dataset and the GCS logos into three
HTML templates (`dashboard/template_head.html`, `template_body.html`,
`template_js.html`), producing one self-contained `outputs/index.html`.
A second step wraps that file in a client-side, AES-256-GCM password lock
screen (`outputs/vercel/index.html`) - the actual access control, since
Vercel's own deployment protection doesn't cover a bare `*.vercel.app`
alias.

### Building it locally

```bash
python generate_report.py                                    # writes dashboard/data/kpi-data.json
python3 dashboard/build_dashboard.py                          # -> outputs/index.html
python3 dashboard/build_locked.py --password 'your-password'  # -> outputs/vercel/index.html
```

Open `outputs/vercel/index.html` directly in a browser to preview the
lock screen and, once unlocked, the dashboard itself - no server needed.

### Automation and deployment

`.github/workflows/dashboard-data.yml` runs the whole thing daily
(06:00 UTC, plus `workflow_dispatch` for a manual run), commits the
refreshed `dashboard/data/kpi-data.json` and the newly-locked
`outputs/vercel/index.html` back to the repo, and stops there - it does
**not** deploy directly. Instead, connect this repository to a Vercel
project once (Vercel dashboard → Add New Project → import this repo,
framework preset "Other", no build command, output directory
`outputs/vercel`): Vercel's native Git integration then deploys
automatically on every push to `main`, since the push itself is what
lands a new `outputs/vercel/index.html` in the repo.

Required repository secret in addition to the ones above:

| Secret | Purpose |
|---|---|
| `DASHBOARD_PASSWORD` | The password the lock screen decrypts with - share this with whoever needs dashboard access, rotate by changing the secret (next daily run re-locks with it) |

### What's on the dashboard

- **KPI cards** for all 5 stages (Retained/Referred/New Intermediaries/
  Presentations/Calls-Meetings), split by BDM, for the selected period
  (Year to Date or any individual month)
- **Funnel by BDM**: a stacked bar per stage showing each BDM's share
- **Trend**: a monthly line chart across the year, one line per stage
- **Stage performance table**: BDM columns + Grand Total, with
  **Retained Clients** expandable into its Country/Program of Interest
  breakdown - the same dimension as the Excel workbook's row drill-down
- Both annual per-BDM targets (Retained Clients, Calls/Meetings)

Numbers refresh once a day; for the full Month/Week/Day drill-down with
live formulas, the Excel workbook remains the source of record.
