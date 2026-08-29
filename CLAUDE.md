# CLAUDE.md

Guardrails for working in this repo. Full setup/architecture docs live in
[`README.md`](./README.md) - this file is what to check *before* acting,
not a duplicate of it.

**Project**: pulls KPI data from HubSpot and publishes it as a
password-locked web dashboard on Vercel, refreshed daily via GitHub
Actions. See README for the full pipeline
(`How it works: HubSpot → GitHub → Vercel`).

## Rule: check both skills before any `dashboard/` change

Before considering **any** change to `dashboard/` done - the three
`template_*.html` files, `build_dashboard.py`, or `build_locked.py` -
check it against **both**:

- **`.claude/skills/web-design-guidelines`** - interaction/accessibility
  rules (focus states, semantic HTML, `prefers-reduced-motion`, touch
  targets, keyboard support, ARIA).
- **`gcs-design-system`** - brand tokens (colors, fonts, spacing: Night
  Blue `#000957`, Electric Blue `#3F8CFF`, Yrsa/Heebo, the `--chart-1..5`
  categorical palette).

This is a hard requirement, not a nicety: a change that skips this check
is not finished, even if it looks right.

## Things that will bite you if you don't know them

- **Credentials never get hardcoded or committed.** `HUBSPOT_ACCESS_TOKEN`
  comes from `.env` locally (gitignored) or the repo secret of the same
  name in CI; `DASHBOARD_PASSWORD` is a repo secret only. Never print or
  log a token value.
- **Retained Clients links to a deal; Referred Clients links to a
  contact** - not interchangeable, and not just a UI choice. A retained
  client always has a closed deal, so it click-throughs to the **deal**
  record (`meta.deal_url_base`); a referred contact doesn't need one, so
  it click-throughs to the **contact** record (`meta.contact_url_base`).
  Respect this asymmetry if a future stage gets click-through added.
- **Vercel's Output Directory override toggle.** A typed value in
  Project Settings → Output Directory is silently ignored unless its
  override toggle is switched on. This caused a real production 404
  once - see README's pipeline section for the exact settings.
- **`outputs/` is gitignored, but `outputs/vercel/index.html` is
  force-added by CI anyway.** `dashboard-data.yml` runs
  `git add -f outputs/vercel/index.html` deliberately, because landing
  that one file in a commit is what triggers the Vercel deploy. Don't
  "fix" the `.gitignore`, and don't be confused when a gitignored path
  shows up in `git status` after a workflow run.
- **`dashboard/data/kpi-data.json` is generated, not hand-edited**, and
  the daily workflow auto-commits it - a manual push can race it. Resolve
  with `git fetch` + `git merge` (fast-forward when possible); use
  `git checkout --ours dashboard/data/kpi-data.json` only when the local
  commit has a schema change the remote's older auto-commit doesn't have.
- **The Excel workbook is dormant, not deleted.** `build_workbook()` and
  its styling helpers still exist in `generate_report.py` but aren't
  called from `main()` any more (see README's "Excel workbook (dormant)"
  section). Don't resurrect it into the pipeline without confirming
  that's actually wanted.
- **`reports/reports.xlsx` and `outputs/vercel/index.html` are git-tracked
  build artifacts, not source.** Never hand-edit them, and never
  `rm -rf` a directory that might contain one without checking
  `git status` first - this exact mistake has happened before in this
  repo and had to be reverted with `git checkout --`.
- **If replicating this pattern in another project**, confirm which
  branch is actually set as that repo's GitHub default branch - Vercel's
  Git integration deploys from whichever branch is configured as
  production there.
