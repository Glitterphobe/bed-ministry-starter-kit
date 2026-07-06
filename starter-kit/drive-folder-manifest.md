# Drive Folder Manifest — "Bed Ministry Starter Kit"

## Public GitHub repo — PUBLISHED 2026-07-05

Canonical source: <https://github.com/Glitterphobe/bed-ministry-starter-kit>
(public, fresh history — initial release commit only; this private repo's
history is never pushed there because old commits contain unaudited docx
screenshots).

Syncing changes from this repo to the public one: copy the changed files over
a clone of the public repo and commit there (the tree mapping is: root code
files + `public-dashboard/` + `playbook/` + `affiliate/` copied as-is;
`starter-kit/public-repo-README.md` → its `README.md`;
`starter-kit/LICENSE.md` → its `LICENSE.md`;
`starter-kit/drive-folder-manifest.md` → its `starter-kit/`).

Still to add there: `starter-kit/volunteer-docs/` and the Launch Guide, once
the Phase 4 screenshot audit clears them.

The church-facing front door. This file records what belongs in the shared
Google Drive folder, its sharing settings, and the IDs the founder must track.
Update it whenever the folder or template changes.

## Contents

| Item | Notes |
|------|-------|
| `START HERE — New City Launch Guide.pdf` | Phase 4 deliverable |
| `Bed Ministry Workbook TEMPLATE` (Google Sheet) | Built from a **blank** spreadsheet (never from the live SA workbook — zero PII incl. version history). Post-setup state: run `setupEverything()` then `setupForm()`; zero data rows; TOTALS row intact; Config placeholders (`you@yourchurch.org`). Its linked Bed Request Form copies along with it. |
| `Volunteer Guides/` (8 PDFs) | City-neutral versions — Phase 4 deliverable (screenshot audit required first) |
| `Data Care one-pager.pdf` | Ships day one — the workbook holds children's data |
| `How to update your copy.pdf` | One page |
| `Operational Playbook/` | Empty placeholder until chapters are captured |

## Sharing settings

- Folder: **viewer** access. Decision pending (founder, Phase 3): publish the
  link openly in the repo README vs. share church-by-church after a
  conversation.
- Template workbook: viewer — churches use **File → Make a copy**.
- Owner: recommended a dedicated Google account (e.g.
  `bedministrystarterkit@gmail.com`), not a personal one, with recovery info
  and a successor recorded.

## IDs to record (fill in at Phase 3)

| What | Value |
|------|-------|
| Drive folder URL | _(pending)_ |
| Template workbook ID | _(pending)_ |
| Template bound-script `scriptId` (for `clasp push` on releases) | _(pending)_ |
| Template form edit URL | _(pending)_ |
| Optional "Register your city" form URL | _(pending)_ |

## Release checklist (every code release)

1. Update `CHANGELOG.md` + `SYSTEM_VERSION` in `Code.js`.
2. `clasp push` to the **template's** scriptId (never San Antonio's).
3. Open the template, run **Check system health**, submit one test form
   response, then delete the test row.
4. If form questions changed: rebuild/adjust the template's linked form and
   re-verify with the health check.
5. Note in the CHANGELOG which version San Antonio is running.
