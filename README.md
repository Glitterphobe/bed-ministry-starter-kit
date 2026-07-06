# Bed Ministry Starter Kit

Everything a church needs to start a bed ministry in its city: the request &
inventory system that has run the San Antonio Child of God Bed Ministry since
2021, volunteer guides for every role, and a step-by-step launch guide.

**Churches: start with the shared Google Drive folder** (link coming with the
first release — until then, open an issue or contact the maintainer). You
copy a template workbook and follow the *New City Launch Guide*; no
programming knowledge is needed.

This repository is the canonical, maintained source: the Apps Script code,
city-neutral documentation, and the operational playbook.

> **Status:** the code (v3.6) is complete and running the San Antonio
> ministry's workflow. The template workbook, launch guide, and city-neutral
> volunteer guides are being packaged now; the operational playbook is being
> captured chapter by chapter.

## What's in the kit

- **The system** — a Google Sheets workbook + Apps Script that handles bed
  requests (Google Form → waiting list), pick lists with substitutions,
  delivery archiving, 39-item inventory reconciliation, and a public impact
  dashboard. One menu click ("🚀 New City Setup") configures a fresh copy.
- **Volunteer guides** (8) — entering requests, packing, receiving donations,
  inventory counts, recording deliveries, fixing mistakes, a complete manual,
  and a ministry-lead guide. *(Being generalized — coming to
  `starter-kit/volunteer-docs/`.)*
- **New City Launch Guide** — copy the workbook to first test delivery in
  about two hours. *(In progress.)*
- **Operational playbook** — bed-building specs, sourcing, costs, volunteers,
  legal, facility, delivery logistics, referral partnerships, fundraising.
  Structure and interview plan in [playbook/OUTLINE.md](playbook/OUTLINE.md);
  content is being captured from the founding ministry.

## What you can customize

- **Ministry Options** (built in): turn off bears, books, plaques, bedrails,
  or whole bed types; set your ministry's name. Start small — flipping an
  option back on later requires no migration.
- **Everything cosmetic**: form wording, book titles, dashboard branding.
- **Not yet customizable**: the 39-item catalog's structure (the inventory
  "slots"). Treat slots as roles — your thin mattress lives in the
  "Mattress 6in" slot whatever its actual size. If your ministry genuinely
  needs different slots, open an issue; a configurable catalog is the
  headline candidate for v4.

## Getting updates

Your copied workbook's script never auto-updates. Check **Bed Ministry →
Setup → About / version** in your sheet against [CHANGELOG.md](CHANGELOG.md);
to update, paste the new [Code.js](Code.js) over your script (Extensions →
Apps Script). All your data lives in the sheet — updates never touch it.

## Support

- Self-service first: **Bed Ministry → Check system health** diagnoses the
  common problems, and the *Fixing a Mistake* guide covers volunteer errors.
- Then: [GitHub issues](../../issues). This project is maintained by a
  volunteer founder — responses are best-effort, there is no SLA.

## Name and license

**Free for churches, ministries, and nonprofits serving their communities.
Not for commercial use or resale.**

- Code: [PolyForm Noncommercial 1.0.0](LICENSE.md) — use, modify, and share
  freely for noncommercial purposes (charitable and religious organizations
  explicitly included). Technically this makes the project "source-available"
  rather than OSI "open source"; for every church that will ever use it, the
  difference is nil.
- Documents (guides, playbook): [CC BY-NC 4.0](https://creativecommons.org/licenses/by-nc/4.0/)
  — copy, adapt, translate, and rebrand with an attribution line ("Adapted
  from the Bed Ministry Starter Kit"); no selling.
- **Your ministry's name is yours to choose.** "Child of God Bed Ministry" is
  the San Antonio ministry's name and is not granted by these licenses.
