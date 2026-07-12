# Volunteer-Doc Audit — Phase 4 (2026-07-09)

Result: **PASS.** The eight current guides in `docs/` contain **zero embedded
images** (no screenshots, therefore no PII in images), and their PDFs contain
zero image objects. Document metadata is clean (`python-docx` / `Un-named`
authors). The unaudited-screenshot concern applies only to *old git history*,
which is already excluded from the public repo (fresh-history release).

What was checked, per file: embedded media (`word/media/*` — none anywhere),
`docProps/thumbnail.jpeg` (six files have one; all blank white), full text of
every XML part (document, footers, footnotes), PDF image objects, and
core-property metadata.

Two starter-kit blockers were found — **in text, not images** — and are fixed
in the city-neutral copies in this folder:

1. **Branding** — every guide's subtitle (and two guides' page footers) said
   "Child of God Bed Ministry", a name the licenses do not grant. Now plain
   "Bed Ministry" — usable as printed, and cities may rebrand under CC BY-NC.
2. **Live San Antonio form URL** — *Entering a Bed Request* embedded the live
   SA request form's public URL (visible text + hyperlink). Now a placeholder
   ("ask your ministry lead for your form's link"); the hyperlink target points
   at Google Forms generically.

Kept as-is deliberately: "Tracking Number (Care Portal / APN#)" matches the
FORM_SCHEMA question title verbatim, so guides agree with the form every new
city gets. The v3.6 `setupForm()` ships that exact wording.

Provenance of these copies: byte-identical to `docs/` originals except the
string replacements above (script: session scratchpad `make_neutral.py`);
PDFs re-exported from the edited .docx via Word. These eight .docx + .pdf are
**cleared for the public repo and the Drive folder's `Volunteer Guides/`**.

Maintenance note: `docs/` remains the SA-branded working set. If a guide is
revised there, regenerate its neutral copy (same two replacements) rather than
editing both by hand.
