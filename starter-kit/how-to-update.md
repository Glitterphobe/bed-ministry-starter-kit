# How to Update Your Copy

*Bed Ministry Starter Kit — your copy never updates itself; here is the five-minute manual update.*

Your workbook's code is frozen at whatever version you copied. Updates are
opt-in, and **an update never touches your data** — every request, delivery,
and inventory count lives in the sheet itself; the code is a separate layer
you can swap.

## 1. Check whether you're behind

- In your workbook: **Bed Ministry → Setup → About / version** — note your
  version (for example, v3.6).
- Compare against **CHANGELOG.md** in the starter-kit repository (the About
  dialog shows the address). If there's a newer version, the changelog says
  what changed and whether any extra steps apply.

## 2. Paste in the new code

1. In the repository, open **Code.js** and copy the entire file
   (on GitHub: the *Raw* button, then Select All → Copy).
2. In your workbook: **Extensions → Apps Script**. You'll see a file
   called **Code.gs** — same file, the editor just uses `.gs`.
3. Click inside it, Select All, Paste (replacing everything), then **Save**
   (Ctrl+S / ⌘S).
4. **If the changelog mentions a manifest change:** in the Apps Script
   editor go to **Project Settings** and tick *Show "appsscript.json"
   manifest file*, then paste the repository's `appsscript.json` over it
   the same way. (Skip this unless the changelog says so.)

## 3. Verify

- Back in the sheet (refresh the tab): **Bed Ministry → Check system
  health**. Run any Setup item it flags — after an update, the changelog
  and the health check together tell you if a one-time setup step is
  needed.
- If Google asks you to authorize again (it does when an update needs a new
  permission), approve it from the ministry account — the *Launch Guide*'s
  Step 2 shows the warning screens, and its troubleshooting section covers
  the "insufficient permissions" case (authorize once from the Apps Script
  editor's **Run** button).

## Rolling back

Pasted an update and something's off? Paste the previous version's
`Code.js` from the repository's release history the same way, and run the
health check. Your data is untouched either way.
