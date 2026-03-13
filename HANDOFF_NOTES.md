# Handoff Notes — March 13, 2026

## What Was Done This Session

### New scripts
- **`FindDeleteEmptyFootnotes.jsx`** — step-through cleanup for unmatched `{{fn:N}}` markers left after `InsertFootnotes.jsx`. Scrolls to each marker in the document view, shows ~60 characters of surrounding context, and asks **Yes** (delete) / **No** (keep and continue). Each deletion is a separate undo step.

### Changes to `InsertFootnotes.jsx`
- After inserting matched footnotes, unmatched markers are now highlighted red automatically using a `Char_UnmatchedMarker` character style (created if absent: CMYK red). The end-of-script report tells you to run `FindDeleteEmptyFootnotes.jsx` to review them.

### Changes to `clean_docx.py`
Three fixes, all in `_mark_superscripts` / `process_run()`:

1. **Repeat-reference ("fake footnote") handling** — When Word uses `customMarkFollows="1"` (author reuses a footnote number by creating an empty footnote body and typing a custom mark like "4"), the script now emits `{{4}}` (plain superscript) instead of `{{fn:5}}` (which would become an unmatched red marker in InDesign because the footnote body is empty). `CleanUp.jsx` then converts `{{4}}` → superscripted `4` via `Char_Superscript`.

2. **Character-style superscript detection** — Added `_get_superscript_style_ids()` which scans `styles.xml` for character styles that define `vertAlign=superscript` (catches Word's built-in "Footnote Reference" style). Runs using that style that weren't caught by `run.font.superscript` are now wrapped as `{{text}}`. This prevents `_strip_character_styles` from silently stripping them back to plain text.

3. **Custom mark text leak fix** — When `customMarkFollows="1"`, the custom mark text ("4") was previously left behind in the run alongside `{{fn:N}}`, producing "4{{fn:5}}" in output. This was fixed in the same session by removing all `<w:t>` elements before appending the new marker text.

### `README.md`
- Added note about repeat-reference handling to the `clean_docx.py` description.
- Corrected `FindDeleteEmptyFootnotes.jsx` button labels from "OK/Cancel" to "Yes/No" (that is how InDesign renders `confirm()` dialogs).

---

## Confirmed Working Workflow

```
Phase 1 — python clean_docx.py YourMonograph.docx
  → YourMonograph_clean.docx
  → YourMonograph_footnotes.txt

Phase 2 — InDesign
  1. Place _clean.docx (Shift+click for autoflow)
  2. TableStyler.jsx
  3. CleanUp.jsx          ← MUST come before InsertFootnotes
  4. InsertFootnotes.jsx  (select _footnotes.txt when prompted)
  5. TitleCaseHeadings.jsx
  6. FindDeleteEmptyFootnotes.jsx  (when ready to clean up red markers)
```

**Critical order constraint:** CleanUp must run **before** InsertFootnotes. If you run CleanUp after InsertFootnotes, its doc-wide `.+` GREP operation (step 0, clearing unnamed char styles) strips the "Footnote Reference" character style from native InDesign footnote reference marks, making them unsuperscripted.

---

## Known Remaining Issues / Next Steps

- **Italic recovery** — journal names and Latin terms still need manual italic re-application after the workflow. This is expected and unchanged.
- **Two-column layout** — Phase 3 (switch Master B to 2 columns, adjust pagination, apply table styles) has not been started for Cry1Ab yet.
- **TitleCaseHeadings `specificFixes` array** — add any new terms discovered during layout review (e.g. protein names, abbreviations) to keep Title Case corrections up to date.
- **Other monographs** — Cry1Ac, EPSPS, PAT/BAR have not been processed yet. The pipeline is ready; just run `clean_docx.py` on each source `.docx` and follow the workflow.
- **footnotes.txt format note** — `export_footnotes_txt` writes only footnotes with non-empty body text. Repeat-reference entries (empty footnote bodies, fn:5 and fn:6 in Cry1Ab) are intentionally excluded. This is correct behavior.

---

## File State at Handoff

| File | Status |
|---|---|
| `clean_docx.py` | Updated (all three fixes above) |
| `InsertFootnotes.jsx` | Updated (unmatched marker highlighting) |
| `FindDeleteEmptyFootnotes.jsx` | New |
| `TableStyler.jsx` | Unchanged |
| `CleanUp.jsx` | Unchanged |
| `TitleCaseHeadings.jsx` | Unchanged |
| `docs/Cry1Ab_FFS_original_clean.docx` | Regenerated with latest fixes |
| `docs/Cry1Ab_FFS_original_footnotes.txt` | 4 entries (fn:1–fn:4); repeat refs excluded |
| `docs/Cry1Ab_FFS_original_footnotes.docx` | Regenerated |
