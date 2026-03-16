# Handoff Notes — March 16, 2026

## What Was Done This Session

### Per-monograph term scanner added to clean_docx.py

`clean_docx.py` now emits a third output: `_newterms.txt`. It scans the original
.docx (before cleaning, so italic markup is still present) and reports candidate
terms not in the known lists, grouped by category:

| Category | Detection method |
|---|---|
| Scientific names | Binomial patterns in italic runs only |
| Gene names | Lowercase token patterns (e.g. `cry1Fa`) in italic runs |
| Protein names | Cry/Vip patterns in full text |
| Event names | Known-prefix patterns (MON, NK, MIR, DAS-XXXXX-X, etc.) |
| Regulatory abbreviations | All-caps 3–6 char tokens, ≥2 occurrences in body text |

Known lists live at the top of the scanner section in `clean_docx.py`. Add
confirmed terms there after reviewing `_newterms.txt`. Terms that need italic
recovery go in `latinTerms` in CleanUp.jsx; title-case fixes go in `specificFixes`
in TitleCaseHeadings.jsx.

### Automation audit completed

Reviewed all existing scripts against the template guide and project instructions.
Key findings:

- **TableStyler.jsx already handles full table styling** — detects approval vs.
  simple tables automatically, applies TStyle_Approvals/TStyle_Simple, styles all
  header cells, Crop/Event Name columns, "x" cells, and highlighted cells. Bottom-row
  style is handled by the table style definition itself. No manual table styling needed.
- **Tables stay in `_clean.docx`** — they are not removed during pre-processing.
  No `[INSERT TABLE X HERE]` markers are needed.
- **Heading style remapping is largely automated** — CleanUp.jsx remaps all standard
  Word heading styles; clean_docx.py heuristically assigns Head_Section, Head_CropName,
  Head_SubsectionUnnumbered, Table_Heading, Table_FootNote from body formatting.
- **Unmatched footnote markers are already flagged** — InsertFootnotes.jsx highlights
  red with Char_UnmatchedMarker; FindDeleteEmptyFootnotes.jsx handles cleanup.

---

## Confirmed Working Workflow

```
Phase 1 — python3 clean_docx.py YourMonograph.docx
  → YourMonograph_clean.docx
  → YourMonograph_footnotes.txt
  → YourMonograph_newterms.txt     ← review before proceeding

Phase 2 — InDesign (single column)
  1. Verify Master B is 1 column
  2. Place _clean.docx (Shift+click for autoflow)
  3. TableStyler.jsx               ← detects & styles all tables automatically
  4. CleanUp.jsx                   ← MUST come before InsertFootnotes
  5. InsertFootnotes.jsx           (select _footnotes.txt when prompted)
  6. TitleCaseHeadings.jsx
  7. Manual italic recovery        (journal names only; Latin terms automated)
  8. FindDeleteEmptyFootnotes.jsx  (when ready to clean up red markers)

Phase 3 — Two-column layout & final polish (NOT YET DONE for Cry1Ab)
  1. Switch to 2 columns: Cmd+A per spread → Cmd+B → 2 cols, 0.1667" gutter
     (manual, spread by spread — script automation crashes InDesign on reflow)
  2. Pagination pass (breaks, widows/orphans)
  3. Final layout polish
```

**Critical order constraint:** CleanUp must run **before** InsertFootnotes.

---

## Known Remaining Issues / Next Steps

### Immediate — automation improvements

1. **Ref_Entry style detection** — reference list paragraphs are not yet
   auto-assigned the `Ref_Entry` paragraph style. Could be done in clean_docx.py
   (heuristic: hanging-indent paragraphs near the end of the document, or paragraphs
   following a "References" heading) or in CleanUp.jsx.

2. **Journal name italic recovery** — the only remaining manual italic task after
   CleanUp.jsx runs. A future pass could scan the original .docx for frequently
   italicized multi-word Title Case phrases not in the Latin terms list and add them
   to a GREP-driven recovery pass in CleanUp.jsx.

### Layout & content (Cry1Ab Phase 3 not started)
- **Two-column layout** — manual switch not yet done
- **Italic recovery** — journal names still need manual italic re-application

### Other monographs
- Cry1Ac, EPSPS, PAT/BAR not yet processed

---

## File State at Handoff

| File | Status |
|---|---|
| `clean_docx.py` | Updated — per-monograph term scanner added (_newterms.txt output) |
| `CleanUp.jsx` | Unchanged — Latin italic recovery still in step 6 |
| `InsertFootnotes.jsx` | Unchanged |
| `FindDeleteEmptyFootnotes.jsx` | Unchanged |
| `TableStyler.jsx` | Unchanged |
| `TitleCaseHeadings.jsx` | Unchanged |
| `docs/Cry1Ab_FFS_original_clean.docx` | Ready for InDesign placement |
| `docs/Cry1Ab_FFS_original_footnotes.txt` | 4 entries (fn:1–fn:4) |
| `docs/Cry1Ab_FFS_original_newterms.txt` | 20 candidates across 5 categories |
