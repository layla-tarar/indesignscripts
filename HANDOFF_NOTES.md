# Handoff Notes — March 16, 2026

## What Was Done This Session

### Latin term italic recovery added to CleanUp.jsx (step 6)

After the override-clearing pass strips all Word character formatting, a new step
re-applies italic to standard Latin/scientific terms. Four weight-context groups:

| Character style | Paragraph styles |
|---|---|
| `Char_Italic` | Body_Text, Body_BulletL1, Body_Footnote, Table_Heading |
| `Char_Regular` | Table_FootNote (fully italic base — terms go roman to contrast) |
| `Char_SemiboldItalic` | Head_SubsectionNumbered, Head_SubsectionUnnumbered, Head_SubSubSectionUnnumbered |
| `Char_BoldItalic` | Head_Section |

Terms covered: in vitro, in vivo, in situ, de novo, ex vivo, ad libitum, in planta,
in silico, sensu stricto, sensu lato, et al., per se, in utero, in ovo.

To add more terms: edit the `latinTerms` array in CleanUp.jsx step 6.
Journal name italics still require manual recovery.

---

## Confirmed Working Workflow

```
Phase 1 — python3 clean_docx.py YourMonograph.docx
  → YourMonograph_clean.docx
  → YourMonograph_footnotes.txt

Phase 2 — InDesign (single column)
  1. Verify Master B is 1 column
  2. Place _clean.docx (Shift+click for autoflow)
  3. TableStyler.jsx
  4. CleanUp.jsx          ← MUST come before InsertFootnotes
  5. InsertFootnotes.jsx  (select _footnotes.txt when prompted)
  6. TitleCaseHeadings.jsx
  7. FindDeleteEmptyFootnotes.jsx  (when ready to clean up red markers)

Phase 3 — Two-column layout & tables (NOT YET DONE for Cry1Ab)
  1. Switch to 2 columns: Cmd+A per spread → Cmd+B → 2 cols, 0.1667" gutter
  2. Pagination pass (breaks, widows/orphans)
  3. Apply table styles (TStyle_Simple / TStyle_Approvals + cell/edge styles)
  4. Final layout polish
```

**Critical order constraint:** CleanUp must run **before** InsertFootnotes.

---

## Known Remaining Issues / Next Steps

### Immediate — automation improvements
1. **Per-monograph term scanner** — extend `clean_docx.py` to emit a `_newterms.txt`
   report of scientific names, gene names, protein names, event names, and regulatory
   body abbreviations found in the document that are not in the known lists.
   Needed before processing Cry1Ac, EPSPS, PAT/BAR.

### Layout & content (Cry1Ab Phase 3 not started)
- **Two-column layout** — manual switch not yet done
- **Table styling** — TStyle_Simple / TStyle_Approvals not yet applied
- **Italic recovery** — journal names still need manual italic re-application
- **TitleCaseHeadings `specificFixes` array** — add new terms found during layout review

### Other monographs
- Cry1Ac, EPSPS, PAT/BAR not yet processed

---

## File State at Handoff

| File | Status |
|---|---|
| `clean_docx.py` | Working correctly |
| `CleanUp.jsx` | Updated — Latin italic recovery added (step 6) |
| `InsertFootnotes.jsx` | Unchanged |
| `FindDeleteEmptyFootnotes.jsx` | Unchanged |
| `TableStyler.jsx` | Unchanged |
| `TitleCaseHeadings.jsx` | Unchanged |
| `docs/Cry1Ab_FFS_original_clean.docx` | Ready for InDesign placement |
| `docs/Cry1Ab_FFS_original_footnotes.txt` | 4 entries (fn:1–fn:4) |
