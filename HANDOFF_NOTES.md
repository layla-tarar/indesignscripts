# Handoff Notes — March 16, 2026

## What Was Done This Session

### Confirmed Working Workflow
- Verified `clean_docx.py` output for Cry1Ab: `{{fn:1}}` through `{{fn:4}}` in body text matched by 4 entries in `_footnotes.txt`, `{{4}}` appears twice (repeat references), 2 empty Word footnote entries excluded.
- Phase 2 (Phase 1 through TitleCaseHeadings.jsx) confirmed working end-to-end for Cry1Ab.

### SwitchToTwoColumns.jsx — abandoned and deleted
- Three script approaches all crashed InDesign (SIGSEGV in Text engine during reflow).
- A fourth approach (one frame at a time, synchronous reflow) completed but InDesign crashed post-script during final redraw.
- A fifth approach (save after each page, skip already-updated frames) still crashed.
- **Resolution:** Script automation for this step is not viable. Two-column switch is now done manually: `Cmd+A` per spread, `Cmd+B` (Text Frame Options) → 2 columns, 0.1667" gutter. Fast enough in practice.
- `SwitchToTwoColumns.jsx` has been deleted from the repo.

### Documentation updated
- `AFSI_Monograph_InDesign_Template_Guide_v12.md` updated to reflect the current `.docx`-based workflow (replaces old `.txt`-strip approach).
- Document Footnote Options noted as pre-configured in the `.indt` template.
- Script names corrected throughout (`CleanupAfterPlacement.jsx` → `CleanUp.jsx`, `ClearTableOverrides.jsx` → `TableStyler.jsx`).
- Two-column switch updated to reflect the manual approach.

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

### Immediate — automation improvements (to do in next session)
1. **Latin term italics in CleanUp.jsx** — add Find/Change → `Char_Italic` for `in vitro`, `in vivo`, `in situ`, `de novo`, `ex vivo`, `ad libitum` etc. Currently requires manual italic recovery.
2. **Per-monograph term scanner** — extend `clean_docx.py` to emit a `_newterms.txt` report of scientific names, gene names, protein names, event names, and regulatory body abbreviations found in the document that are not in the known lists. Needed before processing Cry1Ac, EPSPS, PAT/BAR.

### Layout & content (Cry1Ab Phase 3 not started)
- **Two-column layout** — manual switch not yet done
- **Table styling** — TStyle_Simple / TStyle_Approvals not yet applied
- **Italic recovery** — journal names and Latin terms need manual italic re-application (Latin terms will be partially automated in next session)
- **TitleCaseHeadings `specificFixes` array** — add new terms found during layout review

### Other monographs
- Cry1Ac, EPSPS, PAT/BAR not yet processed

---

## File State at Handoff

| File | Status |
|---|---|
| `clean_docx.py` | Working correctly |
| `CleanUp.jsx` | Unchanged |
| `InsertFootnotes.jsx` | Unchanged |
| `FindDeleteEmptyFootnotes.jsx` | Unchanged |
| `TableStyler.jsx` | Unchanged |
| `TitleCaseHeadings.jsx` | Unchanged |
| `SwitchToTwoColumns.jsx` | **Deleted** — script approach abandoned |
| `docs/Cry1Ab_FFS_original_clean.docx` | Ready for InDesign placement |
| `docs/Cry1Ab_FFS_original_footnotes.txt` | 4 entries (fn:1–fn:4) |
