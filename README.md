# AFSI Monograph InDesign Prep Scripts

Scripts for preparing AFSI protein monograph Word drafts for placement into the InDesign template. Covers the Cry1Ab, Cry1Ac, EPSPS, and PAT/BAR food & feed safety series (and potentially the environmental safety series).

For full style specs, GREP patterns, and template architecture, see [docs/AFSI_Monograph_InDesign_Template_Guide_v12.md](docs/AFSI_Monograph_InDesign_Template_Guide_v12.md).

---

## Scripts

### `clean_docx.py` — Word pre-processing (run before InDesign)

**Requires:** Python 3, `python-docx` (`pip install python-docx`)

```
python clean_docx.py YourMonograph.docx
```

Produces:
- `YourMonograph_clean.docx` — place this in InDesign (tables intact, text cleaned)
- `YourMonograph_footnotes.txt` — footnote key/text pairs for InsertFootnotes.jsx

What it does automatically:
- Marks superscript runs as `{{text}}` and native footnote refs as `{{fn:N}}`
- Handles repeat references: when the author reuses a footnote number (Word's `customMarkFollows` mechanism with an empty footnote body), the display mark (e.g. `4`) is preserved as a plain superscript `{{4}}` rather than generating a `{{fn:N}}` marker that would be unmatched
- Strips Word field codes and character-level overrides (preserves italics)
- Cleans bullet characters (U+2022) and tilde operators (U+223C)
- Infers InDesign paragraph styles from formatting heuristics (all-caps headings, bold paragraphs, table captions, table footnote markers)
- Extracts merged description rows from table tops into `Table_Header` paragraphs

---

### InDesign Scripts (`.jsx`)

Install by copying all `.jsx` files to your InDesign Scripts Panel folder (right-click the "User" folder in the Scripts panel → "Reveal in Finder").

**Run in this order after placing `_clean.docx`:**

#### 1. `TableStyler.jsx`
- Clears all table, cell, and paragraph style overrides imported from Word
- Converts the first row of each table to a header row
- Sets row heights to "At Least" 3pt
- Applies `Table_Span` to each table's container paragraph
- ⚠️ Must run **before** `CleanUp.jsx` (which re-applies `Char_Superscript` to cells)

#### 2. `CleanUp.jsx`
- Removes unnamed/blank imported character styles and replaces usages with `[None]`
- Remaps imported Word paragraph styles → InDesign template styles (e.g., Normal → Body_Text, Heading 1 → Head_Section)
- Clears local formatting overrides (paragraph and character level)
- Fixes double spaces, extra paragraph returns, and multiplication signs (`x` → `×`)
- Converts `{{N}}` and `{{letter}}` markers → strips braces, applies `Char_Superscript`
- Recovers italics on Latin/scientific terms (in vitro, in vivo, in situ, de novo, ex vivo, ad libitum, etc.) stripped by the override-clearing pass — applies the correct character style per paragraph weight context: `Char_Italic` (prose), `Char_Regular` (Table_FootNote, which is fully italic), `Char_SemiboldItalic` (subheadings), `Char_BoldItalic` (Head_Section). To add terms, edit the `latinTerms` array in the script.
- ⚠️ Must run **before** `InsertFootnotes.jsx` — native InDesign footnotes created by InsertFootnotes cause InDesign's Text engine to crash when CleanUp runs doc-wide text operations

#### 3. `InsertFootnotes.jsx`
- Reads `_footnotes.txt` (select it when prompted)
- Replaces every `{{fn:N}}` marker with a native InDesign footnote populated with the correct text
- Processes markers in reverse document order to preserve text positions
- Unmatched markers (no entry in the txt file) are highlighted red using the `Char_UnmatchedMarker` character style (created automatically if absent)
- Run `FindDeleteEmptyFootnotes.jsx` after layout review to clean them up

#### 3b. `FindDeleteEmptyFootnotes.jsx` *(run after InsertFootnotes, when ready)*
- Finds all remaining `{{fn:N}}` markers in the document
- Scrolls to each one and shows surrounding context in a confirmation dialog
- **Yes** = delete the marker; **No** = keep it and move to the next
- Each deletion is a separate undo step

#### 4. `TitleCaseHeadings.jsx`
- Applies Title Case to all heading paragraph styles
- Lowercases articles, prepositions, and conjunctions (APA 7th ed.) unless first word
- Fixes specific terms Title Case breaks (e.g., `Cry1ab` → `Cry1Ab`, `Ge` → `GE`)
- To add terms, edit the `specificFixes` array in the script

---

## Workflow Summary

```
Phase 1 — Word Prep
  └─ python clean_docx.py YourMonograph.docx
       → YourMonograph_clean.docx
       → YourMonograph_footnotes.txt

Phase 2 — InDesign (single column)
  ├─ Place YourMonograph_clean.docx (Shift+click for autoflow)
  ├─ Run: 1. TableStyler.jsx
  ├─ Run: 2. CleanUp.jsx
  ├─ Run: 3. InsertFootnotes.jsx  (select _footnotes.txt when prompted)
  ├─ Run: 4. TitleCaseHeadings.jsx
  ├─ Manual italic recovery (journal names; Latin terms now automated via CleanUp.jsx)
  └─ Run: FindDeleteEmptyFootnotes.jsx  (when ready to clean up red {{fn:N}} markers)

Phase 3 — Two-column layout & final polish
  ├─ Switch Master B to 2 columns
  ├─ Adjust pagination (page/column breaks, widows/orphans)
  ├─ Apply table styles (TStyle_Simple / TStyle_Approvals + cell/edge styles)
  └─ Final layout polish
```

---

## Project Files

```
indesignscripts/
├── clean_docx.py                    # Python pre-processing script
├── CleanUp.jsx                      # InDesign cleanup (step 2)
├── FindDeleteEmptyFootnotes.jsx     # InDesign unmatched marker cleanup (step 3b)
├── InsertFootnotes.jsx              # InDesign footnote insertion (step 3)
├── TableStyler.jsx                  # InDesign table styling (step 1)
├── TitleCaseHeadings.jsx            # InDesign title case (step 4)
├── deprecated_scripts/            # Old scripts kept for reference
└── docs/
    ├── AFSI_Monograph_InDesign_Template_Guide_v12.md  # Full template guide
    ├── AFSI_Monograph_Prep_Project_Instructions_Original.md
    ├── Cry1Ab_FFS_original.docx   # Source monograph draft
    ├── Cry1Ab_FFS_original_clean.docx
    ├── Cry1Ab_FFS_original_footnotes.docx
    ├── Cry1Ab_FFS_original_footnotes.txt
    └── ProteinMonograph_Template_original.indt
```
