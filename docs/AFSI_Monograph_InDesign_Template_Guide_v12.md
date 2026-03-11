# AFSI Protein Monograph — InDesign Template Framework & Guide (v12)

*Updated February 2026 based on analysis of the Cry1Ab FFS draft (Cry1Ab_FFS_27JAN2026_18FEB2026.docx)*

---

## 1. Project Overview

**Goal:** Create a reusable InDesign template (`.indt`) for the updated AFSI protein monograph series that mirrors the structure of the original publications while reflecting current AFSI branding. The template should automate as much formatting as possible through paragraph styles, character styles, GREP styles, and master pages so that flowing in updated Word content requires minimal manual intervention.

**Series scope:** At minimum, four food & feed safety monographs (Cry1Ab, Cry1Ac, EPSPS, PAT/BAR) plus potentially the environmental safety series. All share common structure, so one template should serve the entire set.

---
## 2. Post-Placement Workflow

This is a three-phase workflow. Phase 1 preps the text in Word. Phase 2 places clean text into InDesign and applies all styles in a single-column layout. Phase 3 pastes tables back in, handles footnotes, and switches to two columns for final layout.

---

### Phase 1 — Prepare the text in Word

The Word document contains hidden XML formatting that causes InDesign to crash during recomposition. Instead of fighting overrides, strip all formatting and place clean text.

**Step 1 — Mark superscript text before stripping.**

Your colleagues use manually superscripted numbers (not Word's native footnote system) for shared footnotes, so these will be lost when you strip formatting. Mark them now so you can find them later.

In Word, open Find and Replace (Ctrl+H / Cmd+H):

1. Click in the **Find what** field. Leave it empty. Click **More >>**, then **Format > Font**, check **Superscript**, click OK.
2. Click in the **Replace with** field. Type: `{{^&}}` — this means "wrap the found text in double curly braces." Click **Format > Font**, uncheck **Superscript**, click OK.
3. Click **Replace All**.

Every superscripted character is now wrapped in visible markers: `{{1}}`, `{{2}}`, `{{a}}`, etc. These will survive the plain text export.

**Step 2 — Delete tables, leave markers.**

For each table in the document:

1. Place your cursor on the line immediately above the table (the table heading or caption line).
2. Type a marker on a new line below the caption: `[INSERT TABLE X HERE]` (where X is the table number).
3. Select the entire table and delete it.

Leave all table headings, captions, caption sub-notes, and table footnotes in place — only delete the table grid itself. You will paste the tables back from the original Word file in Phase 3.

**Step 3 — Save a working copy as Plain Text.**

File > Save As > Plain Text (.txt). Choose **UTF-8** encoding. This strips all formatting, styles, embedded XML, and hidden junk. The result is pure text with your `{{}}` superscript markers and `[INSERT TABLE X HERE]` placeholders intact.

Keep the original .docx open or accessible — you will copy tables from it in Phase 3.

---

### Phase 2 — First pass in InDesign (single column, text only)

Work in 1 column for this entire phase. Tables are not present yet.

**Prerequisite:** Verify that InDesign Preferences > Type > Use Typographer's Quotes is enabled before placing.

**Step 1 — Switch Master B to single column.**

Before placing, change your master page layout to single column: open the **Pages** panel, double-click **Master B** to edit it, select the text frame, then Object > Text Frame Options > Columns: **1**. This ensures all auto-created pages during placement use a single column. You will switch Master B back to 2 columns in Phase 3.

**Step 2 — Verify Document Footnote Options** (Type > Document Footnote Options):

- Numbering and Formatting tab:
  - Footnote Reference Number in Text — Position: **Apply Superscript**, Character Style: **Superscript**
  - Footnote Formatting — Paragraph Style: **Body_Footnote**, Separator: **^m** (en space)
- Layout tab:
  - Minimum Space Before First Footnote: **3pt**
  - Space Between Footnotes: **3pt**
  - Allow Split Footnotes: **checked**
  - Rule Above — First Footnote in Column: **On**, Weight: **0.25pt**, Color: **AFSI Light Gray (#91908F)**, Left Indent: **0**, Width: **100pt**, Offset: **3pt**

**Step 3 — Place the .txt file.**

File > Place, select the .txt file. Your cursor becomes a loaded text icon. **Hold Shift and click** on the first page's text frame — this is autoflow. InDesign will automatically create new pages (based on your master page) and add threaded text frames until all the text is placed. No overset, no manual page creation. The text comes in completely clean — no overrides, no hidden formatting, no XML junk.

**Step 4 — Apply Body_text to everything.**

Select All (Cmd/Ctrl+A), click **Body_text** in the Paragraph Styles panel. This gives the entire document the correct base font, size, spacing, and activates all your GREP styles (scientific names, gene names, units, event names, cross-references).

**Step 5 — Apply heading paragraph styles.**

Walk through the document and apply the correct paragraph style to each heading. Since the text is clean, regular clicks work — no Alt/Opt-click needed. Refer to the original Word document for structure:

- [ ] **Head_Section** — major numbered sections (INTRODUCTION, ORIGIN AND FUNCTION OF CRY1AB, etc.)
- [ ] **Head_SubsectionNumbered** — named sub-sections (Mechanism of Cry1Ab insecticidal activity, etc.)
- [ ] **Head_SubSubSectionUnnumbered** — unnumbered sub-sections within a numbered subsection
- [ ] **Head_SubsectionUnnumbered** — sub-sub-headings (Acute toxicity studies on Cry1Ab protein, etc.)
- [ ] **Head_CropName** — crop names (Maize, Cotton, Cowpea, Eucalyptus, Sugarcane)
- [ ] **Table_Heading** — table titles placed above tables by colleagues
- [ ] **Table_Caption** — table captions (Table 1. Regulatory approvals for...)
- [ ] **Table_CaptionSub** — qualifying notes below captions (Approvals after 2020 are highlighted.)
- [ ] **Body_BulletL1** / **Body_BulletL2** — bulleted lists
- [ ] **Table_FootNote** — table-specific footnotes with superscript letters
- [ ] **Ref_Entry** — reference list entries (when present)

**Step 6 — Find/Change operations.**

**Part A — Create native footnotes first (manual):**

The stripped text places all footnote content at the end of the document (since footnotes are just regular text after stripping). Before running the cleanup script, create native footnotes for first occurrences:

For each footnote number (1, 2, 3, 4, etc.):
1. Find the footnote text at the end of the document (e.g., `{{1}} GE crops are crops that have been modified...`). Copy the text (without the `{{1}}` prefix).
2. Use Edit > Find/Change (GREP tab). Find: `\{\{1\}\}` (use the specific number, not `\d+`). Click **Find Next** to jump to the **first** occurrence in the body text.
3. Delete `{{1}}` at that location, then Type > Insert Footnote. Paste the footnote text into the footnote area.
4. Repeat for {{2}}, {{3}}, {{4}}, etc.
5. Delete the block of footnote text at the end of the document once all native footnotes are created.
6. **Footnotes with no content** (the Cry1Ab draft has empty footnotes [5] and [6]): Flag these and confirm with colleagues whether they need content or removal.

**Part B — Run the CleanupAfterPlacement.jsx script:**

File > Scripts > Scripts Panel, double-click **CleanupAfterPlacement.jsx**. The script runs all Find/Change operations in sequence:

1. Double spaces → single space (` {2,}`, not `\s{2,}`)
2. Extra paragraph returns → single return (`\r{2,}` → `\r`)
3. Strip bullet characters (`• ` at start of paragraphs)
4. Tilde operator → standard tilde (`\x{223C}` → `~`)
5. Multiplication signs in stacked event names (`x` → `×`)
6. "Table N." → "Table N:" (scoped to Table_Heading paragraphs)
7. Superscript remaining `{{}}` markers → strips braces, applies Char_Superscript

The script reports how many changes it made for each operation. If you accidentally run Part B before Part A, undo (Cmd+Z) to revert the superscript step, create the native footnotes, then run the script again.

To install: copy CleanupAfterPlacement.jsx to your InDesign Scripts Panel folder (same location as TitleCaseHeadings.jsx).
- [ ] **Title case headings.** Run the **TitleCaseHeadings.jsx** script: File > Scripts > Scripts Panel, navigate to the script, and double-click. The script does three things automatically: (1) applies Title Case to all paragraphs with heading styles, (2) lowercases articles, prepositions, and conjunctions that shouldn't be capitalized (a, an, and, as, at, but, by, for, from, in, into, nor, of, on, or, so, the, to, up, via, with, yet) unless they're the first word, and (3) fixes specific terms like "Ge" → "GE". To install the script, copy TitleCaseHeadings.jsx to your InDesign Scripts Panel folder (find it by right-clicking the "User" folder in the Scripts panel and choosing "Reveal in Finder"). To add more term fixes, edit the `specificFixes` array in the script.
- [ ] **Review title case results.** The script handles most cases, but do a quick scan for: scientific terms that Title Case may have broken (e.g., "Cry1ab" should be "Cry1Ab"), other abbreviations that need ALL CAPS, and any edge cases the script missed.

**Step 7 — Verify GREP automation:**

- [ ] **Scientific names italicized?** Spot-check: Bacillus thuringiensis, B. thuringiensis, Zea mays, Z. mays, Vigna unguiculata, Gossypium hirsutum should all be italic (both full and abbreviated forms).
- [ ] **Gene names italicized, protein names roman?** Spot-check: cry1Ab, cry2Ab, vip3Aa, epsps should be italic. Cry1Ab, Cry2Ab should be roman (upright).
- [ ] **Units staying with numbers?** Spot-check: "4000 mg/kg body weight" and "2.5 ug/g fresh weight" should not break across lines.
- [ ] **Event names not hyphenating?** Spot-check: MON810, Bt11, AAT709A should not split at line breaks.
- [ ] **Cross-references intact?** Spot-check: "Table 1", "Table 8" should not split across lines.

**Step 8 — Manual italic recovery.**

The plain text export strips all italic formatting. Your GREP styles recover most of it (scientific names, gene names), but some italics must be restored manually with the Word document open side-by-side:

- [ ] Journal names in references (GREP can't reliably catch these)
- [ ] Occasional emphasis in body text (words italicized for stress, not convention)
- [ ] Latin terms not in the GREP species list (e.g., in vitro, in vivo, de novo — add these to GREP if they appear frequently)
- [ ] Any other formatting your GREP patterns don't cover

---

### Phase 3 — Two-column layout, tables, and final polish

Tables are the primary crash culprit in InDesign, so this phase switches to two columns first (while the document is still table-free and stable), then builds and formats tables in isolation before inserting them into the flow.

**Step 1 — Switch to two columns.**

In the **Pages** panel, double-click **Master B** to edit it. Select the text frame, then Object > Text Frame Options > Columns: **2**, Gutter: **0.1667"** (1 pica). Return to the document pages — all pages based on Master B will reflow into the two-column layout.

**Step 2 — Adjust pagination.**

With the text now in two columns, do an initial layout pass:

- [ ] **Page breaks and column breaks.** Adjust where sections start, ensure headings don't strand at the bottom of columns (Keep with Next should handle most of this).
- [ ] **Widow/orphan check.** Scan for single lines at the top or bottom of columns. Use Char_No Break on the last 2–3 words of a paragraph to pull a widow back, or insert a discretionary line break.
- [ ] Leave the `[INSERT TABLE X HERE]` markers in place for now — you'll replace them in Step 5.

**Step 3 — Build and format tables in separate text frames.**

**Prerequisite:** Verify that InDesign > Preferences > Clipboard Handling > When Pasting Text and Tables from Other Applications is set to **All Information (Index Markers, Swatches, Styles, etc.)**. If this is set to "Text Only," tables will paste as tab-delimited plain text instead of actual table structures.

Create a new text frame at the end of the document (or on a blank page) as a staging area. For each table:

1. In Word, select the table (click the table move handle at the top-left corner).
2. Copy (Cmd/Ctrl+C).
3. Paste into the staging text frame in InDesign (Cmd/Ctrl+V).

Paste all tables into this staging area before running the scripts.

**Step 4 — Run scripts and apply table styles.**

1. Run **ClearTableOverrides.jsx** (File > Scripts > Scripts Panel). The script processes every table automatically:
   - Clears all table, cell, and paragraph style overrides
   - Converts the first row of each table to a header row
   - Sets all row heights to "At Least" 3pt so rows expand to fit content
   - Applies Table_Span to each table's paragraph

2. Run **CleanupAfterPlacement.jsx** to clean up any double spaces, extra returns, tilde operators, or other junk within cell text.

3. Manually apply styles to each table:
   - Apply the table style: **TStyle_Simple** or **TStyle_Approvals** (or variant).
   - If cell styles don't take, select header cells and Alt/Opt-click the header cell style, then select body cells and Alt/Opt-click the body cell style.
   - Apply edge styles: **CStyle_Header_left_leftalign** on leftmost header cell, **CStyle_Header_right_leftalign** on rightmost (or centered variants for approval tables).
   - Apply bottom row style: Select entire bottom body row, apply **CStyle_BodyBottom** (or **CStyle_BodyApprovalBottom**).
   - Adjust column widths as needed.

**Step 5 — Insert formatted tables into the document flow.**

For each `[INSERT TABLE X HERE]` marker in the main text:

1. In the staging area, click inside the formatted table and select the entire table (Edit > Select All or Cmd/Ctrl+A while cursor is in the table).
2. Cut (Cmd/Ctrl+X).
3. In the main text flow, select the `[INSERT TABLE X HERE]` text and paste (Cmd/Ctrl+V). The formatted table replaces the marker with all styles intact.

**Fallback — if pasting a table into the flow causes crashing:**

Don't paste it inline. Instead:

1. In the main text flow, delete the `[INSERT TABLE X HERE]` marker.
2. Manually adjust the text frame and text to leave enough blank space on the page for the table.
3. Leave the formatted table in its own separate text frame and position it on the page over the blank space.
4. This means the table is not threaded into the text flow — it won't move if text reflows. You'll need to reposition it manually if earlier content changes.

**Step 6 — Final layout polish:**

- [ ] **Table placement.** Confirm each table fits its column or spans correctly. Resize columns within tables as needed.
- [ ] **Table numbering sequential?** Verify Tables 1 through N are in order and body text cross-references match.
- [ ] **Apply post-2020 highlight styling** where noted (CStyle_Highlighted on individual cells in approval tables).
- [ ] **Final GREP spot-check.** Confirm all automation is still working after the two-column reflow and table insertion.

---


## 3. Analysis of the Updated Monograph Draft (Cry1Ab)

### 3.1 Document Structure — Updated

The updated Cry1Ab draft is significantly longer and more complex than the original. The structure is now:

- **Title page (to be designed):** Large title, organization name and address, date, keywords sidebar, AFSI branding elements.
- **Introduction (2–3 pages):** Multi-paragraph introduction including the Codex framework factors (bulleted list), followed by a substantial narrative on global food safety approvals with country-by-country detail and statistics.
- **Approval tables (5–8 pages):** Five separate regulatory approval tables (Tables 1–5), each covering a different crop type or time period, plus a regulatory document references table (Table 6). Tables vary from ~12 to ~21 country columns.
- **Origin and Function of Cry1Ab (3–4 pages):** Includes subsections on Bt biology, mechanism of action, and modifications to the cry1Ab gene. Contains Table 7 (genetic elements/expression cassettes — a text-heavy descriptive table).
- **Expression of Cry1Ab in GE Plants (2–3 pages):** Discussion of expression levels and exposure assessment. Contains Table 8 (expression data with superscript table footnotes).
- **Food and Feed Safety (~4 pages):** Includes subsections on toxicological studies, allergenicity (with sub-sub-sections on bioinformatics and digestibility), compositional analysis (with crop-specific sub-sections: Maize, Cotton, Cowpea, Eucalyptus, Sugarcane), feeding studies, and safety assessment of stacked events.
- **Conclusion (1 page):** Summary of safety findings.
- **Footnotes:** Six document-level footnotes (some currently empty/placeholder).
- **References (not yet included in draft):** Expected 8–15 pages based on original monograph.

### 3.2 Key Structural Differences from the Original

| Feature | Original Monograph | Updated Draft | Template Implication |
|---------|-------------------|---------------|---------------------|
| **Approval tables** | 1 massive table (~20 countries × ~42 events) | 5 separate tables, varying widths (12–21 country columns) | Need flexible table sizing, not just one full-page landscape approach |
| **Total tables** | 2 (approvals + expression) | 8 (Tables 1–8) with 3 distinct table types | Need 3+ table styles, not 2 |
| **Heading hierarchy** | Clean H1 → H2 | Mixed: H1, H2, unnumbered bold subheadings, all-caps crop names, run-in headings | Need additional heading styles (Head_CropName, Head_RunIn, Head_SubsectionUnnumbered) |
| **Bulleted lists** | Minimal | 10-item formal list in Introduction (Codex factors) | Body_BulletL1 style needs full-sentence treatment |
| **Footnotes** | 1 | 6 (some empty/placeholder), plus table-specific footnotes with superscript letters | Need Table_FootNote style; QC check for empty footnotes |
| **Introduction length** | ~1 page | 2–3 pages of narrative before first table | Longer body text run before first table page |
| **Crop sub-sections** | None | 5 crop-specific sub-sections under Compositional Analysis (all-caps headings in Word) | Head_CropName paragraph style; change to title case |
| **Scientific names** | ~3 species | 8+ species across 6 crops | Expanded GREP auto-italicize list |
| **Gene names** | ~3 genes | 8+ genes (cry1Ab, cry2Ab, cry1Fa, cry1Bb, cry2Aa, vip3Aa, cp4-epsps, epsps) | Expanded GREP gene italicize pattern |
| **Event names** | Simple (MON810, Bt11) | Complex stacked names with × signs (e.g., MON87427 × MON89034 × MON810 × MIR162 × MON87411 × MON87419) | GREP for normalizing ×/x and Char_No Break on full event names |
| **Table footnotes** | None | Superscript letters (a, b) below Table 8 | New Table_FootNote paragraph style |

### 3.3 Formatting Issues Identified

All issues from the original analysis remain relevant, plus these new observations:

| Issue | Location | Impact |
|-------|----------|--------|
| **Inconsistent multiplication signs** | Stacked event names throughout — mix of "x" and "×" | Needs normalization via Find/Change |
| **All-caps headings not mapped to styles** | Compositional Analysis sub-sections (MAIZE, COTTON, etc.) | Will import as Body_text; needs post-import remap |
| **Empty footnotes** | Footnotes 5 and 6 are placeholder/empty | QC flag needed in checklist |
| **Bold-formatted "x" markers** | All approval tables use bold "x" for approvals | Need consistent styling — consider AFSI Green or bold center-aligned |
| **Inconsistent heading capitalization** | Mix of Title Case, ALL CAPS, and Sentence case across section headings | GREP or manual cleanup needed post-import |
| **Table caption formatting** | Some captions run to 3+ lines with qualifying notes | Table_Caption style needs to accommodate multi-line captions with sub-text |
| **"Highlighted" approvals** | Tables 2–4 note "Approvals after 2020 are highlighted" | Need a cell style or shading approach for post-2020 highlight |
| **Superscript in table cells** | Table 8 has superscript "a" and "b" within data cells | Character style needed for in-table superscripts |
| **Parenthetical citation style** | Body uses (Author Year) not [1] numbered refs | GREP for orphaned reference numbers less critical; parenthetical refs need no-break treatment instead |

---

## 4. Brand Specifications

### 4.1 Colors (from AFSI_Colors.pdf)

| Name | HEX | RGB | CMYK | Suggested Use |
|------|-----|-----|------|---------------|
| **Green** | #43BEA2 | 67, 190, 162 | 67, 0, 47, 0 | Primary accent — headings, title bar, sidebar backgrounds, table headers |
| **Blue** | #4397D2 | 67, 151, 210 | 70, 29, 0, 0 | Secondary accent — hyperlinks, table header highlights |
| **Gray** | #636466 | 99, 100, 102 | 0, 0, 0, 75 | Body text alternative, secondary headings |
| **Taupe** | #D9D0C7 | 217, 208, 199 | 14, 15, 19, 0 | Sidebar backgrounds, table alternating rows |
| **Light Gray** | #91908F | 145, 144, 143 | 5, 5, 5, 48 | Captions, footnotes, table borders |
| **Brown** | #997655 | 153, 118, 85 | 36, 50, 70, 13 | Tertiary accent (sparingly) |
| **Mustard** | #E5B351 | 229, 179, 81 | 10, 30, 80, 0 | Callout boxes, warnings |
| **Light Green** | #A3BB3F | 163, 187, 63 | 42, 11, 96, 0 | Environmental safety series accent |
| **Dark Green** | #6A8D65 | 106, 141, 101 | 62, 28, 70, 8 | Environmental safety series accent |
| **Light Blue** | #76BEEA | 118, 190, 234 | 50, 10, 0, 0 | Informational callouts |
| **Dark Blue** | #4A679E | 74, 103, 158 | 80, 60, 13, 0 | Formal/institutional elements |

### 4.2 Typography

- **Brand font family:** Source Sans Pro (Regular, Semibold, Bold, Italic, etc.)
- **Recommendation:** Use Source Sans Pro for all text. If a monospaced or serif complement is needed (e.g., for gene names in tables), consider Source Serif Pro for consistency within the Source family, or use italic Source Sans Pro.

### 4.3 Current Website Context

AFSI's current site (foodsystems.org) uses a clean, modern aesthetic with the green (#43BEA2) as the dominant brand color, white space, and sans-serif typography. The updated monographs should feel consistent with this web presence — professional, clean, and science-forward.

---

## 5. Template Architecture

### 5.1 Document Setup

| Setting | Value | Notes |
|---------|-------|-------|
| **Page size** | US Letter (8.5" × 11") | Matches original; standard for US-based non-profit |
| **Margins** | 0.75" all sides (symmetric) | Generous margins give the text block a compact, journal-like framing. Scientific publications typically use 0.75"–1" margins — the density and rigor come from type size, leading, and column structure, not from pushing text to the page edges. Produces a 7.0" wide text area, yielding two columns at ~3.417" each with the 0.1667" gutter — a comfortable width for 9.5pt justified text. |
| **Columns** | 2 (default); override to 1 on table pages as needed | |
| **Bleed** | 0.125" all sides | For any edge-to-edge color blocks |
| **Slug** | 0.25" | For print instructions if needed |
| **Facing pages** | Off (or On if series will be bound) | Original appears to be single-sided |

### 5.2 Master Pages

#### Master A — Title Page
- Full-width layout (single column)
- Top area: Large title block with AFSI Green (#43BEA2) accent bar or background element
- Right sidebar area (approximately 2.5" wide) for keywords box on Taupe (#D9D0C7) background
- Main body area for introduction text
- Bottom area for footnotes
- AFSI logo placement (top or bottom, per current brand guidelines)
- Organization name, address, and date fields

#### Master B — Body (Two-Column)
- Two equal columns with 0.1667" gutter (1 pica — InDesign default; maximizes column width for better justified text)
- Running header with protein name and/or section title (Source Sans Pro Light, 8pt, AFSI Gray)
- Running footer with page number (Source Sans Pro Regular, 9pt, centered or outside-aligned)
- Optional colored rule under header
- **Used for all content pages:** body text, tables, references, everything after the title page

**Handling full-width content on Master B pages:**
When a page needs full-width layout for a wide table or a placed object, simply override the master text frame on that page and change Text Frame Options → Columns to 1. Or, for tables that are placed as standalone objects, ignore the column guides and position the table frame across the full text area. The header/footer treatment remains consistent throughout.

**Handling reference pages:**
The references section uses the same Master B layout. The smaller text size and tighter spacing are entirely controlled by the Ref_Entry paragraph style — no separate master needed.

#### Page Sequence (typical for Cry1Ab)

Based on the actual draft content, expect approximately this flow:

```
Master A  →  Title page
Master B  →  Introduction text (2–3 pages, two-column)
Master B  →  Table 1 (override to single column or place table across full width)
Master B  →  Narrative between tables (two-column)
Master B  →  Table 2 (single column override)
Master B  →  Narrative, Table 3 intro (two-column)
Master B  →  Table 3 (single column override)
Master B  →  Table 4 (single column override)
Master B  →  Table 5 (single column override — may need landscape page rotation)
Master B  →  Table 6 (fits two-column layout)
Master B  →  Origin and Function text, Table 7 (single column override for table)
Master B  →  Expression section, Table 8 (fits two-column layout)
Master B  →  Food and Feed Safety sections (two-column)
Master B  →  Compositional Analysis sections (two-column)
Master B  →  Conclusion (two-column)
Master B  →  References (two-column, Ref_Entry style handles sizing)
```

### 5.3 Layers

| Layer | Contents |
|-------|----------|
| **Background** | Color blocks, sidebar fills, accent bars |
| **Images** | Photos, diagrams, logos |
| **Text** | All text frames |
| **Guides** | Non-printing guides and annotations |

---

## 6. Paragraph Styles

### 6.1 Style Naming Convention

Use a prefix system for easy identification in the styles panel:

- `Title_` — Title page elements
- `Head_` — Headings (all levels)
- `Footer_` — Footer elements (running header, page number)
- `Body_` — Main body text
- `Table_` — Table text
- `Ref_` — Reference list
- `Meta_` — Keywords, copyright, license

### 6.2 Core Paragraph Styles

| Style Name | Font | Size | Leading | Color | Alignment | Space Before/After | Other |
|------------|------|------|---------|-------|-----------|--------------------|-------|
| **Title_title** | Source Sans Pro Bold | 26.5pt | — | Black | Left | 0 / 12pt | Tracking -4 |
| **Title_org** | Source Sans Pro Semibold | 11pt | — | AFSI Dark Gray | Left | 0 / 3pt | Organization name (Title_Subtitle folded into this) |
| **Title_address** | Source Sans Pro Regular | 10pt | — | AFSI Dark Gray | Left | 0 / 3pt | Organization address |
| **Title_date** | Source Sans Pro Regular | 10pt | — | AFSI Dark Gray | Left | 20pt / 0 | Date of publication |
| **Head_Section** | Source Sans Pro Bold | 12pt | 16pt | AFSI Blue #4397D2 | Left | 8pt / 4pt | All caps (via Basic Character Formats > Case: All Caps — source text should be title case); Keep with next 2 lines; GREP styles for scientific names and gene names → Char_BoldItalic |
| **Head_SubsectionNumbered** | Source Sans Pro Semibold | 10.5pt | 14pt | AFSI Gray #636466 | Left | 3pt / 3pt | Keep with next 2 lines; GREP styles for scientific names and gene names → Char_SemiboldItalic |
| **Head_SubsectionUnnumbered** | Source Sans Pro Semibold | 10pt | 13pt | AFSI Gray #636466 | Left | 3pt / 3pt | Keep with next 2 lines; GREP styles for scientific names and gene names → Char_SemiboldItalic |
| **Head_SubSubSectionUnnumbered** | Source Sans Pro Semibold | 10.5pt | 14pt | Black | Left | 3pt / 3pt | Keep with next 2 lines; same as Head_SubsectionNumbered but Black; GREP styles for scientific names and gene names → Char_SemiboldItalic |
| **Head_CropName** | Source Sans Pro Semibold | 10.5pt | 14pt | Black | Left | 3pt / 3pt | Keep with next 2 lines; GREP styles for scientific names and gene names → Char_SemiboldItalic |
| **Head_RunIn** | Source Sans Pro Bold | 9.5pt | 13pt | Black | Left (inline with following text) | 0 / 0 | Run-in heading style — bold text at start of paragraph, followed by regular body text on same line; for minor sub-sections like "Allergenicity prediction based on bioinformatics" |
| **Footer_header** | Source Sans Pro Light | 9.5pt | — | AFSI Dark Gray | Left | — | Running header text in the footer area |
| **Footer_pageno** | Source Sans Pro Light | 9pt | — | AFSI Dark Gray | Right (flush right) | — | Page number in the footer |
| **Body_text** | Source Sans Pro Regular | 9.5pt | 13pt | Black (#000000 or near-black #333333) | Justified, last line left | 0 / 4pt | First line indent: 0 (use space-after instead); Hyphenation ON |
| **Body_BulletL1** | Source Sans Pro Regular | 9.5pt | 13pt | Black | Justified, last line left | 0 / 4pt | Left indent: 0.25", first line left indent: -0.25"; bullet: •; Hyphenation ON (same controls as Body_text — these are full-sentence bullets) |
| **Body_BulletL2** | Source Sans Pro Regular | 9pt | 12pt | Black | Left | 0 / 3pt | Left indent: 0.5", first line left indent: -0.25" |
| **Body_Footnote** | Source Sans Pro Regular | 7.5pt | 10pt | AFSI Gray | Left | 0 / 3pt | GREP style: `^\d+(?=—)` → Char_Superscript (auto-superscripts the footnote number before the em dash) |
| **Table_Heading** | Source Sans Pro Regular | 10pt | 13pt | Black, Tint: 90% | Left | 6pt / 3pt | Span Columns: Span All; GREP style: `Table \d+:` → Char_Semibold; placed above a table when colleagues include a table title separate from the caption |
| **Table_Header** | Source Sans Pro Semibold | 7.5pt | 10pt | White | Center | 2pt / 2pt | White text; background fill is controlled by the cell style (CStyle_Header = AFSI Gray, CStyle_HeaderGreen = AFSI Green) |
| **Table_Header_leftalign** | Source Sans Pro Semibold | 7.5pt | 10pt | White | Left | 2pt / 2pt | Same as Table_Header but left-aligned; used by CStyle_Header_middle_leftalign for TStyle_Simple |
| **Table_Body** | Source Sans Pro Regular | 7pt | 9.5pt | Black | Left or Center | 1pt / 1pt | |
| **Table_BodyCenter_X** | Source Sans Pro Bold | 7.5pt | 10pt | AFSI Green #43BEA2 | Center | 1pt / 1pt | For "x" approval markers — bold green centered |
| **Table_Caption** | Source Sans Pro Semibold | 8.5pt | 12pt | Black | Left | 6pt / 3pt | "Table X." prefix in bold; allow multi-line captions |
| **Table_CaptionSub** | Source Sans Pro Regular | 8pt | 11pt | AFSI Gray | Left | 0 / 3pt | For qualifying notes below table captions (e.g., "Approvals after 2020 are highlighted.") |
| **Table_BodyDescriptive** | Source Sans Pro Regular | 7pt | 9.5pt | Black | Left | 2pt / 2pt | For text-heavy table cells (Table 7 — expression cassettes); allows text wrapping with generous padding |
| **Table_Note** | Source Sans Pro Regular | 7pt | 9.5pt | AFSI Gray | Left | 3pt / 0 | For general table notes |
| **Table_FootNote** | Source Sans Pro Regular Italic | 7.5pt | 10pt | AFSI Gray #636466 | Left | 4pt / 0 | For table-specific footnotes with superscript letters (a, b) placed directly below the table; distinct from document-level footnotes; GREP style: `^.` → Char_Superscript (auto-superscripts the first character) |
| **Table_Span** | (inherits from Basic Paragraph) | — | — | — | — | — | Only setting: Span Columns: Span All. Apply to the paragraph that an inline table sits on, so the table spans the full text area instead of being confined to a single column. No other formatting — the table's own styles control its appearance. |
| **Ref_Entry** | Source Sans Pro Regular | 8pt | 11pt | Black | Justified, last left | 0 / 4pt | Left indent: 0.25", first line left indent: -0.25" |
| **Meta_keywords** | Source Sans Pro Regular | 9.5pt | — | Black | Left | 0 / 3pt | Keywords in sidebar box; normal case |
| **Meta_keywords_header** | Source Sans Pro Semibold | 9.5pt | — | AFSI Dark Gray | Left | 0 / 6pt | "KEY WORDS" label; All caps |
| **Meta_copyright** | Source Sans Pro Semibold | 8pt | — | Black | Left | 6pt / 0 | Year and organization name; Tracking -2 |
| **Meta_creativecommons** | Source Sans Pro Regular | 7.5pt | — | Black | Left | 12pt / 0 | Creative Commons license line |

### 6.3 Key Paragraph Style Settings to Bake In

**For all heading styles:**
- **Keep Options:** Head_Section, Head_SubsectionNumbered, Head_SubSubSectionUnnumbered, Head_SubsectionUnnumbered, Head_CropName → "Keep with Next: 2 lines" and "Keep Lines Together: All Lines"

**For Body_text and other body styles:**
- **Widow/Orphan control:** Set minimum 2 lines at start and end of each paragraph (InDesign: Keep Options → Start/End: 2)

**For Body_text:**
- Hyphenation: ON, but with limits (see Section 8)
- Optical margin alignment: ON (via Story panel, but can be set at the document level)

**For Body_BulletL1:**
- Same justification and hyphenation settings as Body_text (the Codex factors list uses full sentences with semicolons)
- Widow/orphan control: minimum 2 lines at start and end

**For Head_CropName:**
- Apply to crop-specific sub-sections that appear as ALL CAPS in the Word file (MAIZE, COTTON, etc.) — change to title case (Maize, Cotton, Cowpea, Eucalyptus, Sugarcane)
- Same formatting as Head_SubSubSectionUnnumbered (Semibold, Black, 10.5pt)

**For Head_RunIn:**
- This is a run-in heading style where the heading text is bold and continues on the same line as the body text
- In InDesign: set as a character style nested within Body_text via a GREP style, or create as a separate paragraph style with "Next Style: Body_text" and no paragraph break after

**For Ref_Entry:**
- Left indent: 0.25", first line left indent: -0.25" for parenthetical reference citations
- No hyphenation (URLs and author names shouldn't break arbitrarily)

---

## 7. Character Styles

| Style Name | Purpose | Settings |
|------------|---------|----------|
| **Char_Bold** | Inline bold emphasis | Source Sans Pro Bold |
| **Char_Italic** | Inline italic (species names, gene names, journal titles); also used by body text GREP for scientific names | Source Sans Pro Italic |
| **Char_BoldItalic** | Combined emphasis; also used by Head_Section GREP for scientific/gene names in bold headings | Source Sans Pro Bold Italic |
| **Char_Semibold** | Semibold emphasis; also used by Table_Heading GREP for "Table N:" prefix | Source Sans Pro Semibold |
| **Char_SemiboldItalic** | Used by heading GREP styles for scientific/gene names in semibold headings | Source Sans Pro Semibold Italic |
| **Char_Superscript** | Footnote references, chemical formulas | Superscript position |
| **Char_Subscript** | Chemical formulas | Subscript position |
| **Char_Hyperlink** | Hyperlinks in the document body | Source Sans Pro Regular, AFSI Blue (or underline) |
| **Char_RefURL** | URLs in references — smaller size | Source Sans Pro Regular, 7.5pt, AFSI Blue |
| **Char_GeneItalic** | Gene names (cry1Ab) in italic per convention | Source Sans Pro Italic |
| **Char_ProteinRoman** | Protein names (Cry1Ab) in roman — default, but useful for toggling | Source Sans Pro Regular |
| **Char_FootnoteNumberinBody** | Superscript footnote number in body text | Superscript, Source Sans Pro Semibold |
| **Char_TableFootnoteRef** | Superscript letter (a, b) in table cells for table-specific footnotes | Superscript, Source Sans Pro Regular |
| **Char_TableHeaderText** | Override for white text on colored header row | Source Sans Pro Semibold, White |
| **Char_Green** | Colored emphasis | AFSI Green |
| **Char_No Break** | Prevent line breaks within applied text | No Break checked; no other formatting changes |

---

## 8. GREP Styles & Automation

GREP styles are applied within paragraph styles (Paragraph Style Options → GREP Style tab). They automatically apply character styles based on pattern matching, eliminating much of the manual formatting work.

### 8.1 Preventing Orphaned Titles and Honorifics

**Purpose:** Prevent "Dr." or "Mr." or single initials from being separated from the name that follows.

```
# Apply a no-break character style to title + following word
GREP: ((?:Dr|Mr|Mrs|Ms|Prof|St)\.\s\S+)
Style: Char_No Break character style
```

**No Break character style:** Create a character style with "No Break" checked under Basic Character Formats.

### 8.2 Preventing Orphaned Parenthetical Citations

**Purpose:** Keep parenthetical citations (e.g., `(AFSI 2016)`, `(Codex 2003a, 2003b)`) attached to the preceding word. The updated draft uses parenthetical citation style, not numbered references.

```
# Keep parenthetical citations with preceding word
# Matches patterns like (Author Year) and (Author Year; Author Year)
GREP: (\s\([A-Z][A-Za-z\s&]+,?\s?\d{4}[a-z]?(?:[,;]\s?[A-Z][A-Za-z\s&]+,?\s?\d{4}[a-z]?)*\))
Style: Char_No Break
```

**Note:** This pattern may be too aggressive for very long multi-citation parentheticals. Test with real content and adjust. For short citations like `(AFSI 2016)` or `(Codex 2003a)`, it works well.

### 8.3 Non-Breaking Spaces for Common Pairs

**Purpose:** Keep units with their numbers, abbreviations with context.

```
# Number + unit (e.g., "5 mg", "130 kDa", "90 days")
GREP: (\d+\.?\d*\s?(?:mg|µg|ug|ng|g|kg|kDa|Da|mL|µL|L|bp|kb|days?|hours?|min|sec|%))
Style: Char_No Break

# Compound unit expressions (e.g., "ug/g fresh weight", "mg/kg body weight", "mg/kg bodyweight")
GREP: (\d+\.?\d*\s?(?:ug|µg|mg)/(?:g|kg)\s(?:fresh\sweight|body\s?weight|dry\sweight|bw|fw|dw))
Style: Char_No Break

# "Table X", "Figure X", "Event X", etc.
GREP: ((?:Table|Tables|Figure|Fig\.|Appendix|Section|Event|Line)\s\d+(?:[-–]\d+)?)
Style: Char_No Break
```

### 8.4 Auto-Italicizing Scientific Names

**Purpose:** Automatically italicize genus/species names that appear frequently in the text.

```
# Common species in the Cry1Ab monograph — expanded for 6 crops
GREP: (Bacillus thuringiensis|B\. thuringiensis|Zea mays|Z\. mays|Gossypium hirsutum|G\. hirsutum|Oryza sativa|O\. sativa|Manduca sexta|M\. sexta|Vigna unguiculata|V\. unguiculata|Eucalyptus\ssp\.|Saccharum\ssp\.|Arabidopsis thaliana|A\. thaliana)
Style: Char_Italic

# Bt subspecies
GREP: ((?:subsp\.|var\.)\s\w+)
Style: Char_Italic

# Generic genus-species pattern (capitalized genus + lowercase species)
# Use with caution — may over-match
GREP: (?<!\w)([A-Z][a-z]+\s[a-z]{2,})(?=[\s,\.\);])
Style: Char_Italic
```

**Note:** The generic pattern should be tested carefully. You may prefer to handle the most common species explicitly and do a manual pass for edge cases. For other monographs (Cry1Ac, EPSPS, PAT/BAR), different species may appear — add them to the explicit list as needed.

### 8.5 Auto-Italicizing Gene Names

**Purpose:** Gene names (lowercase italic) vs. protein names (roman) follow biological convention.

```
# Gene names — expanded for all genes appearing in the updated draft
# Matches: cry1Ab, cry2Ab, cry1Fa, cry1Bb, cry2Aa, vip3Aa, cp4-epsps, epsps, pat, bar
GREP: (?<!\w)(cry\d[A-Z][a-z]\d?|cry\d[A-Z]{2}\d?|vip\d[A-Z][a-z]?\d?|pat|bar|epsps|cp4[\s-]epsps)(?!\w)
Style: Char_GeneItalic
```

**Caution:** This will need refinement for each monograph since the gene/protein naming varies. The pattern now captures the broader range of Cry gene variants (cry1Fa, cry1Bb, cry2Aa, cry2Ab) that appear in the stacked events discussion. Test thoroughly.

### 8.6 Weight-Matched Italics for Headings

**Purpose:** The body text GREP styles (Sections 8.4–8.5) apply Char_Italic and Char_GeneItalic, which are Regular weight. In headings, this would downgrade Bold/Semibold text to Regular Italic, looking wrong. These GREP styles use weight-matched character styles instead.

**Applied to Head_Section** (Bold headings):

```
# Scientific names
GREP: (Bacillus thuringiensis|B\. thuringiensis|Zea mays|Z\. mays|Gossypium hirsutum|G\. hirsutum|Oryza sativa|O\. sativa|Manduca sexta|M\. sexta|Vigna unguiculata|V\. unguiculata|Eucalyptus\ssp\.|Saccharum\ssp\.|Arabidopsis thaliana|A\. thaliana)
Style: Char_BoldItalic

# Subspecies
GREP: ((?:subsp\.|var\.)\s\w+)
Style: Char_BoldItalic

# Gene names
GREP: (?<!\w)(cry\d[A-Z][a-z]\d?|cry\d[A-Z]{2}\d?|vip\d[A-Z][a-z]?\d?|pat|bar|epsps|cp4[\s-]epsps)(?!\w)
Style: Char_BoldItalic
```

**Applied to Head_SubsectionNumbered, Head_SubsectionUnnumbered, Head_SubSubSectionUnnumbered, Head_CropName** (Semibold headings):

```
# Scientific names
GREP: (Bacillus thuringiensis|B\. thuringiensis|Zea mays|Z\. mays|Gossypium hirsutum|G\. hirsutum|Oryza sativa|O\. sativa|Manduca sexta|M\. sexta|Vigna unguiculata|V\. unguiculata|Eucalyptus\ssp\.|Saccharum\ssp\.|Arabidopsis thaliana|A\. thaliana)
Style: Char_SemiboldItalic

# Subspecies
GREP: ((?:subsp\.|var\.)\s\w+)
Style: Char_SemiboldItalic

# Gene names
GREP: (?<!\w)(cry\d[A-Z][a-z]\d?|cry\d[A-Z]{2}\d?|vip\d[A-Z][a-z]?\d?|pat|bar|epsps|cp4[\s-]epsps)(?!\w)
Style: Char_SemiboldItalic
```

**Note:** These are the same patterns from Sections 8.4 and 8.5 — only the character style target changes. When you add new species or gene names to the body text patterns, add them to the heading patterns too.

### 8.7 Preventing Bad Hyphenation in Scientific Terms

**Purpose:** Prevent hyphenation of key technical terms, event names, and regulatory body abbreviations.

```
# Protein and event names — expanded for all events in the updated draft
GREP: (Cry1Ab|Cry1Ac|Cry1Bb|Cry1F[a]?|Cry2Ab|Cry2Aa|Cry3Bb1|Cry34Ab1|Cry35Ab1|Vip3Aa|EPSPS|PAT/BAR)
Style: Char_No Break

# Transformation event names (single events)
GREP: (MON810|MON801|MON802|MON809|MON863|MON88017|MON89034|MON87427|MON87411|MON87419|Bt11|BT176|Bt176|Bt10|COT67B|NK603|GA21|MIR604|MIR162|TC1507|T304-40?|DBN9936|DBN9336|GHB119|GHB614|GHB811|COT102|MON88701|AAT709A|LP007-1|LP026-2|CTC175-?A|CTC20BT|1521K059)
Style: Char_No Break

# Regulatory body abbreviations — prevent hyphenation
GREP: (CTNBio|FSANZ|FZANS|EFSA|USEPA|USFDA|US\sFDA|CFIA|J-BCH|NBMA|CSIR|NBA|USDA|APHIS|OECD|ISAAA|FAO|WHO)
Style: Char_No Break

# Prevent hyphenation of "Bacillus thuringiensis" and "B. thuringiensis"
GREP: (Bacillus\sthuringiensis|B\.\sthuringiensis)
Style: Char_No Break
```

### 8.8 Normalizing and Protecting Stacked Event Names

**Purpose:** Stacked event names use multiplication signs (×) and can be very long. The draft inconsistently uses "x" and "×". This two-step approach first normalizes, then protects.

**Step 1 — Find/Change (run once after import, not a GREP style):**

```
# Normalize lowercase "x" between event names to proper multiplication sign ×
# Run as Find/Change GREP, not as a paragraph GREP style
Find: (?<=\d)\sx\s(?=[A-Z])
Change: \s×\s
```

**Step 2 — GREP style (baked into paragraph styles):**

```
# Keep short stacked event names together (2–3 events)
# For very long stacks (4+ events), these will be too long for No Break
# and should be allowed to break at × signs
GREP: ([A-Z][A-Za-z0-9]+\s×\s[A-Z][A-Za-z0-9]+(?:\s×\s[A-Z][A-Za-z0-9]+)?)
Style: Char_No Break
```

**Note:** Some stacked event names in the draft are extremely long (e.g., "MON87427 × MON89034 × MON810 × MIR162 × MON87411 × MON87419"). These will need to be allowed to break. The GREP above protects 2–3 event stacks. For longer names, manual soft returns at × signs may be needed during the review pass.

### 8.9 Discretionary Hyphens for Long URLs

Rather than a GREP style, consider applying a "URL Break" character style that allows breaks after `/` and `.` but not mid-word. In practice, for the reference list:

- Set the Ref_Entry paragraph style to **No Hyphenation**
- Use Find/Change (GREP) to insert discretionary line breaks (`\n` soft return or zero-width spaces) after URL path separators

Alternatively, you can use this GREP style on the Ref_Entry style:

```
# Allow break opportunities after / in URLs
GREP: (https?://\S+)
Style: Char_RefURL (with a slightly smaller size and permissive break settings)
```

### 8.10 Tightening Justification to Reduce Rivers

In the **Justification** settings of Body_text and Body_BulletL1:

| Setting | Minimum | Desired | Maximum |
|---------|---------|---------|---------|
| Word Spacing | 85% | 100% | 110% |
| Letter Spacing | -2% | 0% | 2% |
| Glyph Scaling | 98% | 100% | 102% |

These tighter ranges, combined with the Adobe Paragraph Composer (not Single-line), will reduce visible white rivers in justified text.

### 8.11 Auto-Semibold for Table Heading Prefix

**Applied to:** Table_Heading paragraph style

| GREP Pattern | Character Style | What It Matches |
|-------------|----------------|-----------------|
| `Table \d+:` | Char_Semibold | "Table 1:", "Table 2:", etc. — the prefix and colon |

This GREP style works in tandem with the Find/Change step in the checklist that converts "Table 1." (period) to "Table 1:" (colon). Once the colon is in place, this GREP style automatically applies Char_Semibold to the prefix, making it visually distinct from the rest of the heading text.

### 8.12 Auto-Superscript for Table Footnote Letters

**Applied to:** Table_FootNote paragraph style

| GREP Pattern | Character Style | What It Matches |
|-------------|----------------|-----------------|
| `^.` | Char_Superscript | The first character of the paragraph (a, b, c, etc.) |

Table footnotes start with a letter or number that should be superscripted to match the in-cell reference. This GREP style handles it automatically — no manual Char_Superscript application needed.

### 8.13 Auto-Superscript for Document Footnote Numbers

**Applied to:** Body_Footnote paragraph style

| GREP Pattern | Character Style | What It Matches |
|-------------|----------------|-----------------|
| `^\d+(?=—)` | Char_Superscript | One or more digits at the start of the paragraph, only when followed by an em dash |

InDesign's native footnotes follow a "number + em dash" pattern (e.g., "1—GE crops are..."). The lookahead `(?=—)` ensures only the number is matched and superscripted, not the em dash itself. This eliminates the need to manually apply Char_Superscript to footnote numbers.

### 8.14 Widow and Orphan Prevention Summary

InDesign doesn't have a single "GREP for widows" — this is handled through a combination of:

1. **Keep Options** on all body paragraph styles (minimum 2 lines at start and end)
2. **Keep Options** on heading styles (Keep with Next: 2 lines, Keep Lines Together: All Lines)
3. **Balance Ragged Lines** on left-aligned styles (like headings)
4. **Manual review** for final polish — use the Char_No Break character style on the last 2–3 words of a paragraph to pull a widow back, or insert a soft return
5. **Span/Split columns:** If a heading falls at a column break, "Keep with Next" ensures it pulls content forward

---

## 9. Table Styling Strategy

### 9.1 Table Types Overview

The updated draft contains 8 tables across 3 distinct structural types:

| Table | Content | Type | Width Needs | Rows | Columns |
|-------|---------|------|-------------|------|---------|
| **Table 1** | Maize singles pre-2010 | Approval matrix | Full-width portrait or landscape | ~8 events | ~18 countries |
| **Table 2** | Maize singles post-2010 | Approval matrix | May fit portrait full-width | ~4 events | ~12 countries |
| **Table 3** | MON810 breeding stacks | Approval matrix | Full-width portrait | ~21 events | ~15 countries |
| **Table 4** | Bt11 breeding stacks | Approval matrix | Full-width portrait or landscape | ~28 events | ~19 countries |
| **Table 5** | Cotton/cowpea/eucalyptus/rice/sugarcane | Approval matrix | Landscape likely required | ~20+ events across 6 crops | ~21 countries |
| **Table 6** | Regulatory document references | Text reference | Fits two-column body | ~9 entries | 3 columns |
| **Table 7** | Expression cassette genetic elements | Descriptive/text-heavy | Full-width portrait | ~7 events | 5 columns (with long text) |
| **Table 8** | Expression data (µg/g fresh weight) | Data table | Fits two-column body | ~20 rows | 5 columns |

### 9.2 Table Style: Approval Matrix (Tables 1–5)

These are the most complex elements. Strategy varies by table width:

**For narrower tables (Table 2 — ~12 country columns):**
- Can fit on a portrait full-width page (override Master B columns to 1)
- Rotated column headers (90° text rotation) for country names
- Standard 7pt text in cells

**For wider tables (Tables 3–4 — 15–19 country columns):**
- Portrait full-width page (override columns to 1), tighter column spacing
- Rotated headers essential
- May need 6.5pt text in cells

**For the widest table (Table 5 — 21 country columns + crop + event):**
- **Landscape page rotation** within the document (InDesign: select page, use Pages panel to rotate)
- Or: Split across two pages (crops 1–3, crops 4–6)

**Common elements for all approval tables:**
- "x" markers: Center-aligned, AFSI Green Bold for visual clarity (Table_BodyCenter_X style)
- Species grouping: Use merged cells for the crop/species column, with species name in italic
- Empty cells: Leave blank (no dash or "—")
- **Post-2020 highlighting:** For Tables 2–4 that note "Approvals after 2020 are highlighted," use AFSI Blue (#4397D2) at 25% tint as cell fill. Create a cell style `CStyle_Highlighted` for this.

### 9.3 Table Style: Text Reference (Table 6)

Simple 3-column table (Crop, Event, Regulatory Document):
- Fits within the two-column body layout
- Use standard TStyle_Standard
- Left-aligned text in all columns
- Regulatory Document column will contain multiple comma-separated references — allow text wrapping

### 9.4 Table Style: Descriptive/Text-Heavy (Table 7)

Table 7 (expression cassettes) has 5 columns with paragraph-length content in the Promoter, Enhancer, and Gene/Protein Modifications columns:
- **Full-width portrait page** (override Master B columns to 1)
- Use Table_BodyDescriptive paragraph style (7pt, left-aligned, generous cell padding)
- Minimum row height: 0.5" to accommodate multi-line text
- Column proportions (approximate): Event 12% / Promoter 22% / Enhancer 20% / Gene/Modifications 32% / Reference 14%
- Italic gene names within cells will need Char_GeneItalic applied (GREP should handle this automatically)

### 9.5 Table Style: Data Table (Table 8)

Table 8 (expression data) has 5 columns and fits within the two-column layout:
- Columns: Crop/Event, Year/Trial Locations, Tissue, Cry1Ab Levels, Reference
- Merged cells for crop/event spanning multiple tissue rows
- Superscript table footnotes (a, b) within cells — apply Char_TableFootnoteRef
- Table footnote text below table — apply Table_FootNote style
- Alternating row fills (white / Taupe at 30%)

### 9.6 How Table Styles and Cell Styles Work Together

Understanding the relationship between these style types is essential to making them save you time rather than cause frustration.

**The nesting model (built from the inside out):**

- **Paragraph Style** controls text appearance (font, size, color, alignment).
- **Cell Style** wraps around the paragraph style. It adds cell fill color, cell insets (padding), vertical text alignment, and optionally cell strokes. Critically, a cell style *contains* a paragraph style — when applied, it formats both the cell and the text inside it.
- **Table Style** wraps around the cell styles. It controls the outer table border, alternating row/column strokes (interior lines), alternating fills, and *assigns* which cell styles apply to which table regions (header rows, body rows, footer rows, left/right columns).

**The key rule for strokes:** Both table styles and cell styles can define strokes (lines), and when they conflict, **cell styles win**. This means if your cell style specifies any stroke value — even "0pt" — it will override whatever the table style says about interior lines. The solution is to set cell style strokes to **(Ignore)**, which is different from "None" or "0pt." Setting strokes to (Ignore) means "I have no opinion about strokes — let the table style handle all the lines." In the Strokes and Fills tab of the Cell Style dialog, you achieve this by deselecting all four sides of the stroke proxy (clicking the blue lines so they turn gray). When all sides are gray/deselected, the weight and color fields will show as grayed out, indicating (Ignore).

**The override problem with Word imports:** When you place a Word file, InDesign preserves all of Word's manual table formatting as overrides. These overrides sit at the top of the precedence chain, so they block your styles from working. The fix is in the Post-Placement Cleanup Checklist (Section 2).

**Precedence hierarchy (highest priority wins):**

1. Manual cell formatting (Cell Options dialog box)
2. Directly applied cell style
3. Cell style assigned by a table style
4. Manual table formatting (Table Options dialog box)
5. Table style

**The practical takeaway:** Let the table style own all the lines (strokes). Let the cell styles own fills and paragraph styles. Set cell style strokes to (Ignore) so they never fight with the table style. This keeps things predictable.

### 9.7 Building the Base Table Style — Step by Step

The following instructions create the base AFSI table: 1pt AFSI Gray outer border, 0.5pt AFSI Gray interior lines, AFSI Gray header row with white text, and subtle 0.25pt light gray column dividers within the header. Once this is working, it can be duplicated and modified for the approval matrix and other variations described in Sections 9.2–9.5.

**Step 1 — Confirm paragraph styles exist.**

You need these paragraph styles before creating cell styles (see Section 6.2 for full specs):

- **Table_Header:** Source Sans Pro Semibold, 7.5pt, white, centered
- **Table_Header_leftalign:** Source Sans Pro Semibold, 7.5pt, white, left-aligned
- **Table_Body:** Source Sans Pro Regular, 7pt, black, left-aligned

**Step 2 — Create Cell Style: CStyle_Header**

This is the style for interior header cells (not the leftmost or rightmost). It defines the gray fill and a subtle right-side column divider.

Open the Cell Styles panel (Window > Styles > Cell Styles). From the panel menu, choose New Cell Style.

- **General tab:**
  - Style Name: CStyle_Header
  - Paragraph Style: Table_Header
- **Text tab:**
  - Cell Insets: Top 0.0556" (4pt), Bottom 0.0556", Left 0.0556", Right 0.0556"
  - Vertical Justification: Align Center
- **Strokes and Fills tab:**
  - Cell Stroke proxy: Deselect ALL four sides first (click each blue line so it turns gray). Then select ONLY the right side (click it so it turns blue). Set: Weight 0.25pt, Color: AFSI Light Gray (#91908F), Type: Solid. The top, bottom, and left remain at (Ignore) — the table style and table border handle those edges.
  - Cell Fill Color: AFSI Gray (#636466), Tint: 100%
- Click OK.

**Step 3 — Create Cell Style: CStyle_HeaderLeft**

This is for the leftmost header cell only. It matches CStyle_Header but adds a 1pt left stroke to stay flush with the table border.

From the Cell Styles panel menu, choose New Cell Style (or duplicate CStyle_Header and edit).

- **General tab:**
  - Style Name: CStyle_HeaderLeft
  - Paragraph Style: Table_Header
- **Text tab:**
  - Same as CStyle_Header (4pt insets all sides, Align Center)
- **Strokes and Fills tab:**
  - Cell Stroke proxy: Deselect ALL sides. Then select the **left** side: Weight 1pt, Color: AFSI Gray (#636466), Type: Solid. Then deselect left, select the **right** side: Weight 0.25pt, Color: AFSI Light Gray (#91908F), Type: Solid. Top and bottom remain at (Ignore).
  - Cell Fill Color: AFSI Gray (#636466), Tint: 100%
- Click OK.

**Step 4 — Create Cell Style: CStyle_HeaderRight**

This is for the rightmost header cell only. It adds a 1pt right stroke to stay flush with the table border.

From the Cell Styles panel menu, choose New Cell Style (or duplicate CStyle_Header and edit).

- **General tab:**
  - Style Name: CStyle_HeaderRight
  - Paragraph Style: Table_Header
- **Text tab:**
  - Same as CStyle_Header (4pt insets all sides, Align Center)
- **Strokes and Fills tab:**
  - Cell Stroke proxy: Deselect ALL sides. Then select ONLY the **right** side: Weight 1pt, Color: AFSI Gray (#636466), Type: Solid. Top, bottom, and left remain at (Ignore).
  - Cell Fill Color: AFSI Gray (#636466), Tint: 100%
- Click OK.

**Step 5 — Create Cell Style: CStyle_Header_middle_leftalign**

Same as CStyle_Header but uses the left-aligned paragraph style. This is the default header cell style for TStyle_Simple.

Duplicate CStyle_Header and edit:

- **General tab:**
  - Style Name: CStyle_Header_middle_leftalign
  - Paragraph Style: Table_Header_leftalign
- **Text tab:**
  - Same as CStyle_Header (4pt insets all sides, Align Center)
- **Strokes and Fills tab:**
  - Same as CStyle_Header (right side only, 0.25pt AFSI Light Gray; Cell Fill: AFSI Gray)
- Click OK.

**Step 6 — Create Cell Style: CStyle_Header_left_leftalign**

Leftmost header cell with left-aligned text. Matches CStyle_HeaderLeft but uses the left-aligned paragraph style.

Duplicate CStyle_HeaderLeft and edit:

- **General tab:**
  - Style Name: CStyle_Header_left_leftalign
  - Paragraph Style: Table_Header_leftalign
- **Text tab:**
  - Same as CStyle_HeaderLeft (4pt insets all sides, Align Center)
- **Strokes and Fills tab:**
  - Same as CStyle_HeaderLeft (left side 1pt AFSI Gray, right side 0.25pt AFSI Light Gray; Cell Fill: AFSI Gray)
- Click OK.

**Step 7 — Create Cell Style: CStyle_Header_right_leftalign**

Rightmost header cell with left-aligned text. Matches CStyle_HeaderRight but uses the left-aligned paragraph style.

Duplicate CStyle_HeaderRight and edit:

- **General tab:**
  - Style Name: CStyle_Header_right_leftalign
  - Paragraph Style: Table_Header_leftalign
- **Text tab:**
  - Same as CStyle_HeaderRight (4pt insets all sides, Align Center)
- **Strokes and Fills tab:**
  - Same as CStyle_HeaderRight (right side only, 1pt AFSI Gray; Cell Fill: AFSI Gray)
- Click OK.

**Step 8 — Create Cell Style: CStyle_Body**

From the Cell Styles panel menu, choose New Cell Style.

- **General tab:**
  - Style Name: CStyle_Body
  - Paragraph Style: Table_Body
- **Text tab:**
  - Cell Insets: Top 0.0278" (2pt), Bottom 0.0278", Left 0.0556" (4pt), Right 0.0556"
  - Vertical Justification: Align Center
- **Strokes and Fills tab:**
  - Cell Stroke: ALL four sides deselected (gray) — fully at (Ignore). The table style handles all body cell lines.
  - Cell Fill Color: (Ignore) — leave the fill unspecified so the table style can control alternating fills later.
- Click OK.

**Step 9 — Create Cell Style: CStyle_BodyBottom**

This is for the bottom row of body cells only. It matches CStyle_Body but adds a 1pt bottom stroke to stay flush with the table border.

From the Cell Styles panel menu, choose New Cell Style (or duplicate CStyle_Body and edit).

- **General tab:**
  - Style Name: CStyle_BodyBottom
  - Paragraph Style: Table_Body
- **Text tab:**
  - Same as CStyle_Body (2pt top/bottom, 4pt left/right insets, Align Center)
- **Strokes and Fills tab:**
  - Cell Stroke proxy: Deselect ALL sides. Then select ONLY the **bottom** side: Weight 1pt, Color: AFSI Gray (#636466), Type: Solid. Top, left, and right remain at (Ignore).
  - Cell Fill Color: (Ignore)
- Click OK.

**Step 10 — Create Table Style: TStyle_Simple**

Open the Table Styles panel (Window > Styles > Table Styles). From the panel menu, choose New Table Style.

- **General tab:**
  - Style Name: TStyle_Simple
  - Cell Styles section (dropdown menus at the bottom of General):
    - Header Rows: CStyle_Header_middle_leftalign
    - Footer Rows: [None]
    - Body Rows: CStyle_Body
    - Left Column: [None]
    - Right Column: [None]
- **Table Setup tab:**
  - Table Border — Weight: 1pt, Color: AFSI Gray (#636466), Type: Solid
  - Table Spacing — Space Before: 0.0833" (6pt), Space After: 0.0833" (6pt)
- **Row Strokes tab:**
  - Alternating Pattern: Every Other Row
  - First: 1 Rows, Weight: 0.5pt, Color: AFSI Gray (#636466), Type: Solid
  - Next: 1 Rows, Weight: 0.5pt, Color: AFSI Gray (#636466), Type: Solid
  - (Both set to the same values = consistent lines between every row, no visual alternation.)
  - Skip First: 0, Skip Last: 0
- **Column Strokes tab:**
  - Alternating Pattern: Every Other Column
  - First: 1 Columns, Weight: 0.5pt, Color: AFSI Gray (#636466), Type: Solid
  - Next: 1 Columns, Weight: 0.5pt, Color: AFSI Gray (#636466), Type: Solid
  - Skip First: 0, Skip Last: 0
- **Fills tab:**
  - Alternating Pattern: None
- Click OK.

**Step 11 — Test on a new table.**

Create a test table: Table > Insert Table. Set 5 body rows, 4 columns, 1 header row. In the Table Style dropdown at the bottom of the Insert Table dialog, choose TStyle_Simple. Click OK.

You should see: 1pt AFSI Gray outer border, 0.5pt AFSI Gray interior lines in the body, AFSI Gray header row with white left-aligned text, and 0.5pt AFSI Gray column dividers in the header (from the table style's column strokes). The header column dividers will be 0.5pt at this point — that's expected because the table style assigned CStyle_Header_middle_leftalign to all header cells, and only interior cells should have the 0.25pt dividers.

**Step 12 — Apply the edge and bottom row cell styles.**

The table style can only assign one cell style to all header cells and one to all body cells. To get correct edge behavior, manually apply the edge and bottom styles:

1. Click in the **leftmost** header cell. In the Cell Styles panel, click **CStyle_Header_left_leftalign** (for TStyle_Simple) or **CStyle_HeaderLeft** (for centered-header table styles).
2. Click in the **rightmost** header cell. In the Cell Styles panel, click **CStyle_Header_right_leftalign** (for TStyle_Simple) or **CStyle_HeaderRight** (for centered-header table styles).
3. Select the **entire bottom row** of body cells. In the Cell Styles panel, click **CStyle_BodyBottom**.

The table now has: 1pt edges flush with the table border on all sides — left and right header edges, and the bottom body edge. Interior lines are handled by the table style's row and column strokes.

This is three manual applications per table (left header, right header, bottom row). For tables with only 2 columns, use CStyle_Header_left_leftalign on the left cell and CStyle_Header_right_leftalign on the right cell (no interior cells needed).

**Step 13 — Applying styles to pasted tables (the override-clearing workflow).**

Tables pasted from Word will always carry manual formatting overrides that block your styles from applying. This is not a bug — InDesign preserves Word's formatting intentionally.

**Automated step:** Run **ClearTableOverrides.jsx** to clear all table, cell, and paragraph overrides on every table in the document at once, and convert first rows to header rows. This replaces manual steps 1–7 below.

**Manual steps (after running the script):**

1. **Apply the table style:** Click **TStyle_Simple** (or TStyle_Approvals, etc.) in the Table Styles panel.
2. **Force cell styles if needed:** If the cell styles still don't apply after step 1, select all header cells and Alt/Opt-click the appropriate header cell style, then select all body cells and Alt/Opt-click the appropriate body cell style. This pushes through any remaining resistance.
3. **Apply edge header styles:** Click in the leftmost header cell and apply CStyle_HeaderLeft (or the rotated/green variant). Click in the rightmost header cell and apply CStyle_HeaderRight (or variant).
4. **Apply bottom row style:** Select the entire bottom row of body cells and apply CStyle_BodyBottom (or CStyle_BodyApprovalBottom for approval tables).

The Alt/Opt-click is critical — it means "apply this style AND clear all overrides." A regular click applies the style but leaves overrides in place, which is why styles seem to not work.

**Why this happens:** InDesign's precedence chain puts manual formatting above styles. When Word's formatting comes in via paste, InDesign treats it as manual formatting (overrides), which sits at the top of the chain and blocks everything below it. Clearing overrides removes that layer and lets your styles take control.

### 9.8 Building the Approval Matrix Table Style — Step by Step

This style is for the large regulatory approval tables (Tables 1–5) that use rotated country-name headers and centered "x" markers. Rotated headers allow 18–21 country columns to fit on a portrait page without going to landscape.

**Step 1 — Create Cell Style: CStyle_HeaderRotated**

Duplicate CStyle_Header and edit (right-click > Duplicate Style, then double-click to edit):

- **General tab:**
  - Style Name: CStyle_HeaderRotated
  - Paragraph Style: Table_Header
- **Text tab:**
  - Cell Insets: Top 0.0556" (4pt), Bottom 0.0556", Left 0.0278" (2pt), Right 0.0278" (2pt) — narrower left/right since the text is rotated and these become the visual top/bottom of the header cell
  - Vertical Justification: Align Bottom (text anchors near the body rows)
  - Rotation: 270°
- **Strokes and Fills tab:**
  - Same as CStyle_Header: right side only, 0.25pt AFSI Light Gray (#91908F). Everything else at (Ignore).
  - Cell Fill Color: AFSI Gray (#636466)
- Click OK.

**Step 2 — Create Cell Style: CStyle_HeaderRotatedLeft**

Duplicate CStyle_HeaderLeft and edit:

- **General tab:** Style Name: CStyle_HeaderRotatedLeft
- **Text tab:** Same as CStyle_HeaderRotated (2pt left/right insets, Align Bottom, Rotation 270°)
- **Strokes and Fills tab:** Same as CStyle_HeaderLeft (left side 1pt AFSI Gray, right side 0.25pt AFSI Light Gray). Cell Fill: AFSI Gray (#636466).
- Click OK.

**Step 3 — Create Cell Style: CStyle_HeaderRotatedRight**

Duplicate CStyle_HeaderRight and edit:

- **General tab:** Style Name: CStyle_HeaderRotatedRight
- **Text tab:** Same as CStyle_HeaderRotated (2pt left/right insets, Align Bottom, Rotation 270°)
- **Strokes and Fills tab:** Same as CStyle_HeaderRight (right side 1pt AFSI Gray). Cell Fill: AFSI Gray (#636466).
- Click OK.

**Step 4 — Create Cell Style: CStyle_BodyApproval**

Duplicate CStyle_Body and edit:

- **General tab:**
  - Style Name: CStyle_BodyApproval
  - Paragraph Style: Table_BodyCenter_X (bold, AFSI Green, centered — for the "x" markers)
- **Text tab:**
  - Cell Insets: 2pt all sides (tighter than standard body cells since the cells only hold a single "x" or are empty)
  - Vertical Justification: Align Center
- **Strokes and Fills tab:**
  - Same as CStyle_Body: all sides at (Ignore), Cell Fill at (Ignore).
- Click OK.

**Step 5 — Create Cell Style: CStyle_BodyApprovalBottom**

This is for the bottom row of body cells in approval tables. It matches CStyle_BodyApproval but adds a 1pt bottom stroke to stay flush with the table border.

Duplicate CStyle_BodyApproval and edit:

- **General tab:**
  - Style Name: CStyle_BodyApprovalBottom
  - Paragraph Style: Table_BodyCenter_X
- **Text tab:**
  - Same as CStyle_BodyApproval (2pt all sides, Align Center)
- **Strokes and Fills tab:**
  - Cell Stroke proxy: Deselect ALL sides. Then select ONLY the **bottom** side: Weight 1pt, Color: AFSI Gray (#636466), Type: Solid. Top, left, and right remain at (Ignore).
  - Cell Fill Color: (Ignore)
- Click OK.

**Step 6 — Create Table Style: TStyle_Approvals**

Duplicate TStyle_Simple and edit:

- **General tab:**
  - Style Name: TStyle_Approvals
  - Cell Styles:
    - Header Rows: CStyle_HeaderRotated
    - Body Rows: CStyle_BodyApproval
    - Footer Rows: [None]
    - Left Column: [None]
    - Right Column: [None]
- **Table Setup tab:**
  - Table Border: 1pt, AFSI Gray (#636466), Solid (same as TStyle_Simple)
- **Row Strokes tab:**
  - Alternating Pattern: Every Other Row
  - First: 1 Rows, Weight: 0.25pt, Color: AFSI Light Gray (#91908F), Type: Solid
  - Next: 1 Rows, Weight: 0.25pt, Color: AFSI Light Gray (#91908F), Type: Solid
  - (Thinner than TStyle_Simple's 0.5pt — keeps the dense approval grid from feeling heavy)
  - Skip First: 0, Skip Last: 0
- **Column Strokes tab:**
  - Alternating Pattern: Every Other Column
  - First: 1 Columns, Weight: 0.25pt, Color: AFSI Light Gray (#91908F), Type: Solid
  - Next: 1 Columns, Weight: 0.25pt, Color: AFSI Light Gray (#91908F), Type: Solid
  - Skip First: 0, Skip Last: 0
- **Fills tab:**
  - Alternating Pattern: None (the approval tables don't use alternating row fills)
- Click OK.

**Step 7 — Test on a new table.**

Create a test table: Table > Insert Table. Set 8 body rows, 12 columns, 1 header row. Choose TStyle_Approvals from the Table Style dropdown before clicking OK.

Type country names in the header cells (e.g., "Australia", "Brazil", "Canada", "South Africa") and "x" in scattered body cells. The header text should appear rotated vertically with the AFSI Gray fill, and the body cells should show centered bold green "x" markers.

Then apply the edge and bottom styles: CStyle_HeaderRotatedLeft on the leftmost header cell, CStyle_HeaderRotatedRight on the rightmost, and select the entire bottom body row and apply CStyle_BodyApprovalBottom.

**Step 8 — Test on an imported table.**

Run ClearTableOverrides.jsx first, then apply TStyle_Approvals, the edge header styles, and CStyle_BodyApprovalBottom on the bottom row.

**Design notes for approval tables:**

- The first column (event names like "MON810" or "MON89034 × MON88017") will need to be wider than the country columns. Drag it wider manually after the style is applied — column widths aren't controlled by styles, so this won't create an override conflict.
- Header row height is driven by the longest country name. At 7.5pt, "South Africa" runs about 0.6" tall. This is normal and gives the table the vertical-header look common in scientific publications.
- If a table has more than ~18 country columns and still doesn't fit on a portrait page, consider reducing the Table_Header paragraph style size to 7pt for that table (local override), or splitting the table across two pages.

### 9.9 Additional Cell and Table Styles

Once TStyle_Simple and TStyle_Approvals are working, create these additional styles by duplicating and modifying. To duplicate: right-click the style in the panel, choose Duplicate Style, then double-click to edit.

**Additional Cell Styles:**

| Cell Style | Based On | Changes from Base |
|-----------|----------|-------------------|
| **CStyle_HeaderGreen** | CStyle_Header | Cell Fill: AFSI Green (#43BEA2) instead of AFSI Gray |
| **CStyle_HeaderGreenLeft** | CStyle_HeaderLeft | Cell Fill: AFSI Green (#43BEA2); Left stroke: 1pt AFSI Green |
| **CStyle_HeaderGreenRight** | CStyle_HeaderRight | Cell Fill: AFSI Green (#43BEA2); Right stroke: 1pt AFSI Green |
| **CStyle_BodyDescriptive** | CStyle_Body | Cell Insets: 4pt all sides; Paragraph Style: Table_BodyDescriptive; Vertical Justification: Align Top |
| **CStyle_BodyDescriptiveBottom** | CStyle_BodyDescriptive | Bottom stroke: 1pt AFSI Gray (same pattern as CStyle_BodyBottom) |
| **CStyle_Highlighted** | CStyle_Body | Cell Fill: AFSI Blue (#4397D2), Tint: 25% (for post-2020 approvals — apply manually to individual cells after table style is applied) |
| **CStyle_Species** | CStyle_Body | Paragraph Style: a bold italic variant for species grouping rows |

**Note on bottom row styles:** Every table type needs a "Bottom" variant of its body cell style with a 1pt AFSI Gray bottom stroke, to keep the last row flush with the 1pt table border. The pattern is always the same: duplicate the body cell style, deselect all stroke sides, select only the bottom side, set 1pt AFSI Gray. CStyle_BodyBottom, CStyle_BodyApprovalBottom, and CStyle_BodyDescriptiveBottom follow this pattern.

**Note on the Green header edge styles:** The left and right stroke colors should match the header fill (AFSI Green), not the table border (AFSI Gray). This way the header edges blend seamlessly with the green fill while the table border frames the overall table in gray.

**Additional Table Styles:**

| Table Style | Based On | Key Changes |
|------------|----------|-------------|
| **TStyle_Standard** | TStyle_Simple | Header Rows cell style: CStyle_HeaderGreen. (Remember to manually apply CStyle_HeaderGreenLeft and CStyle_HeaderGreenRight to edge cells.) |
| **TStyle_ApprovalsWide** | TStyle_Approvals | Body Rows cell style: a duplicate of CStyle_BodyApproval with 1pt all-side insets for tables with 20+ columns. |
| **TStyle_Descriptive** | TStyle_Simple | Body Rows cell style: CStyle_BodyDescriptive. |

---

## 10. Importing Word Content — Workflow

### 10.1 Pre-Import Checklist (for the Word files from colleagues)

Before working with Word files, standardize the Word styles where possible:

1. **Ask colleagues to use consistent Word heading styles:**
   - Heading 1 = major section heads (INTRODUCTION, ORIGIN AND FUNCTION OF CRY1AB, etc.)
   - Heading 2 = subsections (Mechanism of Cry1Ab insecticidal activity, etc.)
   - Heading 3 = sub-sub-sections (Acute toxicity studies, Allergenicity prediction, etc.)
   - Normal = body text
   - **Note:** The current draft uses inconsistent heading levels. Some sub-sections are bold Normal text, some are all-caps Normal text. These will need manual style application in InDesign.

2. **Crop sub-sections:** Ask colleagues to apply Heading 3 to crop names (MAIZE, COTTON, etc.) or flag them with a consistent marker. Currently these appear as all-caps body text.

3. **References:** The current draft uses parenthetical citations (Author Year). The reference list is not yet included. When it arrives, ask for plain text with consistent formatting.

4. **Tables:** Tables will be copy-pasted from the original Word file into InDesign individually during Phase 3 (see Section 2). They are excluded from the plain text export to avoid XML corruption.

5. **Table footnotes:** Ask colleagues to keep table-specific footnotes (superscript a, b) as regular text immediately below the table, not as Word's built-in footnote function.

6. **Footnotes:** Colleagues use manually superscripted numbers for shared footnotes (e.g., multiple sentences referencing footnote 4). These are incompatible with Word's native footnote system and with InDesign's auto-numbering. The workflow in Section 2 marks these before stripping formatting so they can be restored in InDesign.

7. **Images/charts:** Supply separately at 300dpi minimum.

### 10.2 Text Preparation in Word (Phase 1 detail)

The Word documents contain hidden XML formatting that causes InDesign to crash during recomposition. The solution is to strip all formatting before placement. See Section 2, Phase 1 for the step-by-step procedure. The key operations:

1. **Mark superscripts:** Find/Replace in Word with Find Format: Superscript, Replace with: `{{^&}}` (wraps all superscripted text in double curly braces that survive plain text export).
2. **Delete tables:** Remove all table grids but leave headings, captions, and footnotes. Add `[INSERT TABLE X HERE]` markers.
3. **Save as Plain Text (.txt):** UTF-8 encoding. This eliminates all hidden XML, manual formatting, embedded styles, and No Break overrides.

### 10.3 Plain Text Placement and First Pass (Phase 2 detail)

See Section 2, Phase 2 for the full checklist. The key principle: since the placed text has zero formatting, there are no overrides to fight. All paragraph styles apply cleanly on first click, and GREP styles activate immediately once Body_text is applied.

**Find/Change operations summary:**

1. Create native footnotes manually (Part A — first occurrence of each {{N}} marker)
2. Run **CleanupAfterPlacement.jsx** (handles all remaining Find/Change operations automatically):
   - Double spaces → single space
   - Extra returns → single return
   - Strip bullet characters
   - Tilde operator → standard tilde
   - Multiplication signs normalization
   - Table heading periods → colons
   - Superscript remaining {{}} markers
3. Title case headings: Run TitleCaseHeadings.jsx script (applies Title Case, lowercases articles/prepositions, fixes specific terms like GE)
4. Review title case results: Check for broken scientific terms and other edge cases

**Manual work required:**

- Apply paragraph styles to all headings (refer to original Word doc for structure)
- Restore italics not covered by GREP (journal names, occasional emphasis, Latin terms beyond the species list)

### 10.4 Two-Column Layout and Table Insertion (Phase 3 detail)

See Section 2, Phase 3 for the full checklist. The key principle: switch to two columns first while the document is table-free and stable, then build tables in isolation before inserting them.

**Workflow:**

1. Switch Master B to 2 columns, adjust pagination
2. Paste all tables from Word into a **staging text frame** (separate from the main flow)
3. Run **ClearTableOverrides.jsx** and **CleanupAfterPlacement.jsx** on the staging area
4. Manually apply table styles, edge header styles, and bottom row styles per table
5. Cut each formatted table from staging and paste into the main flow (replacing `[INSERT TABLE X HERE]` markers)
6. **Fallback:** If pasting inline crashes InDesign, use standalone text frames positioned over blank space instead

---

## 11. Step-by-Step Template Creation Sequence

Here is the recommended order of operations for building this template in InDesign:

**Phase 1 — Document Foundation**
1. Create new document with specifications from Section 5.1
2. Set up color swatches for all 11 AFSI brand colors (Section 4.1)
3. Install/verify Source Sans Pro font family is available
4. Create the two master pages (A and B) with guides and frames

**Phase 2 — Styles**
5. Build all character styles (Section 7) — including Char_No Break, Char_GeneItalic, Char_TableFootnoteRef
6. Build all paragraph styles (Section 6) — start with Body_text as the foundation, then build others. Remember to include the new styles: Head_SubSubSectionUnnumbered, Head_SubsectionUnnumbered, Head_CropName, Head_RunIn, Table_BodyCenter_X, Table_CaptionSub, Table_BodyDescriptive, Table_FootNote
7. Add GREP styles to the relevant paragraph styles (Section 8) — expanded patterns for species, genes, events, units
8. Set justification and hyphenation settings on Body_text, Body_BulletL1, and Ref_Entry
9. Build table styles and cell styles (Section 9.7 for base style, Section 9.8–9.9 for AFSI-branded variants)

**Phase 3 — Master Page Detailing**
10. Design Master A (title page layout) — place text frames with style assignments, sidebar box, accent elements
11. Design Master B (two-column body) — column guides, running header/footer with auto page numbers
12. Add AFSI logo to appropriate masters
13. Test page sequence flow per Section 5.2 page sequence

**Phase 4 — Placeholder Content**
14. Add placeholder text frames showing the typical flow: title page → intro body (2–3 pages) → approval tables → body → references
15. Type sample content into each frame to verify styles render correctly
16. Build sample tables for each of the 3 table types:
    - Approval matrix (Table 1 or 2 — smallest first to test)
    - Descriptive/text-heavy (Table 7)
    - Data table (Table 8 with table footnotes)
17. Test the GREP styles with real content from the Cry1Ab draft

**Phase 5 — Test & Refine**
18. Place the actual Cry1Ab Word document and verify the import mapping
19. Run the post-import cleanup sequence (Section 10.3)
20. Review for orphans, widows, bad breaks, and hyphenation issues
21. Verify the GREP patterns catch all scientific names, gene names, and event names in the actual content
22. Export a test PDF and check all colors, fonts, and alignment
23. Adjust GREP patterns and style settings as needed
24. Save as `.indt` (InDesign Template)

---

## 12. File Naming & Organization

```
AFSI_Monographs/
├── Templates/
│   ├── AFSI_Monograph_FFS_Template.indt      ← Food & Feed Safety template
│   ├── AFSI_Monograph_ENV_Template.indt      ← Environmental Safety template (if different)
│   └── AFSI_Monograph_Styles.indd            ← Style library source file
├── Scripts/
│   ├── TitleCaseHeadings.jsx                 ← Applies Title Case to all heading styles
│   ├── CleanupAfterPlacement.jsx             ← Runs all Find/Change operations in sequence
│   └── ClearTableOverrides.jsx               ← Clears Word overrides on all tables at once
├── Assets/
│   ├── AFSI_Logo_Color.ai (or .eps / .svg)
│   ├── AFSI_Logo_White.ai
│   └── Source_Sans_Pro/ (font files if needed)
├── Content/
│   ├── Cry1Ab_FFS_Updated.docx               ← Word files from colleagues
│   ├── Cry1Ac_FFS_Updated.docx
│   ├── EPSPS_FFS_Updated.docx
│   └── PAT_FFS_Updated.docx
└── Output/
    ├── Cry1Ab_FFS_2025.pdf
    └── ...
```

---

## 13. Quick-Reference Checklist for Each New Monograph

When a new Word file arrives from your colleagues, use this checklist:

**Pre-import:**
- [ ] Open template `.indt` → Save As new `.indd` with protein name
- [ ] Review Word file for heading consistency — note any all-caps or bold-only headings that need remapping

**Import & cleanup:**
- [ ] Place Word file into body text frames with style mapping (Section 10.2)
- [ ] Run the Post-Placement Workflow (Section 2): Phase 1 in Word, Phase 2 text styling in single column, Phase 3 tables and two-column layout

**Tables:**
- [ ] Rebuild approval matrix tables (Tables 1–5) in InDesign with appropriate TStyle
- [ ] Determine which tables need full-width layout (override to single column) vs. landscape rotation
- [ ] Apply post-2020 highlight styling where noted (CStyle_Highlighted)
- [ ] Rebuild or verify Table 7 (expression cassettes — text-heavy)
- [ ] Rebuild or verify Table 8 (expression data) with table footnotes
- [ ] Verify Table 6 (regulatory references) imported cleanly
- [ ] Add Table_FootNote paragraphs below tables with superscript footnotes

**Automated formatting verification:**
- [ ] Verify GREP auto-formatting: scientific names italicized correctly
- [ ] Verify GREP auto-formatting: gene names (lowercase italic) vs. protein names (roman)
- [ ] Verify Char_No Break on event names, units, regulatory abbreviations
- [ ] Verify parenthetical citations not orphaned at line breaks
- [ ] Check multiplication signs normalized in stacked event names

**Layout review:**
- [ ] Check and fix widow/orphan issues (especially at column and page breaks)
- [ ] Verify table pages have column overrides applied (single column for wide tables)
- [ ] Update running header text if it includes the protein name
- [ ] Update title page: protein name, date, keywords
- [ ] Update copyright year and any boilerplate text

**QC flags:**
- [ ] Check for empty/placeholder footnotes
- [ ] Verify table numbering is sequential (Tables 1–N)
- [ ] Verify body text cross-references to tables are correct
- [ ] Check that "Table 6" references in body text point to the right table (numbering may shift between proteins)

**Final:**
- [ ] Preflight: check for overset text, missing fonts, RGB colors
- [ ] Export PDF (Press Quality or High Quality Print)
- [ ] Export PDF (Smallest File Size) for web posting on foodsystems.org

---

## 14. Notes for Our Collaborative Next Steps

Once you're ready to start building, we can work together on:

1. **Detailed master page wireframes** — I can create visual mockups or precise specifications for each master page layout with exact frame coordinates
2. **GREP pattern testing** — I now have the actual draft content and can refine the regex patterns against real text
3. **Table reconstruction strategy** — With 8 tables across 3 types, we should discuss whether to build approval matrices as InDesign native tables, linked Excel, or a hybrid approach. Table 5 (21 country columns) is the critical test case.
4. **Style-by-style walkthrough** — I can guide you through creating each paragraph and character style with exact InDesign dialog settings
5. **Word template for colleagues** — Optionally, I can help create a Word template with matching style names so future monograph files import more cleanly (especially for heading levels and crop sub-section formatting)
6. **Automation scripts** — If you want to go further, InDesign supports scripting (JavaScript/ExtendScript) for batch operations like placing content, updating text variables, or running find/change sequences
7. **Table 5 landscape test** — We should prototype the widest table (Table 5) early to determine if landscape rotation is necessary or if portrait full-width with rotated headers and 6.5pt text is sufficient

---

## Appendix A: Complete GREP Pattern Reference

For quick copy-paste when building paragraph styles. All patterns below should be added to the GREP Style tab of the relevant paragraph styles.

### Apply to Body_text, Body_BulletL1, Body_Footnote:

| GREP Pattern | Character Style | Purpose |
|-------------|----------------|---------|
| `((?:Dr\|Mr\|Mrs\|Ms\|Prof\|St)\.\s\S+)` | Char_No Break | Keep titles with names |
| `(\d+\.?\d*\s?(?:mg\|µg\|ug\|ng\|g\|kg\|kDa\|Da\|mL\|µL\|L\|bp\|kb\|days?\|hours?\|min\|sec\|%))` | Char_No Break | Number + unit |
| `(\d+\.?\d*\s?(?:ug\|µg\|mg)/(?:g\|kg)\s(?:fresh\sweight\|body\s?weight\|dry\sweight\|bw\|fw\|dw))` | Char_No Break | Compound units |
| `((?:Table\|Tables\|Figure\|Fig\.\|Appendix\|Section\|Event\|Line)\s\d+(?:[-–]\d+)?)` | Char_No Break | Cross-references |
| `(Bacillus thuringiensis\|B\. thuringiensis\|Zea mays\|Z\. mays\|Gossypium hirsutum\|G\. hirsutum\|Oryza sativa\|O\. sativa\|Manduca sexta\|M\. sexta\|Vigna unguiculata\|V\. unguiculata\|Eucalyptus\ssp\.\|Saccharum\ssp\.\|Arabidopsis thaliana\|A\. thaliana)` | Char_Italic | Scientific names |
| `((?:subsp\.\|var\.)\s\w+)` | Char_Italic | Subspecies/variety |
| `(?<!\w)(cry\d[A-Z][a-z]\d?\|cry\d[A-Z]{2}\d?\|vip\d[A-Z][a-z]?\d?\|pat\|bar\|epsps\|cp4[\s-]epsps)(?!\w)` | Char_GeneItalic | Gene names |
| `(Cry1Ab\|Cry1Ac\|Cry1Bb\|Cry1F[a]?\|Cry2Ab\|Cry2Aa\|Cry3Bb1\|Vip3Aa\|EPSPS\|PAT/BAR)` | Char_No Break | Protein names |
| `(MON810\|MON801\|MON802\|MON809\|MON863\|MON88017\|MON89034\|MON87427\|MON87411\|MON87419\|Bt11\|BT176\|Bt176\|Bt10\|COT67B\|NK603\|GA21\|MIR604\|MIR162\|TC1507\|T304-40?\|DBN9936\|DBN9336\|GHB119\|GHB614\|GHB811\|COT102\|MON88701\|AAT709A\|LP007-1\|LP026-2\|CTC175-?A\|CTC20BT\|1521K059)` | Char_No Break | Event names |
| `(CTNBio\|FSANZ\|FZANS\|EFSA\|USEPA\|USFDA\|US\sFDA\|CFIA\|J-BCH\|NBMA\|CSIR\|NBA\|USDA\|APHIS\|OECD\|ISAAA\|FAO\|WHO)` | Char_No Break | Regulatory bodies |
| `(Bacillus\sthuringiensis\|B\.\sthuringiensis)` | Char_No Break | Keep Bt together |

### Apply to Head_Section only (Bold italic):

| GREP Pattern | Character Style | Purpose |
|-------------|----------------|---------|
| `(Bacillus thuringiensis\|B\. thuringiensis\|Zea mays\|Z\. mays\|Gossypium hirsutum\|G\. hirsutum\|Oryza sativa\|O\. sativa\|Manduca sexta\|M\. sexta\|Vigna unguiculata\|V\. unguiculata\|Eucalyptus\ssp\.\|Saccharum\ssp\.\|Arabidopsis thaliana\|A\. thaliana)` | Char_BoldItalic | Scientific names |
| `((?:subsp\.\|var\.)\s\w+)` | Char_BoldItalic | Subspecies/variety |
| `(?<!\w)(cry\d[A-Z][a-z]\d?\|cry\d[A-Z]{2}\d?\|vip\d[A-Z][a-z]?\d?\|pat\|bar\|epsps\|cp4[\s-]epsps)(?!\w)` | Char_BoldItalic | Gene names |

### Apply to Head_SubsectionNumbered, Head_SubsectionUnnumbered, Head_SubSubSectionUnnumbered, Head_CropName (Semibold italic):

| GREP Pattern | Character Style | Purpose |
|-------------|----------------|---------|
| `(Bacillus thuringiensis\|B\. thuringiensis\|Zea mays\|Z\. mays\|Gossypium hirsutum\|G\. hirsutum\|Oryza sativa\|O\. sativa\|Manduca sexta\|M\. sexta\|Vigna unguiculata\|V\. unguiculata\|Eucalyptus\ssp\.\|Saccharum\ssp\.\|Arabidopsis thaliana\|A\. thaliana)` | Char_SemiboldItalic | Scientific names |
| `((?:subsp\.\|var\.)\s\w+)` | Char_SemiboldItalic | Subspecies/variety |
| `(?<!\w)(cry\d[A-Z][a-z]\d?\|cry\d[A-Z]{2}\d?\|vip\d[A-Z][a-z]?\d?\|pat\|bar\|epsps\|cp4[\s-]epsps)(?!\w)` | Char_SemiboldItalic | Gene names |

### Apply to Ref_Entry only:

| GREP Pattern | Character Style | Purpose |
|-------------|----------------|---------|
| `(https?://\S+)` | Char_RefURL | URL styling |
| `((?:Dr\|Mr\|Mrs\|Ms\|Prof\|St)\.\s\S+)` | Char_No Break | Keep titles with names |

### Apply to Table_Heading only:

| GREP Pattern | Character Style | Purpose |
|-------------|----------------|---------|
| `Table \d+:` | Char_Semibold | Auto-semibold on "Table N:" prefix |

### Apply to Table_FootNote only:

| GREP Pattern | Character Style | Purpose |
|-------------|----------------|---------|
| `^.` | Char_Superscript | Auto-superscript the first character (footnote letter or number) |

### Apply to Body_Footnote only:

| GREP Pattern | Character Style | Purpose |
|-------------|----------------|---------|
| `^\d+(?=—)` | Char_Superscript | Auto-superscript footnote number before em dash |

---

## Appendix B: Version History

| Version | Date | Changes |
|---------|------|---------|
| v1 | Feb 2026 | Initial template guide based on original Cry1Ab monograph PDF |
| v2 | Feb 2026 | Updated based on analysis of Cry1Ab_FFS_27JAN2026_18FEB2026.docx draft. Major changes: 5 approval tables (not 1), 8 total tables (not 2), expanded heading hierarchy, new paragraph styles (Head_SubsectionUnnumbered, Head_CropName, Head_RunIn, Table_BodyCenter_X, Table_CaptionSub, Table_BodyDescriptive, Table_FootNote), expanded GREP patterns for species/genes/events/units, updated table styling strategy with 3 table types, revised master page sequence, enhanced post-import cleanup workflow, updated per-monograph checklist |
| v3 | Feb 2026 | Simplified master pages from 4 (A, B, C, D) to 2 (A and B). Master C (full-width) and Master D (references) eliminated — full-width table pages are handled by overriding Master B column settings or placing tables as standalone objects; reference page sizing is controlled entirely by the Ref_Entry paragraph style. Corrected column gutter from 0.25" to 0.1667" (1 pica, InDesign default) to maximize column width for better justified text. |
| v4 | Mar 2026 | Replaced all references to "hanging indent" with InDesign-native terminology: Left Indent + First Line Left Indent (negative value). Clarifies that "hanging: 0.25\"" means Left Indent: 0.25\", First Line Left Indent: -0.25\". |
| v5 | Mar 2026 | Removed misleading "GREP styles for italic journal names" note from Ref_Entry. Journal name italicization is not automatable via GREP and should be preserved from Word import or applied manually. |
| v6 | Mar 2026 | Changed margins from asymmetric (0.75"/0.875") to symmetric 0.75" all sides. Removed false binding gutter rationale. Added scientific journal aesthetic rationale: generous symmetric margins create a compact text block that signals peer-reviewed publication. Updated text area calculation to 7.0" wide, columns to ~3.417" each. |
| v7 | Mar 2026 | Added Section 2: Post-Placement Cleanup Checklist for the immediate steps after placing a Word file (Find/Change cleanup, heading remapping, GREP verification, QC flags). All subsequent sections renumbered. |
| v8 | Mar 2026 | Replaced Section 9.6 summary list with detailed step-by-step instructions for creating table and cell styles, including conceptual explanation of how table styles, cell styles, and paragraph styles nest together, the (Ignore) vs None distinction for strokes, and the override precedence hierarchy. Added override-clearing step to Section 2 Post-Placement Cleanup Checklist. |
| v9 | Mar 2026 | Updated base table style from Black to AFSI Gray throughout (table border, row strokes, column strokes, header fill). Table border restored to 1pt. Added CStyle_HeaderLeft and CStyle_HeaderRight cell styles to keep header edge strokes flush with the 1pt table border while allowing 0.25pt AFSI Light Gray interior column dividers in the header row. Step-by-step instructions now reflect the three-header-style approach. Added corresponding Green edge variants (CStyle_HeaderGreenLeft, CStyle_HeaderGreenRight) to the additional styles table. |
| v10 | Mar 2026 | Added full step-by-step build for TStyle_Approvals with rotated header cell styles (CStyle_HeaderRotated, CStyle_HeaderRotatedLeft, CStyle_HeaderRotatedRight) and CStyle_BodyApproval. Expanded Step 9 of the base table workflow into a detailed override-clearing procedure for imported tables, explaining the Alt/Opt-click sequence and why Word formatting blocks styles. |
| v11 | Mar 2026 | Added CStyle_BodyBottom (for TStyle_Simple) and CStyle_BodyApprovalBottom (for TStyle_Approvals) to handle 1pt bottom border flush with the table border, matching the edge header cell style pattern. Updated Step 9 in Section 9.7 and test steps in Section 9.8 to include bottom row application. Added step 11 to the import workflow for bottom row styles. Also: No Break override stripping step, body text size/spacing normalization via Find/Change, Document Footnote Options checklist, single-column-first workflow note, Find/Change field clarifications, Body_text space after changed to 4pt, Head_Section changed to AFSI Blue with 8pt/4pt spacing, Table_Heading paragraph style with Span Columns and GREP auto-semibold on "Table N:", Char_Semibold character style, Table_Span utility paragraph style, CStyle_Highlighted changed from Mustard to AFSI Blue 25% tint. |
| v12 | Mar 2026 | Major workflow overhaul. Replaced the direct Word-to-InDesign placement workflow with a three-phase approach to eliminate crashes caused by hidden Word XML formatting. Phase 1: prep in Word (mark superscripts with {{ }} markers via Find/Replace, delete tables leaving [INSERT TABLE X HERE] placeholders, save as UTF-8 plain text). Phase 2: place clean .txt file in single-column InDesign, apply all paragraph styles manually, run Find/Change operations, restore superscript markers with automatic Char_Superscript application, recover manual italics from side-by-side Word reference, handle footnotes via hybrid approach (native InDesign footnotes for first occurrences, Char_Superscript for repeated references). Phase 3: paste tables one at a time from original Word file, run override-clearing sequence per table, apply table/cell styles, switch to two columns for final layout pass. Rewrote Section 10 to match (10.2 Text Preparation in Word, 10.3 Plain Text Placement, 10.4 Table Paste-In). Removed obsolete import mapping table and No Break/size normalization Find/Change steps (no longer needed with clean text). |
