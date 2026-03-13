# Handoff Notes — March 13 Session
## Purpose
These notes describe every **feature/behaviour change** made during the March 13 session.
The scripts in this folder (`*_session_mar13.*`) are the end-of-session versions.
The base (working) version is git commit `ef7f423`.

The goal is: add features one at a time to the working base, test after each addition,
and isolate which change introduced the InDesign crash (EXC_BAD_ACCESS / SIGSEGV in
InDesign's Text engine, Thread 0, Text Walker → CScriptProvider::HandleMethodOnObjects).

---

## Changes to test (add back one at a time)

### 1. `TitleCaseHeadings.jsx` — Add `Title_Title` to style list
**What it does:** Adds `"Title_Title"` as the first entry in the `stylesToFix` array,
so the document title paragraph also gets Title Case applied.
**Risk level:** Low — this is a trivial one-line array addition.
**How to add back:** Add `"Title_Title"` to the top of the `stylesToFix` array.

---

### 2. `TableStyler.jsx` — Increase rotated header row height multiplier
**What it does:** In the `calcRotatedTextHeight` function, increases the character-width
multiplier used to estimate row height for rotated header text from `0.50` to `0.58`.
This was needed because "New Zealand" (11 chars) was getting cut off in approval table
country headers.
**Risk level:** Low — changes only a numeric constant affecting row height calculation.
**How to add back:** Find the comment about `avg Latin char width` and change `0.50` to `0.58`.

---

### 3. `CleanUp.jsx` — Remove unnamed/blank imported character styles (Pre-0b)
**What it does:** Before the main cleanup steps, iterates all character styles in the
document and deletes any whose name is `""` (empty) or starts with `"Unnamed Style,"`.
Any usages of those styles are replaced with `[None]` when the style is removed.
**Why needed:** Word import brings in unnamed/blank character styles that pollute the
InDesign document and can cause unexpected formatting.
**Risk level:** Medium — involves calling `characterStyle.remove(noneStyle)` which
walks text. This COULD be the crash source — test carefully.
**How to add back:**
```javascript
var noneCharStyle = doc.characterStyles.itemByName("[None]");
var allCharStyles = doc.characterStyles.everyItem().getElements();
var unnamedCharCount = 0;
for (var u = 0; u < allCharStyles.length; u++) {
    var cStyle = allCharStyles[u];
    var cName = cStyle.name;
    if (cName === "" || cName.indexOf("Unnamed Style,") === 0) {
        try {
            cStyle.remove(noneCharStyle);
            unnamedCharCount++;
        } catch (e) {}
    }
}
// report: "Unnamed/blank character styles removed: " + unnamedCharCount
```

---

### 4. `CleanUp.jsx` — Protein → gene name correction (Step 4)
**What it does:** When a capitalised protein name (e.g., "Cry1Ab") is immediately
followed by the word "gene", replaces it with the lowercase gene name convention
(e.g., "cry1Ab gene"). Uses GREP lookahead `\bProteinName(?=\s+gene\b)`.
**Proteins covered:**
- Bt series: Cry1Ab, Cry1Ac, Cry1Bb, Cry1F, Cry1Fa, Cry2Ab, Cry2Aa, Cry3Bb1, Vip3Aa
- EPSPS series: CP4-EPSPS, CP4 EPSPS, EPSPS
- PAT/BAR: PAT, BAR
**Risk level:** Low — uses targeted GREP patterns with lookahead, doesn't use `.+`.
**How to add back:** Add a `proteinToGene` array and loop running `doGrepChange` for
each pair with pattern `"\\b" + protein + "(?=\\s+gene\\b)"`.

---

### 5. `CleanUp.jsx` — Table_Header → Table_Heading legacy remap (Pre-0)
**What it does:** Adds `{ from: "Table_Header", to: "Table_Heading" }` to the
`styleRemap` array in Pre-0. This handles documents processed by the old version of
`clean_docx.py` that stamped `Table_Header` (a cell style name) instead of
`Table_Heading` (the paragraph style name for table title lines).
**Risk level:** Medium — the style remap uses text-walking operations. See notes
on the crash investigation below.
**How to add back:** Add `{ from: "Table_Header", to: "Table_Heading" }` to the
`styleRemap` array.

---

### 6. `clean_docx.py` — Fix Table_Header → Table_Heading (4 occurrences)
**What it does:** Corrects `clean_docx.py` to use `"Table_Heading"` (the InDesign
paragraph style) instead of `"Table_Header"` (an InDesign cell style) in all places:
- `_ASSIGNED_STYLES` set
- `_extract_description_rows` function
- `_infer_paragraph_styles` docstring and logic
**Risk level:** None for InDesign — this is a Python-only fix. No crash risk.
**How to add back:** Replace all 4 occurrences of `"Table_Header"` with `"Table_Heading"`
in clean_docx.py (use replace_all).

---

### 7. `clean_docx.py` — Insert empty paragraph after every table
**What it does:** Adds a new function `_insert_paragraph_after_tables()` that inserts
a blank `<w:p>` element after every `<w:tbl>` in the Word document body.
**Why needed:** When InDesign places the Word file, it anchors tables in the following
paragraph. If that paragraph immediately has body text, CleanUp.jsx's `\r{2,} → \r`
collapse (or some other mechanism) was causing those paragraphs to receive `Table_Span`
style. The guaranteed empty paragraph ensures a clean separation.
**Risk level:** None for InDesign — Python-only change. No crash risk.
**How to add back:**
```python
def _insert_paragraph_after_tables(doc: Document) -> None:
    body = doc.element.body
    tbl_tag = qn("w:tbl")
    for child in list(body):
        if child.tag != tbl_tag:
            continue
        empty_p = OxmlElement("w:p")
        child.addnext(empty_p)
```
Call it in `main()` after `_extract_description_rows` and before `_register_assigned_styles`.

---

## Crash investigation summary

The crash is: `EXC_BAD_ACCESS (SIGSEGV)` in InDesign 2026 (21.2.0.30) on macOS 26.3.1.
Thread 0 crash trace: `Text module → Text Walker → CScriptProvider::HandleMethodOnObjects`
This is InDesign's GREP find/change engine (Text Walker) crashing inside the Text module.

Three attempts were made to fix it without success:
1. Changed `doc.stories` iteration → `doc.textFrames` iteration for `clearOverrides`
2. Added `app.findChangeTextOptions.includeFootnotes = false` before all GREP operations
3. Replaced all `doc.changeGrep()` with `.+` pattern → direct paragraph iteration,
   AND changed script order to run CleanUp BEFORE InsertFootnotes

None resolved the crash. TableStyler also started crashing at the end (force quit required),
suggesting the document or InDesign session itself may be in a bad state.

**Recommended next-session approach:**
- Start fresh: revert all scripts to the working `ef7f423` base
- Re-run the full workflow (TableStyler → CleanUp → InsertFootnotes) with the base scripts
  to confirm they still work
- Add features back ONE AT A TIME in the order listed above, testing after each
- The most likely crash candidates are items 3 and 5 (both involve text-walking operations)
- Check whether the crash reproduces WITHOUT InsertFootnotes having run (run TableStyler
  then CleanUp only) — this would rule out the footnote hypothesis

**Files saved for reference:**
- `CleanUp_session_mar13_crashing.jsx` — final CleanUp.jsx from this session
- `TableStyler_session_mar13.jsx` — TableStyler.jsx with 0.58 multiplier
- `TitleCaseHeadings_session_mar13.jsx` — TitleCaseHeadings.jsx with Title_Title
- `clean_docx_session_mar13.py` — clean_docx.py with all Python fixes
