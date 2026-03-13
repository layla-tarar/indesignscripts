// CleanUp.jsx
// Run on the active document containing the placed _clean.docx content.
//
// REQUIRED SCRIPT ORDER:
//   1. TableStyler.jsx      — resets and styles all tables (clears cell overrides)
//   2. CleanUp.jsx          — text cleanup + superscript restoration (this script)
//   3. InsertFootnotes.jsx  — replaces {{fn:N}} markers with native footnotes
//
// CleanUp must run AFTER TableStyler (so cleared cell overrides are not re-applied)
// and BEFORE InsertFootnotes (so native footnote text objects do not exist when we
// run doc-wide text operations — their presence causes InDesign's Text engine to crash).
//
// CleanUp's Step 6 regex (\{\{\d+\}\}) does NOT match {{fn:N}} markers (the colon
// prevents a match), so the footnote markers are safely left for InsertFootnotes.
//
// The following are already handled by clean_docx.py (before placement):
//   - Bullet character stripping (U+2022)
//   - Tilde operator (U+223C) → standard tilde
//   - "Table N." → "Table N:" in table description rows
//   - Superscript/footnote markers wrapped as {{...}}
//   - Character style overrides stripped from runs
//
// This script handles what remains:
//   Pre-0.  Remap imported Word paragraph styles → InDesign template styles
//           (done by direct paragraph iteration, NOT doc.changeGrep, to avoid
//            InDesign's Text Walker which crashes on certain document states)
//   Pre-0b. Remove unnamed/blank imported character styles (empty name or "Unnamed Style,...")
//   0.  Clear local formatting overrides (via paragraph iteration, not GREP)
//   1.  Double spaces → single space
//   2.  Consecutive paragraph returns → single return
//   3.  Multiplication sign normalization (digit x Capital → digit × Capital)
//   4.  Protein → gene name correction: "Cry1Ab gene" → "cry1Ab gene", etc.
//   5.  "Table N." → "Table N:" in Table_Caption and Table_Heading styled paragraphs
//   6.  Superscript {{N}} markers → apply Char_Superscript character style

var doc = app.activeDocument;
var report = [];

app.findGrepPreferences = NothingEnum.nothing;
app.changeGrepPreferences = NothingEnum.nothing;

// GREP helper for specific-pattern find/change (steps 1–6 only).
// NOTE: This is NOT used for any ".+" broad pattern — those are handled by
// paragraph iteration below to avoid InDesign's Text Walker crash.
function doGrepChange(findWhat, changeTo, findParaStyle, changeCharStyle) {
    app.findGrepPreferences = NothingEnum.nothing;
    app.changeGrepPreferences = NothingEnum.nothing;

    app.findGrepPreferences.findWhat = findWhat;
    app.changeGrepPreferences.changeTo = changeTo;

    if (findParaStyle) {
        try {
            app.findGrepPreferences.appliedParagraphStyle =
                doc.paragraphStyles.itemByName(findParaStyle);
        } catch (e) { /* style not found — search all */ }
    }
    if (changeCharStyle) {
        app.changeGrepPreferences.appliedCharacterStyle =
            doc.characterStyles.itemByName(changeCharStyle);
    }

    var results = doc.changeGrep();
    app.findGrepPreferences = NothingEnum.nothing;
    app.changeGrepPreferences = NothingEnum.nothing;
    return results.length;
}

// --- Pre-0 + Fallback + Step 0: Single paragraph iteration pass ---
//
// We combine three operations into one textFrame paragraph loop to avoid
// doc.changeGrep() with ".+" patterns, which trigger InDesign's Text Walker
// on every text object in the document and crash (EXC_BAD_ACCESS) when any
// of those objects are in an unstable state (e.g., after TableStyler restructures
// header rows).
//
// Operations per paragraph:
//   A. Remap Word paragraph style → InDesign template style (Pre-0)
//   B. Remap [Basic Paragraph] → Body_Text (Fallback)
//   C. Clear local formatting overrides — clears both paragraph-level formatting
//      AND character style overrides applied to runs within the paragraph (Step 0)
//
// We iterate doc.textFrames only (not doc.stories) to exclude footnote text,
// table cells, and master page items from this pass.

var styleRemap = [
    { from: "Normal",                   to: "Body_Text" },
    { from: "Body Text",                to: "Body_Text" },
    { from: "Default",                  to: "Body_Text" },
    { from: "Default Paragraph Font",   to: "Body_Text" },
    { from: "Heading 1",                to: "Head_Section" },
    { from: "Heading 2",                to: "Head_SubsectionUnnumbered" },
    { from: "Heading 3",                to: "Head_SubSubSectionUnnumbered" },
    { from: "Title",                    to: "Title_Title" },
    { from: "List Paragraph",           to: "Body_BulletL1" },
    { from: "Table_Header",             to: "Table_Heading" },
];

// Build a lookup map: fromStyleName → toStyle object
var styleMap = {};
for (var r = 0; r < styleRemap.length; r++) {
    try {
        var toStyle = doc.paragraphStyles.itemByName(styleRemap[r].to);
        if (toStyle.isValid) styleMap[styleRemap[r].from] = toStyle;
    } catch(e) {}
}
var bodyTextStyle = doc.paragraphStyles.itemByName("Body_Text");

var remapCount = 0;
var clearOverrideCount = 0;
var seenStories = {};

var frames = doc.textFrames.everyItem().getElements();
for (var s = 0; s < frames.length; s++) {
    try {
        // Track stories so threaded frames don't process the same story twice
        var sid = frames[s].parentStory.id;
        if (seenStories[sid]) continue;
        seenStories[sid] = true;

        var paras = frames[s].paragraphs.everyItem().getElements();
        for (var p = 0; p < paras.length; p++) {
            // A/B. Remap paragraph style
            try {
                var styleName = paras[p].appliedParagraphStyle.name;
                if (styleMap[styleName]) {
                    paras[p].appliedParagraphStyle = styleMap[styleName];
                    remapCount++;
                } else if (styleName === "[Basic Paragraph]" && bodyTextStyle.isValid) {
                    paras[p].appliedParagraphStyle = bodyTextStyle;
                    remapCount++;
                }
            } catch(e) {}

            // C. Clear local overrides (paragraph formatting + character style overrides)
            try {
                paras[p].clearOverrides(OverrideType.LOCAL_ONLY);
                clearOverrideCount++;
            } catch(e) {}
        }
    } catch(e) {}
}
report.push("Paragraph styles remapped: " + remapCount);
report.push("Local overrides cleared: " + clearOverrideCount + " paragraphs");

// Clean up the now-unused imported Word styles from the style list.
// This is best-effort — if a style still has usages (e.g., in table cells
// or footnotes we skipped), the remove() will fail silently.
for (var r = 0; r < styleRemap.length; r++) {
    try {
        var fromStyle = doc.paragraphStyles.itemByName(styleRemap[r].from);
        var toStyleForRemove = doc.paragraphStyles.itemByName(styleRemap[r].to);
        if (fromStyle.isValid && toStyleForRemove.isValid) {
            fromStyle.remove(toStyleForRemove);
        }
    } catch(e) {}
}

// --- Pre-0b. Remove unnamed/blank imported character styles ---
// Word imports often bring in character styles with empty names or names like
// "Unnamed Style, ..." — delete them. Any usages were already cleared by
// clearOverrides(LOCAL_ONLY) in the paragraph loop above.
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
report.push("Unnamed/blank character styles removed: " + unnamedCharCount);

// --- 1. Double spaces → single space ---
var count = doGrepChange(" {2,}", " ");
report.push("Double spaces fixed: " + count);

// --- 2. Consecutive paragraph returns → single return ---
count = doGrepChange("\\r{2,}", "\\r");
report.push("Extra returns fixed: " + count);

// --- 3. Multiplication sign normalization ---
count = doGrepChange("(?<=\\d) x (?=[A-Z])", " \u00D7 ");
report.push("Multiplication signs normalized: " + count);

// --- 4. Protein → gene name correction ---
var proteinToGene = [
    ["Cry1Ab",  "cry1Ab"],  ["Cry1Ac",  "cry1Ac"],
    ["Cry1Bb",  "cry1Bb"],  ["Cry1F",   "cry1F"],
    ["Cry1Fa",  "cry1Fa"],  ["Cry2Ab",  "cry2Ab"],
    ["Cry2Aa",  "cry2Aa"],  ["Cry3Bb1", "cry3Bb1"],
    ["Vip3Aa",  "vip3Aa"],
    ["CP4-EPSPS", "cp4-epsps"], ["CP4 EPSPS", "cp4 epsps"], ["EPSPS", "epsps"],
    ["PAT", "pat"], ["BAR", "bar"],
];
var geneNameCount = 0;
for (var g = 0; g < proteinToGene.length; g++) {
    app.findGrepPreferences = NothingEnum.nothing;
    app.changeGrepPreferences = NothingEnum.nothing;
    app.findGrepPreferences.findWhat = "\\b" + proteinToGene[g][0] + "(?=\\s+gene\\b)";
    app.changeGrepPreferences.changeTo = proteinToGene[g][1];
    geneNameCount += doc.changeGrep().length;
}
app.findGrepPreferences = NothingEnum.nothing;
app.changeGrepPreferences = NothingEnum.nothing;
report.push("Protein \u2192 gene name corrections: " + geneNameCount);

// --- 5. "Table N." → "Table N:" in caption/heading paragraphs ---
var captionStyles = ["Table_Caption", "Table_Heading"];
var captionCount = 0;
for (var cs = 0; cs < captionStyles.length; cs++) {
    try {
        captionCount += doGrepChange("(Table \\d+)\\.", "$1:", captionStyles[cs]);
    } catch (e) {}
}
report.push("Table caption periods \u2192 colons: " + captionCount);

// --- 6. Superscript {{N}} and {{letter}} markers → Char_Superscript ---
// Pattern \{\{\d+\}\} matches {{1}}, {{2}} etc.
// Pattern \{\{[a-zA-Z]\}\} matches {{a}}, {{b}} etc.
// Neither matches {{fn:N}} — the colon in fn:N prevents a match — so footnote
// markers are safely ignored and left for InsertFootnotes.jsx to process.
try {
    var supCount = 0;
    supCount += doGrepChange("\\{\\{(\\d+)\\}\\}", "$1", null, "Char_Superscript");
    supCount += doGrepChange("\\{\\{([a-zA-Z])\\}\\}", "$1", null, "Char_Superscript");
    report.push("Superscript markers restored: " + supCount);
} catch (e) {
    report.push("Superscript fix skipped \u2014 'Char_Superscript' character style not found.");
}

// --- Report ---
alert(
    "Cleanup complete!\n\n" +
    report.join("\n") +
    "\n\nScript order: TableStyler \u2192 CleanUp \u2192 InsertFootnotes"
);
