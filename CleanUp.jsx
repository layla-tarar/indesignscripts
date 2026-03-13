// CleanUp.jsx
// Run on the active document containing the placed _clean.docx content.
//
// REQUIRED SCRIPT ORDER (all scripts operate on the same _clean.docx placement):
//   1. TableStyler.jsx      — resets and styles all tables (clears cell overrides)
//   2. InsertFootnotes.jsx  — replaces {{fn:N}} markers with native footnotes
//   3. CleanUp.jsx          — text cleanup + superscript restoration (this script)
//      Step 5 (Char_Superscript) must run AFTER TableStyler, which clears cell
//      character overrides, and AFTER InsertFootnotes, which removes {{fn:N}} markers.
//
// The following are already handled by clean_docx.py (before placement):
//   - Bullet character stripping (U+2022)
//   - Tilde operator (U+223C) → standard tilde
//   - "Table N." → "Table N:" in table description rows
//   - Superscript/footnote markers wrapped as {{...}}
//   - Character style overrides stripped from runs
//
// This script handles what remains:
//   1. Double spaces → single space
//   2. Consecutive paragraph returns → single return
//   3. Multiplication sign normalization (digit x Capital → digit × Capital)
//   4. "Table N." → "Table N:" in Table_Caption and Table_Heading styled paragraphs
//      (standalone caption lines in the body text, separate from description rows)
//   5. Superscript {{N}} markers → apply Char_Superscript character style
//      Handles plain {{N}}, {{letter}}, and in-cell markers like {{a}}, {{b}}.
//      Does NOT match {{fn:N}} or {{en:N}} — those are native footnote refs.

var doc = app.activeDocument;
var report = [];

app.findGrepPreferences = NothingEnum.nothing;
app.changeGrepPreferences = NothingEnum.nothing;

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
    return results.length;
}

// --- Pre-0. Remap imported Word paragraph styles → InDesign template styles ---
// paragraphStyle.remove(replacement) reassigns all uses AND deletes the imported style.
// Run before override clearing so overrides are wiped against the correct styles.
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
];
var remapCount = 0;
for (var r = 0; r < styleRemap.length; r++) {
    try {
        var fromStyle = doc.paragraphStyles.itemByName(styleRemap[r].from);
        var toStyle   = doc.paragraphStyles.itemByName(styleRemap[r].to);
        if (fromStyle.isValid && toStyle.isValid) {
            fromStyle.remove(toStyle);
            remapCount++;
        }
    } catch (e) {}
}
report.push("Word paragraph styles remapped + deleted: " + remapCount);

// Targeted remap: Table_Header → Table_Heading (paragraphs only, not child styles)
// Uses paragraph iteration instead of paragraphStyle.remove() to avoid reparenting
// any styles that are based on Table_Header.
try {
    var tableHeaderStyle  = doc.paragraphStyles.itemByName("Table_Header");
    var tableHeadingStyle = doc.paragraphStyles.itemByName("Table_Heading");
    if (tableHeaderStyle.isValid && tableHeadingStyle.isValid) {
        var thStories = doc.stories.everyItem().getElements();
        var tableHeaderCount = 0;
        for (var ths = 0; ths < thStories.length; ths++) {
            var thParas = thStories[ths].paragraphs.everyItem().getElements();
            for (var thp = 0; thp < thParas.length; thp++) {
                try {
                    if (thParas[thp].appliedParagraphStyle.name === "Table_Header") {
                        thParas[thp].appliedParagraphStyle = tableHeadingStyle;
                        tableHeaderCount++;
                    }
                } catch (e) {}
            }
        }
        if (tableHeaderCount > 0) report.push("Table_Header \u2192 Table_Heading: " + tableHeaderCount);
    }
} catch (e) {}

// Fallback: apply Body_Text to any paragraphs still on [Basic Paragraph]
// (InDesign assigns [Basic Paragraph] when a placed paragraph has no matching style.)
try {
    app.findGrepPreferences = NothingEnum.nothing;
    app.changeGrepPreferences = NothingEnum.nothing;
    app.findGrepPreferences.findWhat = ".+";
    app.findGrepPreferences.appliedParagraphStyle =
        doc.paragraphStyles.itemByName("[Basic Paragraph]");
    app.changeGrepPreferences.appliedParagraphStyle =
        doc.paragraphStyles.itemByName("Body_Text");
    var fallbackCount = doc.changeGrep().length;
    if (fallbackCount > 0) report.push("Unstyled → Body_Text: " + fallbackCount);
    app.findGrepPreferences = NothingEnum.nothing;
    app.changeGrepPreferences = NothingEnum.nothing;
} catch (e) {}

// --- Pre-0b. Remove unnamed/blank imported character styles ---
// Word import brings in unnamed or "Unnamed Style,N" character styles that pollute
// the document. Remove them before the override-clearing pass, replacing any usages
// with [None] so no text is left referencing a deleted style.
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

// --- 0. Clear character style overrides and local formatting overrides ---
// Strips character styles and direct run formatting imported from Word
// (fonts, sizes, colors, bold, underline, etc.) so InDesign paragraph styles
// fully control appearance. Step 5 re-applies Char_Superscript from {{...}} markers.
app.findGrepPreferences = NothingEnum.nothing;
app.changeGrepPreferences = NothingEnum.nothing;
app.findGrepPreferences.findWhat = ".+";
app.changeGrepPreferences.appliedCharacterStyle = doc.characterStyles.itemByName("[None]");
var clearCharCount = doc.changeGrep().length;
app.findGrepPreferences = NothingEnum.nothing;
app.changeGrepPreferences = NothingEnum.nothing;

// Clear all local text overrides (font, size, color, etc.) from every paragraph
var stories = doc.stories.everyItem().getElements();
for (var s = 0; s < stories.length; s++) {
    var paras = stories[s].paragraphs.everyItem().getElements();
    for (var p = 0; p < paras.length; p++) {
        try { paras[p].clearOverrides(OverrideType.LOCAL_ONLY); } catch(e) {}
    }
}
report.push("Character + local formatting overrides cleared: " + clearCharCount + " ranges");

// --- 1. Double spaces → single space ---
var count = doGrepChange(" {2,}", " ");
report.push("Double spaces fixed: " + count);

// --- 2. Consecutive paragraph returns → single return ---
count = doGrepChange("\\r{2,}", "\\r");
report.push("Extra returns fixed: " + count);

// --- 3. Multiplication sign normalization ---
// Matches: digit + space + x + space + uppercase letter
// Replaces with: digit + space + × (U+00D7) + space + uppercase letter
count = doGrepChange("(?<=\\d) x (?=[A-Z])", " \u00D7 ");
report.push("Multiplication signs normalized: " + count);

// --- 4. "Table N." → "Table N:" in caption/heading paragraphs ---
// Handles standalone Table_Caption and Table_Heading paragraphs in the body text.
// (Description rows inside tables are already handled by extract_tables.py.)
var captionStyles = ["Table_Caption", "Table_Heading"];
var captionCount = 0;
for (var cs = 0; cs < captionStyles.length; cs++) {
    try {
        captionCount += doGrepChange("(Table \\d+)\\.", "$1:", captionStyles[cs]);
    } catch (e) {}
}
report.push("Table caption periods → colons: " + captionCount);

// --- 4b. Protein name → gene name when followed by "gene" ---
// e.g. "Cry1Ab gene" → "cry1Ab gene", "PAT gene" → "pat gene"
// Uses lookahead so only the protein name is replaced, not the word "gene".
var proteinToGene = [
    ["Cry1Ab",    "cry1Ab"],
    ["Cry1Ac",    "cry1Ac"],
    ["Cry1Bb",    "cry1Bb"],
    ["Cry1F",     "cry1F"],
    ["Cry1Fa",    "cry1Fa"],
    ["Cry2Ab",    "cry2Ab"],
    ["Cry2Aa",    "cry2Aa"],
    ["Cry3Bb1",   "cry3Bb1"],
    ["Vip3Aa",    "vip3Aa"],
    ["CP4-EPSPS", "cp4-epsps"],
    ["CP4 EPSPS", "cp4 epsps"],
    ["EPSPS",     "epsps"],
    ["PAT",       "pat"],
    ["BAR",       "bar"]
];
var geneCount = 0;
for (var g = 0; g < proteinToGene.length; g++) {
    try {
        geneCount += doGrepChange(
            "\\b" + proteinToGene[g][0] + "(?=\\s+gene\\b)",
            proteinToGene[g][1]
        );
    } catch (e) {}
}
report.push("Protein \u2192 gene name corrections: " + geneCount);

// --- 5. Superscript {{N}} and {{letter}} markers → Char_Superscript ---
// Handles: plain numeric markers {{1}}, {{2}} (non-native superscripts)
//          and table footnote letter markers {{a}}, {{b}}, {{c}} etc.
// Does NOT match {{fn:N}} or {{en:N}} — those are native footnote refs, handled manually.
try {
    var supCount = 0;
    // Numeric markers: {{1}}, {{2}}, etc.
    supCount += doGrepChange("\\{\\{(\\d+)\\}\\}", "$1", null, "Char_Superscript");
    // Single-letter markers: {{a}}, {{b}}, etc. (table footnote markers)
    supCount += doGrepChange("\\{\\{([a-zA-Z])\\}\\}", "$1", null, "Char_Superscript");
    report.push("Superscript markers restored: " + supCount);
} catch (e) {
    report.push("Superscript fix skipped — 'Char_Superscript' character style not found.");
}

// --- Clean up find/change state ---
app.findGrepPreferences = NothingEnum.nothing;
app.changeGrepPreferences = NothingEnum.nothing;

// --- Report ---
alert(
    "Cleanup complete!\n\n" +
    report.join("\n") +
    "\n\nReminder: This script should run AFTER TableStyler.jsx and InsertFootnotes.jsx.\n" +
    "If you haven't run those yet, undo (Cmd+Z) and run them first."
);
