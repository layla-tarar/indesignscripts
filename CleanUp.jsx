// CleanUp.jsx
// Run AFTER placing the _text.docx file in InDesign and applying paragraph styles.
//
// The following are already handled by extract_tables.py (before placement):
//   - Bullet character stripping (U+2022)
//   - Tilde operator (U+223C) → standard tilde
//   - "Table N." → "Table N:" in table description rows (_text file)
//   - Superscript/footnote markers wrapped as {{...}}
//   - Bold stripped from table cells
//
// This script handles what remains:
//   1. Double spaces → single space
//   2. Consecutive paragraph returns → single return
//   3. Multiplication sign normalization (digit x Capital → digit × Capital)
//   4. "Table N." → "Table N:" in Table_Caption and Table_Heading styled paragraphs
//      (standalone caption lines in the body text, separate from description rows)
//   5. Superscript {{N}} markers → apply Char_Superscript character style
//      NOTE: Run this AFTER creating native footnotes for {{fn:N}} markers (Part A).
//            {{fn:N}} markers are handled manually via native footnote insertion.
//            This step only superscripts plain {{N}} and {{letter}} markers.
//            If Part A is not done yet, undo step 5 (Cmd+Z), do Part A, then re-run.

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
    "\n\nReminder: Step 5 (superscript) should run AFTER creating native footnotes (Part A).\n" +
    "{{fn:N}} markers are handled manually — do not use this script to process them.\n" +
    "If needed, undo (Cmd+Z), create footnotes first, then re-run this script."
);
