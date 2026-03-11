// CleanUp.jsx
// Run AFTER placing the .docx file in InDesign and applying paragraph styles.
//
// The following are already handled by extract_tables.py (before placement):
//   - Bullet character stripping (U+2022)
//   - Tilde operator (U+223C) → standard tilde
//   - "Table N." → "Table N:" in description rows
//   - Superscript/footnote markers wrapped as {{...}}
//
// This script handles what remains:
//   1. Double spaces → single space
//   2. Consecutive paragraph returns → single return
//   3. Multiplication sign normalization (digit x Capital → digit × Capital)
//   4. Superscript {{N}} markers → apply Char_Superscript character style
//      NOTE: Run this AFTER creating native footnotes for first occurrences (Part A).
//            If Part A is not done yet, undo step 4 (Cmd+Z), do Part A, then re-run.

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

// --- 4. Superscript {{N}} markers → Char_Superscript ---
try {
    count = doGrepChange("\\{\\{(\\d+)\\}\\}", "$1", null, "Char_Superscript");
    report.push("Superscript markers restored: " + count);
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
    "\n\nReminder: Step 4 (superscript) should run AFTER creating native footnotes.\n" +
    "If needed, undo (Cmd+Z), create footnotes first, then re-run this script."
);
