// CleanupAfterPlacement.jsx
// Runs all Find/Change operations from Phase 2, Step 6 in sequence.
// Run AFTER placing the .txt file and applying Body_text to everything.
// Run BEFORE creating native footnotes (Part A) — this script handles Part B
// (superscripting remaining {{}} markers) at the end.
//
// Run from InDesign: File > Scripts > Scripts Panel, then double-click.

var doc = app.activeDocument;
var results;
var report = [];

// --- Helper function ---
function doGrepChange(findWhat, changeTo, findStyle, changeCharStyle) {
    app.findGrepPreferences = NothingEnum.nothing;
    app.changeGrepPreferences = NothingEnum.nothing;

    app.findGrepPreferences.findWhat = findWhat;
    app.changeGrepPreferences.changeTo = changeTo;

    if (findStyle) {
        app.findGrepPreferences.appliedParagraphStyle = doc.paragraphStyles.itemByName(findStyle);
    }
    if (changeCharStyle) {
        app.changeGrepPreferences.appliedCharacterStyle = doc.characterStyles.itemByName(changeCharStyle);
    }

    var results = doc.changeGrep();
    return results.length;
}

// --- 1. Double spaces to single space ---
var count = doGrepChange(" {2,}", " ");
report.push("Double spaces fixed: " + count);

// --- 2. Extra paragraph returns to single return ---
count = doGrepChange("\\r{2,}", "\\r");
report.push("Extra returns fixed: " + count);

// --- 3. Strip bullet characters (bullet + space at start of paragraph) ---
count = doGrepChange("^\\x{2022} ", "");
report.push("Bullet characters stripped: " + count);

// --- 4. Tilde operator to standard tilde ---
count = doGrepChange("\\x{223C}", "~");
report.push("Tilde operators fixed: " + count);

// --- 5. Normalize multiplication signs in stacked event names ---
// (?<=\d)\sx\s(?=[A-Z]) → space + multiplication sign + space
count = doGrepChange("(?<=\\d)\\sx\\s(?=[A-Z])", " \\x{00D7} ");
report.push("Multiplication signs normalized: " + count);

// --- 6. Convert "Table N." to "Table N:" in table headings ---
try {
    count = doGrepChange("(Table \\d+)\\.", "$1:", "Table_Heading");
    report.push("Table heading periods to colons: " + count);
} catch (e) {
    report.push("Table heading fix skipped (Table_Heading style not found)");
}

// --- 7. Superscript remaining {{}} markers (Part B) ---
// Only run this AFTER you've created native footnotes for first occurrences (Part A).
// If you haven't done Part A yet, this will superscript ALL markers including first occurrences.
try {
    count = doGrepChange("\\{\\{(\\d+)\\}\\}", "$1", null, "Char_Superscript");
    report.push("Superscript markers restored: " + count);
} catch (e) {
    report.push("Superscript fix skipped (Char_Superscript style not found)");
}

// --- Clean up ---
app.findGrepPreferences = NothingEnum.nothing;
app.changeGrepPreferences = NothingEnum.nothing;

// --- Report ---
alert(
    "Cleanup complete!\n\n" +
    report.join("\n") +
    "\n\nReminder: If you haven't created native footnotes yet (Part A),\n" +
    "undo the superscript step (Cmd+Z), create the footnotes first,\n" +
    "then run this script again."
);
