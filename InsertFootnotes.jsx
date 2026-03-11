// InsertFootnotes.jsx
// Replaces every {{fn:N}} marker in the active document with a native InDesign footnote.
// Footnote text is read from the companion _footnotes.txt produced by extract_tables.py.
//
// _footnotes.txt format (UTF-8, tab-separated):
//   fn:1<TAB>First footnote text here.
//   fn:2<TAB>Second footnote. Multi-paragraph text uses \r as paragraph separator.
//
// Workflow:
//   1. python extract_tables.py your_file.docx     → produces _footnotes.txt
//   2. Place _text.docx in InDesign
//   3. Run CleanUp.jsx (all steps — step 5 will skip {{fn:N}}, only hits {{N}}/{{letter}})
//   4. Run InsertFootnotes.jsx — select _footnotes.txt when prompted
//
// Technical notes:
//   - Markers are processed in REVERSE document order so text positions stay valid.
//   - Each footnote is inserted at the marker's start position; the marker is then deleted.
//   - \r in footnote text produces a paragraph break within the footnote body.

(function () {

    if (app.documents.length === 0) {
        alert("No document is open.");
        return;
    }

    var doc = app.activeDocument;

    // --- 1. Select the _footnotes.txt file ---
    var f = File.openDialog(
        "Select the _footnotes.txt file",
        "Text files:*.txt,All files:*.*"
    );
    if (!f) return; // user cancelled

    // --- 2. Read and parse ---
    f.encoding = "UTF-8";
    if (!f.open("r")) {
        alert("Could not open:\n" + f.fsName);
        return;
    }

    var footnotes = {}; // { "fn:1": "text...", ... }
    while (!f.eof) {
        var line = f.readln();
        var tabIdx = line.indexOf("\t");
        if (tabIdx < 1) continue;
        var key  = line.substring(0, tabIdx);  // "fn:1"
        var text = line.substring(tabIdx + 1); // "Footnote text..."
        footnotes[key] = text;
    }
    f.close();

    var fnCount = 0;
    for (var k in footnotes) { fnCount++; }
    if (fnCount === 0) {
        alert("No entries found in:\n" + f.fsName +
              "\n\nExpected tab-separated lines: fn:1<TAB>text");
        return;
    }

    // --- 3. Find all {{fn:N}} markers ---
    app.findGrepPreferences = NothingEnum.nothing;
    app.changeGrepPreferences = NothingEnum.nothing;
    app.findGrepPreferences.findWhat = "\\{\\{fn:(\\d+)\\}\\}";
    var found = doc.findGrep();
    app.findGrepPreferences = NothingEnum.nothing;
    app.changeGrepPreferences = NothingEnum.nothing;

    if (found.length === 0) {
        alert("No {{fn:N}} markers found in the document.\n\n" +
              "Ensure _text.docx has been placed and CleanUp.jsx has run.");
        return;
    }

    // --- 4. Insert footnotes in reverse order ---
    // Reverse order preserves the validity of earlier text positions after each edit.
    var inserted = 0;
    var missing  = [];
    var errors   = [];

    app.doScript(
        function () {
            for (var i = found.length - 1; i >= 0; i--) {
                var match      = found[i];
                var markerText = match.contents; // e.g. "{{fn:1}}"

                // Extract key: "{{fn:1}}" → "fn:1"
                var key = markerText.replace(/^\{\{/, "").replace(/\}\}$/, "");

                var fnText = footnotes[key];
                if (fnText === undefined) {
                    missing.push(markerText);
                    continue;
                }

                try {
                    // Insert footnote at the character position just before the marker.
                    // footnotes.add() inserts a footnote reference char at that point
                    // and returns the Footnote object.
                    var startIP  = match.insertionPoints[0];
                    var footnote = startIP.footnotes.add();

                    // Populate the footnote body. The newly created footnote is empty;
                    // inserting at insertionPoints[-1] appends the text.
                    // \r in fnText produces paragraph breaks inside the footnote.
                    footnote.insertionPoints[-1].contents = fnText;

                    // Delete the original marker. After the footnotes.add() call, the
                    // match object has shifted one character to the right (the inserted
                    // footnote ref char), but it still refers to the same "{{fn:1}}" text.
                    match.remove();

                    inserted++;
                } catch (e) {
                    errors.push(markerText + ": " + e.message);
                }
            }
        },
        ScriptLanguage.JAVASCRIPT,
        undefined,
        UndoModes.ENTIRE_SCRIPT,
        "InsertFootnotes"
    );

    // --- 5. Report ---
    var msg =
        "InsertFootnotes complete!\n\n" +
        "Entries in txt file: " + fnCount         + "\n" +
        "Markers found:       " + found.length    + "\n" +
        "Footnotes inserted:  " + inserted;

    if (missing.length) {
        msg += "\n\nMarkers with no matching txt entry (" + missing.length + "):\n  " +
               missing.slice(0, 10).join("\n  ");
        if (missing.length > 10) {
            msg += "\n  … and " + (missing.length - 10) + " more";
        }
    }

    if (errors.length) {
        msg += "\n\nErrors (" + errors.length + "):\n  " +
               errors.slice(0, 5).join("\n  ");
    }

    alert(msg);

})();
