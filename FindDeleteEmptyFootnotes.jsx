// FindDeleteEmptyFootnotes.jsx
// Finds all {{fn:N}} markers remaining in the document (i.e. markers that
// InsertFootnotes.jsx could not match to a footnote text entry), then steps
// through each one so you can decide to delete or keep it individually.
//
// Typical workflow:
//   1. Run InsertFootnotes.jsx  — inserts matched footnotes, highlights
//      unmatched markers red with the Char_UnmatchedMarker character style.
//   2. Do layout review.  Unmatched markers are visible as red {{fn:N}} text.
//   3. Run this script whenever you are ready to clean up.
//
// For each marker the script:
//   • Scrolls the document view to the marker so you can see it in context.
//   • Shows a dialog with ~60 characters of surrounding text.
//   • OK = delete the marker.  Cancel = keep it and move to the next.
//
// Each deletion is its own undo step (Cmd+Z undoes one at a time).

(function () {

    if (app.documents.length === 0) {
        alert("No document is open.");
        return;
    }

    var doc = app.activeDocument;

    // --- 1. Find all remaining {{fn:N}} markers ---
    app.findGrepPreferences  = NothingEnum.nothing;
    app.changeGrepPreferences = NothingEnum.nothing;
    app.findGrepPreferences.findWhat = "\\{\\{fn:(\\d+)\\}\\}";
    var found = doc.findGrep();
    app.findGrepPreferences  = NothingEnum.nothing;
    app.changeGrepPreferences = NothingEnum.nothing;

    if (found.length === 0) {
        alert("No unmatched {{fn:N}} markers found.\nThe document is clean.");
        return;
    }

    // --- 2. Ask whether to begin the review ---
    var startMsg =
        found.length + " unmatched marker" + (found.length === 1 ? "" : "s") +
        " found.\n\n" +
        "The script will scroll to each one and ask whether to delete it.\n\n" +
        "OK = begin review\nCancel = exit";
    if (!confirm(startMsg)) return;

    // --- 3. Step through each marker ---
    var deleted = 0;
    var kept    = 0;
    var errors  = [];

    for (var i = 0; i < found.length; i++) {
        var match = found[i];

        // Scroll the view to this marker
        try { doc.select(match); } catch (e) {}

        var markerText = "";
        try { markerText = match.contents; } catch (e) { markerText = "{{fn:?}}"; }

        var ctx = getContext(match);

        var prompt =
            "Marker " + (i + 1) + " of " + found.length + ":  " + markerText +
            "\n\n" + ctx +
            "\n\nOK = Delete this marker\nCancel = Keep and continue";

        if (confirm(prompt)) {
            try {
                match.remove();
                deleted++;
            } catch (e) {
                errors.push(markerText + ": " + e.message);
            }
        } else {
            kept++;
        }
    }

    // --- 4. Report ---
    var msg =
        "Review complete.\n\n" +
        "Deleted: " + deleted + "\n" +
        "Kept:    " + kept;
    if (errors.length) {
        msg += "\n\nErrors (" + errors.length + "):\n  " +
               errors.slice(0, 5).join("\n  ");
    }
    alert(msg);


    // -------------------------------------------------------------------------
    // Helper: return a short context string surrounding the matched text.
    // Format:  …before text [{{fn:N}}] after text…
    // -------------------------------------------------------------------------
    function getContext(match) {
        var CHARS = 60;
        try {
            var story    = match.parentStory;
            var startIdx = match.insertionPoints[0].index;
            var endIdx   = match.insertionPoints[match.insertionPoints.length - 1].index;

            var ctxStart = Math.max(0, startIdx - CHARS);
            var ctxEnd   = Math.min(story.characters.length - 1, endIdx + CHARS - 1);

            var before = "", after = "";

            if (startIdx > ctxStart) {
                try {
                    var bc = story.characters.itemByRange(ctxStart, startIdx - 1).contents;
                    before = (bc instanceof Array) ? bc.join("") : String(bc);
                } catch (e) {}
            }

            if (endIdx <= ctxEnd) {
                try {
                    var ac = story.characters.itemByRange(endIdx, ctxEnd).contents;
                    after = (ac instanceof Array) ? ac.join("") : String(ac);
                } catch (e) {}
            }

            // Collapse whitespace and paragraph returns for readable display
            before = before.replace(/[\r\n]+/g, " ").replace(/\s{2,}/g, " ");
            after  = after.replace(/[\r\n]+/g, " ").replace(/\s{2,}/g, " ");

            return (ctxStart > 0 ? "\u2026" : "") +
                   before +
                   "[" + match.contents + "]" +
                   after +
                   (ctxEnd < story.characters.length - 1 ? "\u2026" : "");
        } catch (e) {
            try { return "[" + match.contents + "]"; } catch (e2) { return ""; }
        }
    }

})();
