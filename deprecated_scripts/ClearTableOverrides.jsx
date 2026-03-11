// ClearTableOverrides.jsx
// For every table in the active document:
//   1. Clears all table style overrides (equivalent to Alt/Opt-click [Basic Table])
//   2. Clears all cell style overrides (equivalent to Alt/Opt-click [None])
//   3. Clears all paragraph style overrides within cells
//   4. Converts the first row to a header row
//   5. Sets all row heights to "At Least" 3pt so rows expand to fit content
//   6. Applies Table_Span paragraph style to the paragraph containing the table
//
// Run AFTER pasting tables from Word, BEFORE applying table/cell styles.
// After running this script, you still need to manually:
//   - Apply the correct table style (TStyle_Simple, TStyle_Approvals, etc.)
//   - Apply edge header cell styles (CStyle_HeaderLeft, CStyle_HeaderRight)
//   - Apply bottom row cell style (CStyle_BodyBottom, CStyle_BodyApprovalBottom)
//   - Adjust column widths as needed
//
// Run from InDesign: File > Scripts > Scripts Panel, then double-click.

var doc = app.activeDocument;
var tableCount = 0;
var headerCount = 0;
var spanCount = 0;
var stories = doc.stories.everyItem().getElements();

// Check if Table_Span style exists
var hasTableSpan = false;
try {
    var tableSpanStyle = doc.paragraphStyles.itemByName("Table_Span");
    tableSpanStyle.name; // force resolve — throws error if not found
    hasTableSpan = true;
} catch (e) {
    hasTableSpan = false;
}

for (var s = 0; s < stories.length; s++) {
    var tables = stories[s].tables.everyItem().getElements();

    for (var t = 0; t < tables.length; t++) {
        var table = tables[t];
        tableCount++;

        // --- 1. Clear table style overrides ---
        try {
            table.appliedTableStyle = doc.tableStyles.itemByName("[Basic Table]");
            table.clearTableStyleOverrides();
        } catch (e) {}

        // --- 2. Clear cell style overrides on every cell ---
        try {
            var cells = table.cells.everyItem().getElements();
            for (var c = 0; c < cells.length; c++) {
                cells[c].appliedCellStyle = doc.cellStyles.itemByName("[None]");
                cells[c].clearCellStyleOverrides();
            }
        } catch (e) {}

        // --- 3. Clear paragraph style overrides in every cell ---
        try {
            var cells = table.cells.everyItem().getElements();
            for (var c = 0; c < cells.length; c++) {
                var texts = cells[c].texts.everyItem().getElements();
                for (var tx = 0; tx < texts.length; tx++) {
                    texts[tx].appliedParagraphStyle = doc.paragraphStyles.itemByName("[Basic Paragraph]");
                    texts[tx].clearOverrides(OverrideType.ALL);
                }
            }
        } catch (e) {}

        // --- 4. Convert first row to header row ---
        try {
            if (table.rows.length > 0 && table.headerRowCount === 0) {
                table.rows[0].rowType = RowTypes.HEADER_ROW;
                headerCount++;
            }
        } catch (e) {}

        // --- 5. Set all row heights to "At Least" 3pt ---
        try {
            var rows = table.rows.everyItem().getElements();
            for (var r = 0; r < rows.length; r++) {
                rows[r].autoGrow = true;
                rows[r].height = "3pt";
            }
        } catch (e) {}

        // --- 6. Apply Table_Span to the paragraph containing the table ---
        if (hasTableSpan) {
            try {
                var parentPara = table.storyOffset.paragraphs[0];
                parentPara.appliedParagraphStyle = tableSpanStyle;
                spanCount++;
            } catch (e) {}
        }
    }
}

// --- Report ---
if (tableCount === 0) {
    alert("No tables found in the document.");
} else {
    var msg =
        "Done!\n\n" +
        "Tables processed: " + tableCount + "\n" +
        "Header rows created: " + headerCount + "\n" +
        "Table_Span applied: " + spanCount + "\n\n" +
        "All table, cell, and paragraph overrides have been cleared.\n" +
        "All row heights set to 'At Least 3pt' (rows will expand to fit content).\n\n" +
        "Next steps (manual, per table):\n" +
        "1. Apply table style (TStyle_Simple or TStyle_Approvals)\n" +
        "2. Force cell styles if needed (Alt/Opt-click)\n" +
        "3. Apply edge header styles (CStyle_HeaderLeft, CStyle_HeaderRight)\n" +
        "4. Apply bottom row style (CStyle_BodyBottom or CStyle_BodyApprovalBottom)\n" +
        "5. Adjust column widths as needed";

    if (!hasTableSpan) {
        msg += "\n\nNote: Table_Span paragraph style not found — skipped span application.";
    }

    alert(msg);
}
