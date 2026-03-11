// TableStyler.jsx
// One-shot table formatting script. For every table in the active document:
//
//   Phase 1 — Reset
//     1. Clear table style overrides (→ [Basic Table])
//     2. Clear cell style overrides (→ [None])
//     3. Clear paragraph and character style overrides within cells
//
//   Phase 2 — Structure
//     4. Convert first row to a header row
//     5. Set all row heights to "At Least" 3pt (rows expand to fit content)
//     6. Apply Table_Span paragraph style to the paragraph containing the table
//
//   Phase 3 — Classification & styling
//     7. Detect approval tables (≥10% of body cells in country columns contain "x")
//     8. Apply TStyle_Approvals or TStyle_Simple accordingly
//
// Run AFTER placing the _tables.docx and BEFORE manual column/cell-style adjustments.
// Run from InDesign: File > Scripts > Scripts Panel, then double-click.

(function () {
    if (app.documents.length === 0) {
        alert("No document is open.");
        return;
    }

    var doc = app.activeDocument;

    // --- Resolve styles (fail fast if essential ones are missing) ---
    var approvalStyle = doc.tableStyles.itemByName("TStyle_Approvals");
    var simpleStyle   = doc.tableStyles.itemByName("TStyle_Simple");
    if (!approvalStyle.isValid || !simpleStyle.isValid) {
        alert("Missing required table styles:\n- TStyle_Approvals\n- TStyle_Simple\n\nScript aborted.");
        return;
    }

    var basicTableStyle = doc.tableStyles.itemByName("[Basic Table]");
    var noneCellStyle   = doc.cellStyles.itemByName("[None]");
    var basicParaStyle  = doc.paragraphStyles.itemByName("[Basic Paragraph]");
    var noneCharStyle   = doc.characterStyles.itemByName("[None]");

    var hasTableSpan = false;
    var tableSpanStyle;
    try {
        tableSpanStyle = doc.paragraphStyles.itemByName("Table_Span");
        tableSpanStyle.name; // force resolve — throws if not found
        hasTableSpan = true;
    } catch (e) {}

    // --- Run everything as one undoable action ---
    app.doScript(
        function () {
            var counts = {
                tables:    0,
                headers:   0,
                spans:     0,
                approvals: 0,
                simple:    0
            };
            var debugLines = [];

            var stories = doc.stories.everyItem().getElements();
            for (var s = 0; s < stories.length; s++) {
                var tables = stories[s].tables.everyItem().getElements();
                for (var t = 0; t < tables.length; t++) {
                    var table = tables[t];
                    counts.tables++;

                    // --- Phase 1: Reset ---
                    clearTableOverrides(table, doc, basicTableStyle, noneCellStyle,
                                        basicParaStyle, noneCharStyle);

                    // --- Phase 2: Structure ---
                    if (table.rows.length > 0 && table.headerRowCount === 0) {
                        try {
                            table.rows[0].rowType = RowTypes.HEADER_ROW;
                            counts.headers++;
                        } catch (e) {}
                    }

                    setRowHeights(table, "3pt");

                    if (hasTableSpan) {
                        try {
                            var parentPara = table.storyOffset.paragraphs[0];
                            parentPara.appliedParagraphStyle = tableSpanStyle;
                            counts.spans++;
                        } catch (e) {}
                    }

                    // --- Phase 3: Classify and style ---
                    if (table.rows.length < 1 || table.columns.length < 2) {
                        table.appliedTableStyle = simpleStyle;
                        counts.simple++;
                        continue;
                    }

                    var result = isApprovalTable(table);
                    if (result.isApproval) {
                        table.appliedTableStyle = approvalStyle;
                        counts.approvals++;
                    } else {
                        table.appliedTableStyle = simpleStyle;
                        counts.simple++;
                        if (debugLines.length < 12) {
                            debugLines.push(
                                "Table #" + counts.tables +
                                " → simple (x=" + result.xCells + "/" + result.totalCells +
                                ", " + Math.round(result.xCells / Math.max(1, result.totalCells) * 100) + "%)"
                            );
                        }
                    }
                }
            }

            // --- Report ---
            if (counts.tables === 0) {
                alert("No tables found in the document.");
                return;
            }

            var msg =
                "TableStyler complete!\n\n" +
                "Tables processed:      " + counts.tables    + "\n" +
                "Header rows set:       " + counts.headers   + "\n" +
                "Table_Span applied:    " + counts.spans     + "\n" +
                "Approval tables:       " + counts.approvals + "\n" +
                "Simple tables:         " + counts.simple;

            if (!hasTableSpan) {
                msg += "\n\nNote: Table_Span style not found — skipped.";
            }
            if (debugLines.length) {
                msg += "\n\nNon-approval tables (first " + debugLines.length + "):\n- " +
                       debugLines.join("\n- ");
            }
            msg += "\n\nNext steps (manual, per table):\n" +
                   "1. Apply edge header cell styles (CStyle_HeaderLeft, CStyle_HeaderRight)\n" +
                   "2. Apply bottom row cell style (CStyle_BodyBottom or CStyle_BodyApprovalBottom)\n" +
                   "3. Adjust column widths as needed";

            alert(msg);
        },
        ScriptLanguage.JAVASCRIPT,
        undefined,
        UndoModes.ENTIRE_SCRIPT,
        "TableStyler"
    );


    // -------------------------------------------------------------------------
    // Phase 1 helper: clear all style overrides on a table
    // -------------------------------------------------------------------------
    function clearTableOverrides(table, doc, basicTableStyle, noneCellStyle,
                                  basicParaStyle, noneCharStyle) {
        // 1a. Table style overrides
        try {
            if (basicTableStyle && basicTableStyle.isValid) {
                table.appliedTableStyle = basicTableStyle;
            }
            table.clearTableStyleOverrides();
        } catch (e) {}

        // 1b. Cell style overrides — preserve fill colors before clearing
        try {
            var cells = table.cells.everyItem().getElements();

            // Save fill colors first (clearCellStyleOverrides resets fill to [None])
            var cellFills = [];
            for (var c = 0; c < cells.length; c++) {
                var fill = null;
                try {
                    var fc = cells[c].fillColor;
                    if (fc && fc.isValid && fc.name !== "[None]" && fc.name !== "None") {
                        fill = { color: fc, tint: cells[c].fillTint };
                    }
                } catch (ef) {}
                cellFills.push(fill);
            }

            for (var c = 0; c < cells.length; c++) {
                if (noneCellStyle && noneCellStyle.isValid) {
                    cells[c].appliedCellStyle = noneCellStyle;
                }
                cells[c].clearCellStyleOverrides();
            }

            // Restore fill colors
            for (var c = 0; c < cells.length; c++) {
                if (cellFills[c]) {
                    try {
                        cells[c].fillColor = cellFills[c].color;
                        cells[c].fillTint  = cellFills[c].tint;
                    } catch (er) {}
                }
            }
        } catch (e) {}

        // 1c. Paragraph and character style overrides within cells
        try {
            var cells = table.cells.everyItem().getElements();
            for (var c = 0; c < cells.length; c++) {
                var cell = cells[c];

                if (basicParaStyle && basicParaStyle.isValid) {
                    var paras = cell.paragraphs;
                    for (var p = 0; p < paras.length; p++) {
                        try {
                            paras[p].appliedParagraphStyle = basicParaStyle;
                            paras[p].clearOverrides(OverrideType.ALL, true);
                        } catch (e1) {}
                    }
                }

                if (noneCharStyle && noneCharStyle.isValid) {
                    var texts = cell.texts;
                    for (var j = 0; j < texts.length; j++) {
                        try {
                            texts[j].appliedCharacterStyle = noneCharStyle;
                            texts[j].clearOverrides(OverrideType.ALL, true);
                        } catch (e2) {}
                    }
                }
            }
        } catch (e) {}
    }


    // -------------------------------------------------------------------------
    // Phase 2 helper: set all row heights to "At Least N"
    // -------------------------------------------------------------------------
    function setRowHeights(table, minHeight) {
        try {
            var rows = table.rows.everyItem().getElements();
            for (var r = 0; r < rows.length; r++) {
                rows[r].autoGrow = true;
                rows[r].height = minHeight;
            }
        } catch (e) {}
    }


    // -------------------------------------------------------------------------
    // Phase 3 helper: classify table as approval or simple
    // Heuristic: ≥10% of body cells in country columns (col 2+) contain "x"
    // -------------------------------------------------------------------------
    function isApprovalTable(table) {
        var headerRowIndex  = 0; // first row is always treated as header
        var startBodyRow    = headerRowIndex + 1;
        var startCountryCol = 2; // col 0: Crop, col 1: Event Name, col 2+: countries

        if (startBodyRow >= table.rows.length || startCountryCol >= table.columns.length) {
            return { isApproval: false, totalCells: 0, xCells: 0 };
        }

        var totalCells = 0;
        var xCells     = 0;

        for (var r = startBodyRow; r < table.rows.length; r++) {
            var cells = table.rows[r].cells;
            for (var i = 0; i < cells.length; i++) {
                var cell = cells[i];
                if (getCellColumnIndex(cell) < startCountryCol) continue;

                var span = getCellColumnSpan(cell);
                var txt  = getCellText(cell).toLowerCase().replace(/[^a-z]/g, "");
                if (txt === "x") xCells += span;
                totalCells += span;
            }
        }

        if (totalCells === 0) return { isApproval: false, totalCells: 0, xCells: 0 };

        return {
            isApproval: (xCells / totalCells) >= 0.10,
            totalCells: totalCells,
            xCells:     xCells
        };
    }


    // -------------------------------------------------------------------------
    // Low-level cell helpers
    // -------------------------------------------------------------------------
    function getCellText(cell) {
        var txt = "";
        try { txt = cell.contents; } catch (e) {}
        if (txt === undefined || txt === null) txt = "";
        if (txt instanceof Array) txt = txt.join(" ");
        return String(txt).replace(/^\s+|\s+$/g, "");
    }

    function getCellColumnIndex(cell) {
        try {
            if (cell && cell.parentColumn && typeof cell.parentColumn.index === "number") {
                return cell.parentColumn.index;
            }
        } catch (e) {}
        try {
            var n = Number(cell.columnIndex);
            if (!isNaN(n)) return n;
        } catch (e2) {}
        return -1;
    }

    function getCellColumnSpan(cell) {
        try {
            var n = Number(cell.columnSpan);
            if (!isNaN(n) && n >= 1) return n;
        } catch (e) {}
        return 1;
    }

})();
