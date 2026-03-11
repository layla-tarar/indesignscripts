// TableStyler.jsx
// One-shot table formatting script. For every table in the active document:
//
//   Phase 1 — Reset
//     1. Clear table style overrides (→ [Basic Table])
//     2. Clear cell style overrides (→ [None])
//     3. Clear paragraph and character style overrides within cells
//
//   Phase 2 — Structure (ALL tables)
//     4. Convert first row to a header row
//     5. Set all row heights to "At Least" 3pt (rows expand to fit content)
//     6. Apply Table_Span paragraph style to the paragraph containing the table
//     7. Set table width to 504pt, flush left
//
//   Phase 3 — Classification & styling
//     8. Detect approval tables (≥10% of body cells in country columns contain "x")
//     9. Apply TStyle_Approvals or TStyle_Simple accordingly
//
//   Phase 4 — Approval table detail styling (approval tables only)
//    10. Header cell styles:
//          Leftmost  → CStyle_HeaderRotatedLeft (overridden to CStyle_Header_left if "Crop")
//          Rightmost → CStyle_HeaderRotatedRight
//          Contains "Event Name" → CStyle_Header_middle
//    11. Body column styles:
//          Column whose header contains "Crop"       → CStyle_BodyApproval_Crop
//          Column whose header contains "Event Name" → CStyle_BodyApproval_Event
//    12. "x" cell styles (override column style):
//          Body cell = "x", no fill  → CStyle_BodyApproval_x
//          Body cell = "x", has fill → CStyle_BodyApproval_x_Highlight
//    13. Distribute country columns (not Crop / Event Name) evenly across remaining width
//
// Run AFTER placing the _tables.docx and BEFORE manual bottom-row cell-style adjustments.

(function () {
    if (app.documents.length === 0) {
        alert("No document is open.");
        return;
    }

    var doc = app.activeDocument;

    // --- Resolve required table styles ---
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

    // Resolve approval cell styles — warn but do not abort if missing
    var acs = {
        headerRotLeft:  getCellStyleSafe(doc, "CStyle_HeaderRotatedLeft"),
        headerRotRight: getCellStyleSafe(doc, "CStyle_HeaderRotatedRight"),
        headerCrop:     getCellStyleSafe(doc, "CStyle_Header_left"),
        headerEvent:    getCellStyleSafe(doc, "CStyle_Header_middle"),
        bodyCrop:       getCellStyleSafe(doc, "CStyle_BodyApproval_Crop"),
        bodyEvent:      getCellStyleSafe(doc, "CStyle_BodyApproval_Event"),
        bodyX:          getCellStyleSafe(doc, "CStyle_BodyApproval_x"),
        bodyXHighlight: getCellStyleSafe(doc, "CStyle_BodyApproval_x_Highlight")
    };

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

                    setRowHeights(table, 3);

                    if (hasTableSpan) {
                        try {
                            var parentPara = table.storyOffset.paragraphs[0];
                            parentPara.appliedParagraphStyle = tableSpanStyle;
                            counts.spans++;
                        } catch (e) {}
                    }

                    setTableFullWidth(table, 504);

                    // --- Phase 3: Classify and style ---
                    if (table.rows.length < 1 || table.columns.length < 2) {
                        table.appliedTableStyle = simpleStyle;
                        counts.simple++;
                        continue;
                    }

                    var result = isApprovalTable(table);
                    if (result.isApproval) {
                        table.appliedTableStyle = approvalStyle;
                        // --- Phase 4: Approval detail styling ---
                        styleApprovalTable(table, acs);
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
            msg += "\n\nRemaining manual steps (per table):\n" +
                   "1. Apply bottom-row cell style (CStyle_BodyBottom or CStyle_BodyApprovalBottom)\n" +
                   "2. Review column widths";

            alert(msg);
        },
        ScriptLanguage.JAVASCRIPT,
        undefined,
        UndoModes.ENTIRE_SCRIPT,
        "TableStyler"
    );


    // =========================================================================
    // Phase 1 helper: clear all style overrides on a table
    // =========================================================================
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


    // =========================================================================
    // Phase 2 helpers
    // =========================================================================
    function setRowHeights(table, minHeightPt) {
        try {
            var rows = table.rows.everyItem().getElements();
            for (var r = 0; r < rows.length; r++) {
                rows[r].autoGrow = true;
                rows[r].height   = minHeightPt; // numeric points
            }
        } catch (e) {}
    }

    function setTableFullWidth(table, totalWidth) {
        try { table.width = totalWidth; } catch (e) {}
        try { table.horizontalLayoutAlignment = HorizontalAlignment.LEFT_ALIGN; } catch (e) {}
    }


    // =========================================================================
    // Phase 3 helper: classify table as approval or simple
    // Heuristic: ≥10% of body cells in country columns (col 2+) contain "x"
    // =========================================================================
    function isApprovalTable(table) {
        var startBodyRow    = 1; // first row is always treated as header
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


    // =========================================================================
    // Phase 4: approval table — fix header row height after text rotation
    // TStyle_Approvals applies text rotation to header cells; this runs after
    // the style is applied so InDesign uses the rotated text dimensions when
    // computing the minimum row height.
    // =========================================================================
    function styleApprovalTable(table, styles) {
        if (table.rows.length === 0) return;
        try {
            var hRow = table.rows[0];
            hRow.autoGrow = true;
            hRow.height   = 3; // numeric points (At Least 3pt) — row expands to fit rotated text
        } catch (e) {}
    }


    // -------------------------------------------------------------------------
    // Distribute all non-Crop, non-EventName columns to equal width,
    // filling the space left after the reserved columns.
    // -------------------------------------------------------------------------
    function distributeCountryColumns(table, totalWidth, cropColIdx, eventColIdx) {
        try {
            var cols    = table.columns.everyItem().getElements();
            var numCols = cols.length;

            // Measure reserved column widths
            var reservedWidth = 0;
            var reservedCount = 0;
            for (var c = 0; c < numCols; c++) {
                if (c === cropColIdx || c === eventColIdx) {
                    try { reservedWidth += cols[c].width; } catch (e) {}
                    reservedCount++;
                }
            }

            var countryCount = numCols - reservedCount;
            if (countryCount <= 0) return;

            var countryWidth = (totalWidth - reservedWidth) / countryCount;
            if (countryWidth < 1) return; // safety: don't collapse columns

            for (var c = 0; c < numCols; c++) {
                if (c !== cropColIdx && c !== eventColIdx) {
                    try { cols[c].width = countryWidth; } catch (e) {}
                }
            }
        } catch (e) {}
    }


    // =========================================================================
    // Low-level helpers
    // =========================================================================
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

    function cellHasFill(cell) {
        try {
            var fc = cell.fillColor;
            return fc && fc.isValid && fc.name !== "[None]" && fc.name !== "None";
        } catch (e) {}
        return false;
    }

    function getCellStyleSafe(doc, name) {
        try {
            var s = doc.cellStyles.itemByName(name);
            if (s && s.isValid) return s;
        } catch (e) {}
        return null;
    }

    function setCellStyle(cell, style) {
        if (!style) return;
        try { cell.appliedCellStyle = style; } catch (e) {}
    }

})();
