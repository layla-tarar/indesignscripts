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
// Run AFTER placing the _clean.docx and BEFORE CleanUp.jsx.
// Phase 1 clears all character overrides in table cells — CleanUp.jsx (step 5) must
// run afterward to restore Char_Superscript on {{a}}/{{b}}/{{N}} markers in cells.

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
                        // Capture highlighted "x" cells BEFORE applying table style —
                        // TStyle_Approvals can clear fill overrides, so detection must
                        // happen here while original fills are still on the cells.
                        var xHighlightCells = captureXHighlightCells(table);
                        table.appliedTableStyle = approvalStyle;
                        // --- Phase 4: Approval detail styling ---
                        styleApprovalTable(table, acs, xHighlightCells);
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


    // Scan all body "x" cells for fill colors BEFORE the table style is applied.
    // Returns an array of cell references for cells containing only "x" with a fill.
    // Must be called before table.appliedTableStyle = approvalStyle.
    function captureXHighlightCells(table) {
        var result = [];
        try {
            var cols = table.columns.everyItem().getElements();
            for (var c = 1; c < cols.length; c++) { // skip col 0 (Crop)
                var cells = cols[c].cells.everyItem().getElements();
                for (var i = 1; i < cells.length; i++) { // skip header cell (index 0)
                    try {
                        var cell = cells[i];
                        var txt  = getCellText(cell).toLowerCase().replace(/[^a-z]/g, "");
                        if (txt === "x" && cellHasFill(cell)) result.push(cell);
                    } catch(e) {}
                }
            }
        } catch(e) {}
        return result;
    }


    // =========================================================================
    // Phase 4: approval table detail styling
    //
    // Order of operations (important — later passes override earlier ones):
    //   1. Header row height   — must be after TStyle_Approvals applies rotation
    //   2. Re-assert 504pt width — applying the style can reset table width
    //   3. Header cell styles  — positional (first/last) then content (Event Name)
    //   4. Column body styles  — Event Name column → CStyle_BodyApproval_Event
    //   5. Distribute country columns evenly across remaining width
    //   6. "x" cell overrides  — individual cells, must be AFTER column styles
    // =========================================================================
    function styleApprovalTable(table, styles, xHighlightCells) {
        if (table.rows.length === 0) return;

        var hRow    = table.rows[0];
        var numCols = hRow.cells.length;

        // ---- ALL CELL STYLING FIRST ----
        // Structural changes (width, column distribution) come after, because InDesign
        // can silently reset cell style overrides when table/column geometry is modified.

        // --- A: Header cell styles + locate Event Name column ---
        // Normalize internal whitespace so "Event\rName" matches as "event name".
        var eventColIdx = -1;

        for (var c = 0; c < numCols; c++) {
            try {
                var hCell = hRow.cells[c];
                var raw   = getCellText(hCell);
                var txt   = raw.toLowerCase().replace(/\s+/g, " ");

                if (c === 0)            setCellStyle(hCell, styles.headerCrop);     // CStyle_Header_left
                if (c === numCols - 1)  setCellStyle(hCell, styles.headerRotRight); // CStyle_HeaderRotatedRight
                if (txt.indexOf("event name") !== -1) {
                    setCellStyle(hCell, styles.headerEvent); // CStyle_Header_middle
                    eventColIdx = c;
                }
            } catch(e) {}
        }

        // --- B: Event Name body column style (first pass) ---
        // Use column.cells instead of row.cells[idx] to avoid merged-cell index shifting:
        // when Crop cells span multiple rows, row.cells is shorter and indices shift,
        // so row.cells[eventColIdx] lands on the wrong cell for rows inside the merge.
        if (eventColIdx !== -1) {
            try {
                var eventColCells = table.columns[eventColIdx].cells.everyItem().getElements();
                for (var ci = 1; ci < eventColCells.length; ci++) { // ci=0 is the header cell
                    try { setCellStyle(eventColCells[ci], styles.bodyEvent); } catch(e) {}
                }
            } catch(e) {}
        }

        // --- C: "x" cell overrides (second pass — wins over column styles) ---
        // Apply bodyX to all "x" cells first, then override with bodyXHighlight for
        // pre-captured highlighted cells (fills were captured before table style was
        // applied, since TStyle_Approvals can clear fill overrides).
        for (var ci = 0; ci < numCols; ci++) {
            if (ci === 0 || ci === eventColIdx) continue; // skip Crop and Event Name
            try {
                var colCells = table.columns[ci].cells.everyItem().getElements();
                for (var ri = 1; ri < colCells.length; ri++) { // ri=0 is the header cell
                    try {
                        var cell    = colCells[ri];
                        var cellTxt = getCellText(cell).toLowerCase().replace(/[^a-z]/g, "");
                        if (cellTxt === "x") setCellStyle(cell, styles.bodyX);
                    } catch(e) {}
                }
            } catch(e) {}
        }
        // Apply highlight style to cells that had fills before the table style was set.
        // Clear cell-level overrides after applying so the style's fill wins over any
        // lingering local fill override that survived table style application.
        for (var hi = 0; hi < xHighlightCells.length; hi++) {
            try {
                var hCell = xHighlightCells[hi];
                setCellStyle(hCell, styles.bodyXHighlight);
                // clearOverrides() is unreliable on cell objects; instead, read the fill
                // directly from the cell style and stamp it onto the cell so the imported
                // Word fill color (local override) is replaced with the correct value.
                try {
                    hCell.fillColor = styles.bodyXHighlight.fillColor;
                    hCell.fillTint  = styles.bodyXHighlight.fillTint;
                } catch(e) {}
            } catch(e) {}
        }

        // ---- ALL STRUCTURAL CHANGES AFTER CELL STYLING ----

        // --- D: Header row height ---
        // Crop (col 0) and Event Name are horizontal — exclude from rotated height estimate.
        try {
            var neededPt  = calcRotatedTextHeight(hRow, [0, eventColIdx]);
            hRow.autoGrow = false;
            hRow.height   = neededPt;
        } catch (e) {}

        // --- E: Re-assert 504pt table width ---
        setTableFullWidth(table, 504);

        // --- F: Size Crop and Event Name columns to fit their horizontal text ---
        try { table.columns[0].width = calcColumnWidth(table, 0, 130); } catch(e) {}
        if (eventColIdx !== -1) {
            try { table.columns[eventColIdx].width = calcColumnWidth(table, eventColIdx); } catch(e) {}
        }

        // --- G: Distribute country columns across the remaining width ---
        distributeCountryColumns(table, 504, 0, eventColIdx);
    }

    // Estimate the row height needed for 90°-rotated text.
    // Row height must equal the text's rendered WIDTH (characters become vertical).
    // skipCols: array of column indices to exclude (non-rotated cells like Crop, Event Name).
    // Formula: characters × avg-char-width (≈ 0.50 × ptSize) + padding
    function calcRotatedTextHeight(row, skipCols) {
        var MIN_HEIGHT = 30; // never shorter than 30pt regardless of content
        var PADDING    = 8;  // top + bottom breathing room in points
        var maxPt      = MIN_HEIGHT;
        skipCols = skipCols || [];

        for (var c = 0; c < row.cells.length; c++) {
            // Skip columns whose text is horizontal (not rotated)
            var skip = false;
            for (var i = 0; i < skipCols.length; i++) {
                if (skipCols[i] === c) { skip = true; break; }
            }
            if (skip) continue;

            try {
                var paras = row.cells[c].paragraphs;
                for (var p = 0; p < paras.length; p++) {
                    var txt = "";
                    try { txt = paras[p].contents; } catch(e) {}
                    if (!txt || txt === "\r") continue;

                    var ptSize = 7.5; // Table_Header style: Source Sans Pro Bold 7.5pt
                    try {
                        var s = paras[p].pointSize;
                        if (typeof s === "number" && s > 0) ptSize = s;
                    } catch(e) {}

                    // avg Latin char width ≈ 0.58 × point size (tuned for Source Sans Pro Bold;
                    // sized to fit longer country names like "New Zealand", "European Union")
                    var est = Math.ceil(txt.length * ptSize * 0.58) + PADDING;
                    if (est > maxPt) maxPt = est;
                }
            } catch(e) {}
        }
        return maxPt;
    }


    // Estimate the minimum width needed for a column with horizontal (unrotated) text.
    // Checks every cell in the column (header + body) and returns the widest estimate.
    // maxWidth: optional upper bound in points (prevents runaway widths from long body cells).
    function calcColumnWidth(table, colIdx, maxWidth) {
        var MIN_WIDTH = 20;
        var PADDING   = 12; // left + right cell padding
        var maxPt     = MIN_WIDTH;

        for (var r = 0; r < table.rows.length; r++) {
            try {
                var row = table.rows[r];
                if (colIdx >= row.cells.length) continue;
                var paras = row.cells[colIdx].paragraphs;
                for (var p = 0; p < paras.length; p++) {
                    var txt = "";
                    try { txt = paras[p].contents; } catch(e) {}
                    if (!txt || txt === "\r") continue;

                    var ptSize = 7.5;
                    try {
                        var s = paras[p].pointSize;
                        if (typeof s === "number" && s > 0) ptSize = s;
                    } catch(e) {}

                    var est = Math.ceil(txt.length * ptSize * 0.55) + PADDING;
                    if (est > maxPt) maxPt = est;
                }
            } catch(e) {}
        }
        return (maxWidth && maxPt > maxWidth) ? maxWidth : maxPt;
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
        txt = String(txt).replace(/^\s+|\s+$/g, "");

        // cell.contents is unreliable for header-row cells in InDesign —
        // fall back to reading paragraph text directly.
        if (txt === "") {
            try {
                var paras = cell.paragraphs;
                var parts = [];
                for (var p = 0; p < paras.length; p++) {
                    var pt = "";
                    try { pt = String(paras[p].contents || ""); } catch(e) {}
                    pt = pt.replace(/\r/g, "").replace(/^\s+|\s+$/g, "");
                    if (pt) parts.push(pt);
                }
                txt = parts.join(" ");
            } catch(e) {}
        }
        return txt;
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
