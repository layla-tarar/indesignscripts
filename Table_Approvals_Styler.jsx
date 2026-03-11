// Table Approvals Styler.jsx
// Applies TStyle_Approvals to approval tables and TStyle_Simple to others.

(function () {
    if (app.documents.length === 0) {
        alert("No document is open.");
        return;
    }

    var doc = app.activeDocument;

    // Get table styles
    var approvalStyle = doc.tableStyles.itemByName("TStyle_Approvals");
    var simpleStyle   = doc.tableStyles.itemByName("TStyle_Simple");
    var noneCharStyle = doc.characterStyles.itemByName("[None]");
    var basicParaStyle = doc.paragraphStyles.itemByName("[Basic Paragraph]");

    if (!approvalStyle.isValid || !simpleStyle.isValid) {
        alert("One or both table styles are missing:\n- TStyle_Approvals\n- TStyle_Simple");
        return;
    }

    app.doScript(
        function () {
            var approvalCount = 0;
            var simpleCount = 0;
            var debugLines = [];
            var tableIndex = 1;

            var stories = doc.stories;
            for (var s = 0; s < stories.length; s++) {
                var story = stories[s];
                if (!story.tables || story.tables.length === 0) {
                    continue;
                }

                var tables = story.tables;
                for (var t = 0; t < tables.length; t++) {
                    var table = tables[t];

                    // Ensure the first visual row becomes a header row without
                    // inserting a new blank row above it.
                    try {
                        if (table.rows.length > 0 && table.rows[0].rowType !== RowTypes.HEADER_ROW) {
                            table.rows[0].rowType = RowTypes.HEADER_ROW;
                        }
                    } catch (eHeader) {}

                    // Normalize character/paragraph styling within this table so placement
                    // character styles (like Char_RefURL) don't interfere.
                    try {
                        normalizeTableTextStyles(table, basicParaStyle, noneCharStyle);
                    } catch (e) {
                        // If anything goes wrong here, continue without failing the script.
                    }

                    if (table.rows.length < 1 || table.columns.length < 2) {
                        // Not enough structure to analyze, treat as simple
                        table.appliedTableStyle = simpleStyle;
                        simpleCount++;
                        continue;
                    }

                    // Heuristic: treat the first row as the header row for all tables.
                    // Then classify tables purely based on how many 'x' markers
                    // appear under the country columns in the body.
                    var headerRowIndex = 0;

                    var result = isApprovalTableByBody(table, headerRowIndex);
                    if (result.isApproval) {
                        table.appliedTableStyle = approvalStyle;
                        approvalCount++;
                    } else {
                        table.appliedTableStyle = simpleStyle;
                        simpleCount++;
                        if (debugLines.length < 12) {
                            debugLines.push(
                                "Table #" + tableIndex +
                                " body test: x=" + result.xCells + " / " + result.totalCells +
                                " (" + Math.round((result.xCells / Math.max(1, result.totalCells)) * 100) + "%)"
                            );
                        }
                    }
                    tableIndex++;
                }
            }

            alert(
                "Table approvals script finished.\n\n" +
                "Approval tables: " + approvalCount + "\n" +
                "Non-approval tables: " + simpleCount +
                (debugLines.length ? ("\n\nDebug (first " + debugLines.length + "):\n- " + debugLines.join("\n- ")) : "")
            );
        },
        ScriptLanguage.JAVASCRIPT,
        undefined,
        UndoModes.ENTIRE_SCRIPT,
        "Apply approval/simple table styles"
    );

    // ---- Helpers ----

    // Legacy header-detection helper (no longer used directly).
    // Kept for reference in case we want to reintroduce stricter
    // rules later.
    function findApprovalHeaderRow(table) {
        // Check a bit deeper in case there are 1–2 description/blank rows above the header
        var maxRowToCheck = Math.min(4, table.rows.length); // check row indices 0..3 if present

        for (var r = 0; r < maxRowToCheck; r++) {
            var row = table.rows[r];
            var hasCrop = false;

            for (var i = 0; i < row.cells.length; i++) {
                // Normalize whitespace so multi-line headers like "Event\nName" still match.
                var txt = getCellText(row.cells[i]).toLowerCase().replace(/\s+/g, " ");
                if (txt.indexOf("crop") !== -1) {
                    hasCrop = true;
                }
            }

            if (hasCrop && table.columns.length >= 3) {
                return { matched: true, headerRowIndex: r };
            }
        }

        return { matched: false, headerRowIndex: -1 };
    }

    // Decide if table is approval-type based on body cells under country columns.
    // headerRowIndex is 0 or 1 (which row we detected as header).
    function isApprovalTableByBody(table, headerRowIndex) {
        var startBodyRow = headerRowIndex + 1;
        if (startBodyRow >= table.rows.length) {
            // No body rows
            return { isApproval: false, totalCells: 0, xCells: 0 };
        }

        var startCountryCol = 2; // 0: Crop, 1: Event Name, 2+ : country columns
        if (startCountryCol >= table.columns.length) {
            // No country columns
            return { isApproval: false, totalCells: 0, xCells: 0 };
        }

        var totalCells = 0;
        var xCells = 0;

        for (var r = startBodyRow; r < table.rows.length; r++) {
            var row = table.rows[r];
            var cells = row.cells;
            for (var i = 0; i < cells.length; i++) {
                var cell = cells[i];
                var colIndex = getCellColumnIndex(cell);
                if (colIndex < startCountryCol) {
                    continue;
                }

                // If a cell spans multiple country columns, count it proportionally.
                var span = getCellColumnSpan(cell);
                if (span < 1) span = 1;

                var txt = getCellText(cell);
                var normalized = txt.toLowerCase().replace(/[^a-z]/g, "");

                if (normalized === "x") {
                    xCells += span;
                }
                totalCells += span;
            }
        }

        if (totalCells === 0) {
            return { isApproval: false, totalCells: 0, xCells: 0 };
        }

        var ratio = xCells / totalCells;
        var isApproval = ratio >= 0.10; // 10% or more

        return { isApproval: isApproval, totalCells: totalCells, xCells: xCells };
    }

    function getCellText(cell) {
        var txt = "";
        try {
            txt = cell.contents;
        } catch (e) {
            txt = "";
        }

        if (txt === undefined || txt === null) {
            txt = "";
        }

        if (txt instanceof Array) {
            txt = txt.join(" ");
        }

        return String(txt).replace(/^\s+|\s+$/g, "");
    }

    function buildNoHeaderDebugLine(table, tableNumber) {
        var parts = [];
        var maxRowsToShow = Math.min(4, table.rows.length);
        for (var r = 0; r < maxRowsToShow; r++) {
            var row = table.rows[r];
            var rowTexts = [];
            for (var c = 0; c < row.cells.length; c++) {
                var txt = getCellText(row.cells[c]).replace(/\s+/g, " ");
                if (txt.length > 40) {
                    txt = txt.substring(0, 37) + "...";
                }
                rowTexts.push("[" + txt + "]");
            }
            parts.push("r" + r + ": " + rowTexts.join(" "));
        }
        return "Table #" + tableNumber + " skipped (no header match). rows=" +
            table.rows.length + ", cols=" + table.columns.length +
            ". head rows: " + parts.join(" | ");
    }

    function getRowCellByColumnIndex(row, desiredColumnIndex) {
        // InDesign rows can have merged cells; row.cells[n] is not guaranteed to be column n.
        // Prefer matching by cell.parentColumn.index when available.
        try {
            var cells = row.cells;
            for (var i = 0; i < cells.length; i++) {
                var cell = cells[i];
                if (getCellColumnIndex(cell) === desiredColumnIndex) {
                    return cell;
                }
            }
            // Fallback: direct index if present
            if (cells.length > desiredColumnIndex) {
                return cells[desiredColumnIndex];
            }
        } catch (e) {}
        return null;
    }

    function getCellColumnIndex(cell) {
        // Works across more InDesign versions than cell.columnIndex
        try {
            if (cell && cell.parentColumn && typeof cell.parentColumn.index === "number") {
                return cell.parentColumn.index;
            }
        } catch (e) {}
        // Fallback: some versions expose columnIndex; try it if present
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

    function normalizeTableTextStyles(table, basicParaStyle, noneCharStyle) {
        var cells = table.cells;
        for (var i = 0; i < cells.length; i++) {
            var cell = cells[i];

            // First normalize paragraph styles (like Opt/Alt+click on [Basic Paragraph])
            if (basicParaStyle && basicParaStyle.isValid) {
                var paras = cell.paragraphs;
                for (var p = 0; p < paras.length; p++) {
                    var para = paras[p];
                    try {
                        para.appliedParagraphStyle = basicParaStyle;
                        para.clearOverrides(OverrideType.ALL, true);
                    } catch (e1) {}
                }
            }

            // Then normalize character styles inside the cell
            if (noneCharStyle && noneCharStyle.isValid) {
                var texts = cell.texts;
                for (var j = 0; j < texts.length; j++) {
                    var txt = texts[j];
                    try {
                        txt.appliedCharacterStyle = noneCharStyle;
                        txt.clearOverrides(OverrideType.ALL, true);
                    } catch (e2) {}
                }
            }
        }
    }
})();

