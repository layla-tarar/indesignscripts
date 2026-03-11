import sys
import os

# Ensure we can import locally installed packages (in .python-packages)
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
LOCAL_PACKAGES = os.path.join(SCRIPT_DIR, ".python-packages")
if os.path.isdir(LOCAL_PACKAGES) and LOCAL_PACKAGES not in sys.path:
    sys.path.insert(0, LOCAL_PACKAGES)

from docx import Document  # type: ignore


def extract_tables(source_path: str, output_path: str) -> None:
    """
    Read a source .docx file and write a new .docx
    containing only the tables (in order), with cell text copied.
    """
    if not os.path.isfile(source_path):
        raise FileNotFoundError(f"Source file not found: {source_path}")

    src_doc = Document(source_path)
    out_doc = Document()

    table_index = 1

    for table in src_doc.tables:
        # Optional: add a simple heading / separator before each table
        out_doc.add_paragraph(f"Table {table_index}")

        # --- Detect and extract an optional description row (often merged across columns) ---
        rows = len(table.rows)
        start_row_index = 0

        if rows > 0:
            first_row = table.rows[0]
            first_row_texts = [ (cell.text or "").strip() for cell in first_row.cells ]
            non_empty = [t for t in first_row_texts if t]

            # Heuristic: if all non-empty cells in the first row have the same text,
            # treat this as a description row rather than a data/header row.
            if non_empty and len(set(non_empty)) == 1:
                description_text = non_empty[0]
                # Write the description as a normal paragraph (avoids truncation in cells)
                out_doc.add_paragraph(description_text)
                start_row_index = 1

        # Now work only with the data/header rows we actually want to copy
        data_rows = list(table.rows)[start_row_index:]
        rows = len(data_rows)

        if rows == 0:
            # No remaining rows; just add a blank line and move on
            out_doc.add_paragraph("")
            table_index += 1
            continue

        # Determine the maximum number of visible cells in any remaining row.
        # This is safer for tables that use merged cells.
        max_cols = 0
        for row in data_rows:
            if len(row.cells) > max_cols:
                max_cols = len(row.cells)

        if max_cols == 0:
            # Empty table; just add a blank line and continue
            out_doc.add_paragraph("")
            table_index += 1
            continue

        out_table = out_doc.add_table(rows=rows, cols=max_cols)

        for r, row in enumerate(data_rows):
            # Use the actual number of cells in this row (handles merged cells)
            cols_in_row = len(row.cells)
            for c in range(min(cols_in_row, max_cols)):
                src_cell = row.cells[c]
                out_cell = out_table.cell(r, c)
                # Copy plain text; formatting is not preserved
                out_cell.text = src_cell.text or ""

        # Add a blank paragraph after each table for spacing
        out_doc.add_paragraph("")
        table_index += 1

    out_doc.save(output_path)


def main(argv: list[str]) -> None:
    if len(argv) < 3:
        print("Usage: python extract_tables.py INPUT_DOCX OUTPUT_DOCX")
        sys.exit(1)

    source_path = argv[1]
    output_path = argv[2]

    try:
        extract_tables(source_path, output_path)
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main(sys.argv)

