from __future__ import annotations

import sys
import os
import copy
import re
import tempfile

# Ensure we can import locally installed packages (in .python-packages)
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
LOCAL_PACKAGES = os.path.join(SCRIPT_DIR, ".python-packages")
if os.path.isdir(LOCAL_PACKAGES) and LOCAL_PACKAGES not in sys.path:
    sys.path.insert(0, LOCAL_PACKAGES)

from docx import Document  # type: ignore
from docx.oxml import OxmlElement  # type: ignore
from docx.oxml.ns import qn  # type: ignore
from lxml import etree  # type: ignore


def _make_paragraph_element(text: str, style: str | None = None):
    """Create a bare <w:p> XML element containing the given plain text.
    If style is given, adds <w:pPr><w:pStyle w:val="style"/> so InDesign
    can map the paragraph to the matching paragraph style by name.
    """
    p = OxmlElement("w:p")
    if style:
        pPr = OxmlElement("w:pPr")
        pStyle = OxmlElement("w:pStyle")
        pStyle.set(qn("w:val"), style)
        pPr.append(pStyle)
        p.append(pPr)
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = text
    r.append(t)
    p.append(r)
    return p


def _get_description_row(table) -> str | None:
    """Return description text if first row is a merged title row, else None."""
    if not table.rows:
        return None
    first_row_texts = [(cell.text or "").strip() for cell in table.rows[0].cells]
    non_empty = [t for t in first_row_texts if t]
    if non_empty and len(set(non_empty)) == 1:
        return non_empty[0]
    return None


def _extract_description_rows(doc: Document) -> None:
    """
    For each table whose first row is a merged title/description row (all
    non-empty cells contain the same text), insert that text as a plain
    paragraph immediately before the table, then remove the row from the table.
    """
    tbl_tag = qn("w:tbl")
    tr_tag  = qn("w:tr")
    body    = doc.element.body

    tbl_elements  = [el for el in body if el.tag == tbl_tag]
    table_objects = doc.tables  # python-docx objects, same order as XML elements

    for tbl_el, table in zip(tbl_elements, table_objects):
        description = _get_description_row(table)
        if description is None:
            continue
        # Insert description paragraph immediately before the table,
        # tagged with Table_Header so InDesign maps it to that paragraph style.
        tbl_el.addprevious(_make_paragraph_element(description, style="Table_Header"))
        # Remove the first <w:tr> from the table
        first_tr = tbl_el.find(tr_tag)
        if first_tr is not None:
            tbl_el.remove(first_tr)


def _mark_superscripts(doc: Document) -> None:
    """
    Walk every run in the document (body paragraphs and all table cells).
    - Superscript runs: wrap text as {{text}} and remove the superscript property.
    - Footnote references: replace the <w:footnoteReference> element with {{fn:N}}.
    - Endnote references: replace the <w:endnoteReference> element with {{en:N}}.
    """
    fn_tag = qn("w:footnoteReference")
    en_tag = qn("w:endnoteReference")
    va_tag = qn("w:vertAlign")
    rpr_tag = qn("w:rPr")

    def process_run(run):
        r_el = run._r

        # --- Footnote reference ---
        fn_ref = r_el.find(fn_tag)
        if fn_ref is not None:
            fn_id = fn_ref.get(qn("w:id"), "?")
            r_el.remove(fn_ref)
            rpr = r_el.find(rpr_tag)
            if rpr is not None:
                va = rpr.find(va_tag)
                if va is not None:
                    rpr.remove(va)
            t = OxmlElement("w:t")
            t.text = f"{{{{fn:{fn_id}}}}}"
            r_el.append(t)
            return

        # --- Endnote reference ---
        en_ref = r_el.find(en_tag)
        if en_ref is not None:
            en_id = en_ref.get(qn("w:id"), "?")
            r_el.remove(en_ref)
            rpr = r_el.find(rpr_tag)
            if rpr is not None:
                va = rpr.find(va_tag)
                if va is not None:
                    rpr.remove(va)
            t = OxmlElement("w:t")
            t.text = f"{{{{en:{en_id}}}}}"
            r_el.append(t)
            return

        # --- Regular superscript run ---
        if run.font.superscript:
            run.text = f"{{{{{run.text}}}}}"
            run.font.superscript = False

    def process_paragraphs(paragraphs):
        for para in paragraphs:
            for run in para.runs:
                process_run(run)

    process_paragraphs(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                process_paragraphs(cell.paragraphs)


def _strip_field_codes(doc: Document) -> None:
    """
    Remove Word field code machinery from all paragraphs, keeping only the
    display text (the portion between 'separate' and 'end' markers).
    """
    fldchar_tag      = qn("w:fldChar")
    r_tag            = qn("w:r")
    fldchartype_attr = qn("w:fldCharType")

    def strip_fields(p_el):
        in_instruction = False
        for child in list(p_el):
            if child.tag != r_tag:
                continue
            fldchar = child.find(fldchar_tag)
            if fldchar is not None:
                ftype = fldchar.get(fldchartype_attr, "")
                if ftype == "begin":
                    in_instruction = True
                    p_el.remove(child)
                elif ftype == "separate":
                    in_instruction = False
                    p_el.remove(child)
                elif ftype == "end":
                    p_el.remove(child)
            elif in_instruction:
                p_el.remove(child)

    for para in doc.paragraphs:
        strip_fields(para._p)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    strip_fields(para._p)


def _strip_character_styles(doc: Document) -> None:
    """
    Strip run-level formatting that would override InDesign paragraph styles on placement:
    - Unwrap <w:hyperlink> elements (keeping link text as plain runs)
    - Remove from <w:rPr>: rStyle, rFonts, sz/szCs, color, b/bCs, u, highlight, kern,
      spacing, lang
    Preserves: i/iCs (italics) and vertAlign (superscript — handled separately).
    """
    rpr_tag       = qn("w:rPr")
    hyperlink_tag = qn("w:hyperlink")

    _STRIP_RPR = {
        qn("w:rStyle"),
        qn("w:rFonts"),
        qn("w:sz"),    qn("w:szCs"),
        qn("w:color"),
        qn("w:b"),     qn("w:bCs"),
        qn("w:u"),
        qn("w:highlight"),
        qn("w:kern"),
        qn("w:spacing"),
        qn("w:lang"),
    }

    ppr_tag = qn("w:pPr")

    def strip_run_overrides(paragraphs):
        for para in paragraphs:
            p_el = para._p
            for hl in p_el.findall(hyperlink_tag):
                parent = hl.getparent()
                idx = list(parent).index(hl)
                for child in list(hl):
                    parent.insert(idx, child)
                    idx += 1
                parent.remove(hl)

            ppr = p_el.find(ppr_tag)
            if ppr is not None:
                ppr_rpr = ppr.find(rpr_tag)
                if ppr_rpr is not None:
                    ppr.remove(ppr_rpr)

            for run in para.runs:
                rpr = run._r.find(rpr_tag)
                if rpr is not None:
                    for child in list(rpr):
                        if child.tag in _STRIP_RPR:
                            rpr.remove(child)

    strip_run_overrides(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                strip_run_overrides(cell.paragraphs)


def _clean_runs(doc: Document) -> None:
    """
    Text-level cleanup:
    - Strip leading bullet character (U+2022 + space) from paragraph starts.
    - Replace tilde operator (U+223C) with standard tilde (~).
    - Normalize lone "X" in table cells to lowercase "x".
    """
    BULLET = "\u2022"
    TILDE_OP = "\u223c"

    def process_paragraphs(paragraphs):
        for para in paragraphs:
            for run in para.runs:
                if run.text.startswith(BULLET + " "):
                    run.text = run.text[2:]
                elif run.text.startswith(BULLET):
                    run.text = run.text[1:]
                break

            for run in para.runs:
                if TILDE_OP in run.text:
                    run.text = run.text.replace(TILDE_OP, "~")

    process_paragraphs(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                process_paragraphs(cell.paragraphs)
                if cell.text.strip() == "X":
                    for para in cell.paragraphs:
                        for run in para.runs:
                            run.text = run.text.replace("X", "x")


_FOOTNOTES_REL = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
)
_ENDNOTES_REL = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
)
_SKIP_FOOTNOTE_TYPES = {"separator", "continuationSeparator", "continuationNotice"}


def extract_footnotes(source_path: str, output_path: str) -> int:
    """
    Extract all footnotes/endnotes from source_path into a new .docx.
    Returns the number of notes written.
    """
    if not os.path.isfile(source_path):
        raise FileNotFoundError(f"Source file not found: {source_path}")

    doc = Document(source_path)
    out_doc = Document()
    body = out_doc.element.body
    count = 0

    default_p = body.find(qn("w:p"))
    if default_p is not None:
        body.remove(default_p)

    fn_tag    = qn("w:footnote")
    en_tag    = qn("w:endnote")
    id_attr   = qn("w:id")
    type_attr = qn("w:type")
    t_tag     = qn("w:t")
    ref_tags  = {qn("w:footnoteRef"), qn("w:endnoteRef")}
    rpr_tag   = qn("w:rPr")
    keep_rpr  = {qn("w:i"), qn("w:iCs")}

    def _copy_note_paragraphs(note_el):
        for p_el in note_el.findall(qn("w:p")):
            p_copy = copy.deepcopy(p_el)
            for r_el in p_copy.findall(qn("w:r")):
                non_rpr = [c for c in list(r_el) if c.tag != rpr_tag]
                if len(non_rpr) == 1 and non_rpr[0].tag in ref_tags:
                    r_el.getparent().remove(r_el)
            remaining = p_copy.findall(qn("w:r"))
            if remaining:
                t_el = remaining[0].find(t_tag)
                if t_el is not None and t_el.text:
                    stripped = t_el.text.lstrip(" \t\xa0")
                    if stripped:
                        t_el.text = stripped
                    else:
                        remaining[0].getparent().remove(remaining[0])
            for r_el in p_copy.findall(qn("w:r")):
                rpr = r_el.find(rpr_tag)
                if rpr is not None:
                    for child in list(rpr):
                        if child.tag not in keep_rpr:
                            rpr.remove(child)
                    if len(rpr) == 0:
                        r_el.remove(rpr)
            body.append(p_copy)

    def _extract_from_part(rel_uri, note_tag, label_prefix):
        nonlocal count
        try:
            part = doc.part.part_related_by(rel_uri)
            part_xml = etree.fromstring(part.blob)
        except (KeyError, Exception):
            return
        for note in part_xml.findall(note_tag):
            if note.get(type_attr) in _SKIP_FOOTNOTE_TYPES:
                continue
            note_id = note.get(id_attr, "?")
            body.append(_make_paragraph_element(f"{{{{{label_prefix}:{note_id}}}}}"))
            _copy_note_paragraphs(note)
            body.append(_make_paragraph_element(""))
            count += 1

    _extract_from_part(_FOOTNOTES_REL, fn_tag, "fn")
    _extract_from_part(_ENDNOTES_REL,  en_tag, "en")

    if count == 0:
        return 0

    out_doc.save(output_path)
    return count


def export_footnotes_txt(source_path: str, output_path: str) -> int:
    """
    Write a tab-separated .txt mapping footnote/endnote IDs to plain text.
    Returns the number of entries written.
    """
    if not os.path.isfile(source_path):
        raise FileNotFoundError(f"Source file not found: {source_path}")

    doc = Document(source_path)

    t_tag     = qn("w:t")
    rpr_tag   = qn("w:rPr")
    fn_tag    = qn("w:footnote")
    en_tag    = qn("w:endnote")
    id_attr   = qn("w:id")
    type_attr = qn("w:type")
    ref_tags  = {qn("w:footnoteRef"), qn("w:endnoteRef")}

    entries: list[tuple[str, str]] = []

    def _note_plain_text(note_el) -> str:
        paragraphs = []
        for p_el in note_el.findall(qn("w:p")):
            para_text = ""
            for r_el in p_el.findall(qn("w:r")):
                non_rpr = [c for c in list(r_el) if c.tag != rpr_tag]
                if len(non_rpr) == 1 and non_rpr[0].tag in ref_tags:
                    continue
                for t_el in r_el.findall(t_tag):
                    if t_el.text:
                        para_text += t_el.text
            stripped = para_text.lstrip(" \t\xa0")
            if stripped:
                paragraphs.append(stripped)
        return "\r".join(paragraphs)

    def _collect_from_part(rel_uri, note_tag, label_prefix):
        try:
            part = doc.part.part_related_by(rel_uri)
            part_xml = etree.fromstring(part.blob)
        except (KeyError, Exception):
            return
        for note in part_xml.findall(note_tag):
            if note.get(type_attr) in _SKIP_FOOTNOTE_TYPES:
                continue
            note_id = note.get(id_attr, "?")
            text = _note_plain_text(note)
            if text:
                entries.append((f"{label_prefix}:{note_id}", text))

    _collect_from_part(_FOOTNOTES_REL, fn_tag, "fn")
    _collect_from_part(_ENDNOTES_REL,  en_tag, "en")

    if not entries:
        return 0

    with open(output_path, "w", encoding="utf-8") as fh:
        for key, text in entries:
            fh.write(f"{key}\t{text}\n")

    return len(entries)


def main(argv: list[str]) -> None:
    if len(argv) < 2:
        print("Usage: python clean_docx.py INPUT_DOCX")
        sys.exit(1)

    source_path = argv[1]
    base, ext = os.path.splitext(source_path)
    clean_path         = base + "_clean"     + ext
    footnotes_path     = base + "_footnotes" + ext
    footnotes_txt_path = base + "_footnotes" + ".txt"

    # Footnotes extracted from original (unmodified) source
    try:
        n = extract_footnotes(source_path, footnotes_path)
        if n:
            print(f"Footnotes docx: {footnotes_path} ({n} notes)")
        else:
            print("Footnotes docx: none found")
    except Exception as e:
        print(f"Warning: could not extract footnotes docx: {e}", file=sys.stderr)

    try:
        n_txt = export_footnotes_txt(source_path, footnotes_txt_path)
        if n_txt:
            print(f"Footnotes txt:  {footnotes_txt_path} ({n_txt} notes)")
    except Exception as e:
        print(f"Warning: could not export footnotes txt: {e}", file=sys.stderr)

    # Apply all pre-processing and save with tables intact
    try:
        doc = Document(source_path)
        _strip_field_codes(doc)
        _mark_superscripts(doc)       # must run before _strip_character_styles strips rStyle
        _strip_character_styles(doc)
        _clean_runs(doc)
        _extract_description_rows(doc)
        doc.save(clean_path)
        print(f"Clean file: {clean_path}")
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main(sys.argv)
