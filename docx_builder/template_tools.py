"""Generic DOCX template instrumentation helpers.

Library of reusable building blocks for writing a per-template prep script
(the ``template_prep.py`` a project ships next to its draft).  A prep script
uses these helpers to turn an organisation's official DOCX template into a
Jinja-instrumented template that satisfies the engine's template contract
(see docs/TEMPLATE_CONTRACT.md).

Nothing in this module knows about any specific template: heading texts,
style IDs, table layouts and TOC field instructions are all passed in by
the caller.
"""
from __future__ import annotations
from docx.oxml import OxmlElement
from lxml import etree

NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS_XML = 'http://www.w3.org/XML/1998/namespace'


# ---------------------------------------------------------------------------
# Low-level XML helpers
# ---------------------------------------------------------------------------

def para_style_id(el: etree._Element) -> str:
    """Return the w:pStyle val for a <w:p>, or 'Normal' if absent."""
    ps = el.find(f'.//{{{NS_W}}}pStyle')
    return ps.get(f'{{{NS_W}}}val', 'Normal') if ps is not None else 'Normal'


def para_text(el: etree._Element) -> str:
    return ''.join(t.text or '' for t in el.findall(f'.//{{{NS_W}}}t'))


def find_heading(children: list[etree._Element], text: str) -> etree._Element | None:
    """Find the first Heading paragraph whose text contains `text` (case-insensitive)."""
    for child in children:
        if child.tag == f'{{{NS_W}}}p':
            sid = para_style_id(child)
            if sid.startswith('Heading') and text.lower() in para_text(child).lower():
                return child
    return None


def make_para(style_id: str, text: str, align: str | None = None) -> etree._Element:
    """Build a minimal <w:p> with one run.

    Args:
        style_id: DOCX style ID (e.g. 'Normal', 'Heading1').
        text:     Run text content.
        align:    Optional paragraph alignment ('center', 'left', 'right', 'both').
    """
    p = OxmlElement('w:p')
    pPr = OxmlElement('w:pPr')
    pStyle = OxmlElement('w:pStyle')
    pStyle.set(f'{{{NS_W}}}val', style_id)
    pPr.append(pStyle)
    if align:
        jc = OxmlElement('w:jc')
        jc.set(f'{{{NS_W}}}val', align)
        pPr.append(jc)
    p.append(pPr)
    r = OxmlElement('w:r')
    t = OxmlElement('w:t')
    t.text = text
    t.set(f'{{{NS_XML}}}space', 'preserve')
    r.append(t)
    p.append(r)
    return p


def set_cell_text(cell, text: str) -> None:
    """Replace all runs in the first paragraph of a table cell."""
    para = cell.paragraphs[0]
    p_el = para._p
    # Remove existing runs
    for r in p_el.findall(f'{{{NS_W}}}r'):
        p_el.remove(r)
    # Add a single run
    r = OxmlElement('w:r')
    t = OxmlElement('w:t')
    t.text = text
    t.set(f'{{{NS_XML}}}space', 'preserve')
    r.append(t)
    p_el.append(r)


# ---------------------------------------------------------------------------
# Cleanup helpers
# ---------------------------------------------------------------------------

def remove_range(body: etree._Element, from_el: etree._Element, to_el: etree._Element) -> None:
    """Remove all body children from from_el (inclusive) to to_el (exclusive)."""
    removing = False
    for child in list(body):
        if child is from_el:
            removing = True
        if child is to_el:
            break
        if removing:
            body.remove(child)


def remove_styled_and_empty_between(
    body: etree._Element,
    start_el: etree._Element,
    end_el: etree._Element,
    style_ids: tuple[str, ...] = (),
) -> None:
    """Remove paragraphs with a style in `style_ids`, and empty Normal
    paragraphs, between two elements (both exclusive)."""
    in_range = False
    to_remove: list[etree._Element] = []
    for child in list(body):
        if child is start_el:
            in_range = True
            continue
        if child is end_el:
            break
        if not in_range:
            continue
        if child.tag != f'{{{NS_W}}}p':
            continue
        sid = para_style_id(child)
        text = para_text(child).strip()
        if sid in style_ids or (sid == 'Normal' and not text):
            to_remove.append(child)
    for el in to_remove:
        body.remove(el)


# ---------------------------------------------------------------------------
# Table instrumentation
# ---------------------------------------------------------------------------

def instrument_table(
    table,
    for_tag: str,
    cell_vars: list[str],
    endfor_tag: str,
    data_row_start: int = 2,  # first data row index (0-based)
) -> None:
    """Instrument a table for docxtpl row looping.

    Expects the table to have at least (data_row_start + 3) rows:
      rows 0..data_row_start-1  — header rows (kept as-is)
      row  data_row_start       — becomes the {% tr for %} marker row
      row  data_row_start+1     — becomes the template row with {{ vars }}
      row  data_row_start+2     — becomes the {% tr endfor %} marker row
      rows data_row_start+3+    — deleted
    """
    rows = table.rows
    for_row = rows[data_row_start]
    tmpl_row = rows[data_row_start + 1]
    end_row = rows[data_row_start + 2]

    # For-marker row: tag in first cell, clear the rest
    set_cell_text(for_row.cells[0], for_tag)
    for cell in for_row.cells[1:]:
        set_cell_text(cell, '')

    # Template row: variable reference in each cell
    for ci, var in enumerate(cell_vars):
        if ci < len(tmpl_row.cells):
            set_cell_text(tmpl_row.cells[ci], var)

    # Endfor-marker row
    set_cell_text(end_row.cells[0], endfor_tag)
    for cell in end_row.cells[1:]:
        set_cell_text(cell, '')

    # Delete surplus rows
    tbl_el = table._tbl
    tr_els = tbl_el.findall(f'{{{NS_W}}}tr')
    for tr in tr_els[data_row_start + 3:]:
        tbl_el.remove(tr)


# ---------------------------------------------------------------------------
# Word-native TOC fields (used for List of Figures / List of Tables)
# ---------------------------------------------------------------------------

def make_toc_field_para(instr: str, para_style: str = 'TOC1') -> etree._Element:
    """Build a <w:p> containing a Word TOC field that Word will expand on open.

    The field is marked dirty so Word regenerates it automatically.

    Args:
        instr:      Field instruction, e.g. r'TOC \\h \\z \\t "Caption,1"'.
        para_style: Paragraph style for the field paragraph.
    """
    p = OxmlElement('w:p')
    pPr = OxmlElement('w:pPr')
    pStyle = OxmlElement('w:pStyle')
    pStyle.set(f'{{{NS_W}}}val', para_style)
    pPr.append(pStyle)
    p.append(pPr)

    r_begin = OxmlElement('w:r')
    fc_begin = OxmlElement('w:fldChar')
    fc_begin.set(f'{{{NS_W}}}fldCharType', 'begin')
    fc_begin.set(f'{{{NS_W}}}dirty', 'true')
    r_begin.append(fc_begin)
    p.append(r_begin)

    r_instr = OxmlElement('w:r')
    it = OxmlElement('w:instrText')
    it.set(f'{{{NS_XML}}}space', 'preserve')
    it.text = f' {instr} '
    r_instr.append(it)
    p.append(r_instr)

    r_sep = OxmlElement('w:r')
    fc_sep = OxmlElement('w:fldChar')
    fc_sep.set(f'{{{NS_W}}}fldCharType', 'separate')
    r_sep.append(fc_sep)
    p.append(r_sep)

    r_placeholder = OxmlElement('w:r')
    t_placeholder = OxmlElement('w:t')
    t_placeholder.set(f'{{{NS_XML}}}space', 'preserve')
    t_placeholder.text = 'Right-click to update field.'
    r_placeholder.append(t_placeholder)
    p.append(r_placeholder)

    r_end = OxmlElement('w:r')
    fc_end = OxmlElement('w:fldChar')
    fc_end.set(f'{{{NS_W}}}fldCharType', 'end')
    r_end.append(fc_end)
    p.append(r_end)

    return p


def insert_toc_sections(
    anchor_el: etree._Element,
    sections: list[tuple[str, str]],
    heading_style: str = 'Heading2',
    toc_para_style: str = 'TOC1',
) -> None:
    """Insert (heading, TOC field) pairs immediately after `anchor_el`.

    Args:
        anchor_el: Element after which the sections are inserted, in order.
        sections:  List of (heading_text, toc_field_instruction).
    """
    ref = anchor_el
    for heading_text, toc_instr in sections:
        h = make_para(heading_style, heading_text)
        ref.addnext(h)
        ref = h
        toc_p = make_toc_field_para(toc_instr, para_style=toc_para_style)
        ref.addnext(toc_p)
        ref = toc_p


# ---------------------------------------------------------------------------
# Content-loop insertion
# ---------------------------------------------------------------------------

def insert_content_loop(
    sect_pr: etree._Element,
    entries: list[tuple],
) -> None:
    """Insert content-loop paragraphs immediately before the final sectPr.

    Each entry is (text, style_id, align, numId?, ilvl?):
      text      — paragraph text ({%p ... %} control tags or {{ item.* }} vars)
      style_id  — DOCX style ID for the paragraph
      align     — optional alignment passed to make_para
      numId     — optional w:numPr/w:numId value (list numbering)
      ilvl      — optional w:numPr/w:ilvl value (nested list level)

    addprevious(sect_pr) places each new paragraph directly before sectPr,
    so iterating in forward order produces the correct document order.
    """
    for entry in entries:
        text, style_id, align = entry[0], entry[1], entry[2]
        num_id = entry[3] if len(entry) > 3 else None
        ilvl = entry[4] if len(entry) > 4 else None
        p = make_para(style_id, text, align=align)
        if num_id is not None or ilvl is not None:
            pPr = p.find(f'{{{NS_W}}}pPr')
            if pPr is not None:
                numPr = OxmlElement('w:numPr')
                if num_id is not None:
                    numIdEl = OxmlElement('w:numId')
                    numIdEl.set(f'{{{NS_W}}}val', str(num_id))
                    numPr.append(numIdEl)
                if ilvl is not None:
                    ilvlEl = OxmlElement('w:ilvl')
                    ilvlEl.set(f'{{{NS_W}}}val', str(ilvl))
                    numPr.append(ilvlEl)
                pPr.append(numPr)
        sect_pr.addprevious(p)


def force_decimal_numbering(doc, numid: int, lvl_text: str = '%1.') -> None:
    """Force a numId's level-0 list definition to plain decimal (1., 2., ...).

    Numbered-list content-loop items are assigned this numId directly, but
    the abstractNum it references in the official template carries whatever
    format Word last used there (bullet, lettered, "Annex %1 -", ...) — it
    must be normalized per template, not assumed.
    """
    numbering = doc.part.numbering_part._element
    abstract_id = None
    for num in numbering.findall(f'{{{NS_W}}}num'):
        if num.get(f'{{{NS_W}}}numId') == str(numid):
            ref = num.find(f'{{{NS_W}}}abstractNumId')
            abstract_id = ref.get(f'{{{NS_W}}}val') if ref is not None else None
            break
    if abstract_id is None:
        return
    for an in numbering.findall(f'{{{NS_W}}}abstractNum'):
        if an.get(f'{{{NS_W}}}abstractNumId') != abstract_id:
            continue
        lvl0 = an.find(f'.//{{{NS_W}}}lvl[@{{{NS_W}}}ilvl="0"]')
        if lvl0 is None:
            return
        num_fmt = lvl0.find(f'{{{NS_W}}}numFmt')
        if num_fmt is None:
            num_fmt = OxmlElement('w:numFmt')
            lvl0.append(num_fmt)
        num_fmt.set(f'{{{NS_W}}}val', 'decimal')
        lvl_text_el = lvl0.find(f'{{{NS_W}}}lvlText')
        if lvl_text_el is None:
            lvl_text_el = OxmlElement('w:lvlText')
            lvl0.append(lvl_text_el)
        lvl_text_el.set(f'{{{NS_W}}}val', lvl_text)
        return


def get_or_make_spacing(pPr_el: etree._Element) -> etree._Element:
    """Get or create the w:spacing child of a pPr element."""
    sp = pPr_el.find(f'{{{NS_W}}}spacing')
    if sp is None:
        sp = OxmlElement('w:spacing')
        pPr_el.append(sp)
    return sp
