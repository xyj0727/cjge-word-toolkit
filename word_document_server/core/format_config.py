"""
Default document format of this server: 《岩土工程学报》 (Chinese Journal of Geotechnical Engineering, CJGE).

Every place where the server chooses a format by itself (new documents, heading styles, the heading
fallback, new tables) reads it from here. Explicit tool arguments (font_name, font_size, bold, ...) still
win. MathType equations are not touched.

Word sizes: 二号 22 pt, 四号 14 pt, 五号 10.5 pt, 小五 9 pt, 六号 7.5 pt.
"""
from docx.enum.section import WD_ORIENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm

LATIN = "Times New Roman"

# A4; 版心 17.8 cm x 25.4 cm. The body runs in two columns of 8.53 cm with a 0.74 cm gap, the front
# matter (题名 … Key words) in one column. A new document has no front matter yet, so it starts in one
# column; the docx-editor-cn skill (scripts/apply_format.py) inserts the column break later.
PAGE = {
    "width_cm": 21.0, "height_cm": 29.7,
    "left_cm": 1.65, "right_cm": 1.53, "top_cm": 2.15, "bottom_cm": 2.15,
    "header_cm": 1.5, "footer_cm": 1.0,
    "new_document_columns": 1, "body_columns": 2, "column_space_cm": 0.74,
}

# Paragraph styles: east = Chinese font, size in pt, line = fixed line spacing in pt.
STYLES = {
    "Normal":    {"east": "SimSun", "size": 10.5, "align": "both", "first_chars": 2, "line": 15.6},
    "Title":     {"east": "SimHei", "size": 22, "bold": False, "align": "center", "line": 30, "after": 8},
    "Heading 1": {"east": "SimHei", "size": 14, "bold": False, "align": "left", "line": 18, "before": 15},
    "Heading 2": {"east": "SimHei", "size": 10.5, "bold": False, "align": "left", "line": 15.6},
    "Heading 3": {"east": "SimSun", "size": 10.5, "bold": False, "align": "both", "first_chars": 2, "line": 15.6},
    "Caption":   {"east": "SimHei", "size": 9, "bold": False, "align": "center", "line": 15.6},
}

# Heading sizes used when a heading has to be built without a heading style.
HEADING_FALLBACK_SIZES = {1: 14, 2: 10.5, 3: 10.5}
HEADING_FALLBACK_BOLD = False

# 三线表: top/bottom 0.75 pt, header rule 0.5 pt, no vertical rules; 小五 text, fixed 14 pt.
TABLE = {"top_bottom_pt": 0.75, "header_rule_pt": 0.5, "east": "SimSun", "size": 9, "line": 14}

_ALIGN = {"left": "left", "center": "center", "right": "right", "both": "both"}
_PPR_ORDER = ["w:pStyle", "w:keepNext", "w:keepLines", "w:pageBreakBefore", "w:framePr", "w:widowControl",
              "w:numPr", "w:suppressLineNumbers", "w:pBdr", "w:shd", "w:tabs", "w:suppressAutoHyphens",
              "w:kinsoku", "w:wordWrap", "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE",
              "w:autoSpaceDN", "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
              "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc", "w:textDirection",
              "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr",
              "w:sectPr", "w:pPrChange"]
_RPR_ORDER = ["w:rStyle", "w:rFonts", "w:b", "w:bCs", "w:i", "w:iCs", "w:caps", "w:smallCaps", "w:strike",
              "w:dstrike", "w:outline", "w:shadow", "w:emboss", "w:imprint", "w:noProof", "w:snapToGrid",
              "w:vanish", "w:webHidden", "w:color", "w:spacing", "w:w", "w:kern", "w:position", "w:sz",
              "w:szCs", "w:highlight", "w:u", "w:effect", "w:bdr", "w:shd", "w:fitText", "w:vertAlign",
              "w:rtl", "w:cs", "w:em", "w:lang", "w:eastAsianLayout", "w:specVanish", "w:oMath"]


def _child(parent, tag):
    el = parent.find(qn(tag))
    if el is None:
        el = OxmlElement(tag)
        parent.append(el)
    return el


def _reorder(el, order):
    rank = {qn(t): i for i, t in enumerate(order)}
    for c in sorted(list(el), key=lambda c: rank.get(c.tag, len(order))):
        el.remove(c)
        el.append(c)


def set_run_format(rPr, east, size=None, bold=None):
    """Chinese font `east`, Latin font Times New Roman, optional size (pt) and bold."""
    rf = _child(rPr, "w:rFonts")
    for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        rf.attrib.pop(qn(attr), None)
    for attr in ("w:ascii", "w:hAnsi", "w:cs"):
        rf.set(qn(attr), LATIN)
    rf.set(qn("w:eastAsia"), east)
    if size is not None:
        half = str(int(round(size * 2)))
        _child(rPr, "w:sz").set(qn("w:val"), half)
        _child(rPr, "w:szCs").set(qn("w:val"), half)
    if bold is not None:
        for tag in ("w:b", "w:bCs"):
            _child(rPr, tag).set(qn("w:val"), "1" if bold else "0")
    _reorder(rPr, _RPR_ORDER)


def set_paragraph_format(pPr, spec):
    if spec.get("align"):
        _child(pPr, "w:jc").set(qn("w:val"), _ALIGN[spec["align"]])
    sp = _child(pPr, "w:spacing")
    for attr in list(sp.attrib):
        del sp.attrib[attr]
    sp.set(qn("w:before"), str(int(round(spec.get("before", 0) * 20))))
    sp.set(qn("w:after"), str(int(round(spec.get("after", 0) * 20))))
    sp.set(qn("w:line"), str(int(round(spec["line"] * 20))))
    sp.set(qn("w:lineRule"), "exact")
    ind = _child(pPr, "w:ind")
    for attr in list(ind.attrib):
        del ind.attrib[attr]
    chars = spec.get("first_chars", 0)
    ind.set(qn("w:firstLineChars"), str(int(chars * 100)))
    ind.set(qn("w:firstLine"), str(int(round(chars * spec["size"] * 20))))
    _reorder(pPr, _PPR_ORDER)


def apply_style_format(style, spec):
    """Write one STYLES entry into a paragraph style (fonts, size, colour, spacing, indent)."""
    el = style.element
    rPr = _child(el, "w:rPr")
    set_run_format(rPr, spec["east"], spec["size"], spec.get("bold"))
    color = rPr.find(qn("w:color"))
    if color is not None:
        rPr.remove(color)
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "000000")
    rPr.append(color)
    _reorder(rPr, _RPR_ORDER)
    pPr = el.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        rPr.addprevious(pPr)
    set_paragraph_format(pPr, spec)


def apply_named_styles(doc, names=None):
    for name, spec in STYLES.items():
        if names is not None and name not in names:
            continue
        try:
            style = doc.styles[name]
        except KeyError:
            continue
        apply_style_format(style, spec)


def apply_document_defaults(doc):
    """Document defaults, styles and page setup for a new document."""
    rpr_default = doc.styles.element.find(f"{qn('w:docDefaults')}/{qn('w:rPrDefault')}/{qn('w:rPr')}")
    if rpr_default is not None:
        normal = STYLES["Normal"]
        set_run_format(rpr_default, normal["east"], normal["size"])
        _child(rpr_default, "w:lang").set(qn("w:eastAsia"), "zh-CN")
        _reorder(rpr_default, _RPR_ORDER)
    apply_named_styles(doc)
    for section in doc.sections:
        section.orientation = WD_ORIENT.PORTRAIT
        section.page_width = Cm(PAGE["width_cm"])
        section.page_height = Cm(PAGE["height_cm"])
        section.left_margin = Cm(PAGE["left_cm"])
        section.right_margin = Cm(PAGE["right_cm"])
        section.top_margin = Cm(PAGE["top_cm"])
        section.bottom_margin = Cm(PAGE["bottom_cm"])
        section.header_distance = Cm(PAGE["header_cm"])
        section.footer_distance = Cm(PAGE["footer_cm"])
        cols = _child(section._sectPr, "w:cols")
        cols.set(qn("w:num"), str(PAGE["new_document_columns"]))
        cols.set(qn("w:space"), str(int(round(PAGE["column_space_cm"] / 2.54 * 1440))))


def _border(parent, side, pt):
    el = _child(parent, f"w:{side}")
    for attr in list(el.attrib):
        del el.attrib[attr]
    if pt:
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(int(round(pt * 8))))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
    else:
        el.set(qn("w:val"), "nil")


def apply_three_line_table(table):
    """三线表 for a table created by the server: rules, centred 小五 cell text, fixed 14 pt."""
    tbl = table._tbl
    tblPr = tbl.tblPr
    style = tblPr.find(qn("w:tblStyle"))
    if style is not None:
        tblPr.remove(style)
    _child(tblPr, "w:jc").set(qn("w:val"), "center")
    borders = _child(tblPr, "w:tblBorders")
    for side, pt in (("top", TABLE["top_bottom_pt"]), ("left", 0), ("bottom", TABLE["top_bottom_pt"]),
                     ("right", 0), ("insideH", 0), ("insideV", 0)):
        _border(borders, side, pt)
    rows = tbl.findall(qn("w:tr"))
    for r_index, tr in enumerate(rows):
        for tc in tr.findall(qn("w:tc")):
            if r_index == 0 and len(rows) > 1:
                tcPr = tc.get_or_add_tcPr()
                _border(_child(tcPr, "w:tcBorders"), "bottom", TABLE["header_rule_pt"])
            for p in tc.findall(qn("w:p")):
                set_table_paragraph_format(p)


def set_table_paragraph_format(p):
    pPr = p.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        p.insert(0, pPr)
    set_paragraph_format(pPr, {"align": "center", "line": TABLE["line"], "size": TABLE["size"]})
    for r in p.findall(qn("w:r")):
        rPr = r.find(qn("w:rPr"))
        if rPr is None:
            rPr = OxmlElement("w:rPr")
            r.insert(0, rPr)
        set_run_format(rPr, TABLE["east"], TABLE["size"])
