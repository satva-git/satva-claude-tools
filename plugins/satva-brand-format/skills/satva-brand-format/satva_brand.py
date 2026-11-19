"""Satva brand helpers for .docx (python-docx) and .pptx (python-pptx).
Usage:  import satva_brand as sb
  Slides: prs = sb.new_deck(); sb.cover(prs, "Title", "Subtitle"); s = sb.content(prs, "Slide title"); ...; prs.save(path)
  Docs:   doc = sb.new_doc(); sb.doc_cover(doc, "Business Requirement Document", info_rows, amendments); sb.h1(doc,"Acknowledgment:") ...; doc.save(path)
  PDF:    sb.to_pdf(path)  (LibreOffice, keeps header/footer/theme)
"""
import base64, io, os, subprocess, copy
try:
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.oxml.ns import qn as qn_p
    from docx import Document
    from docx.shared import Pt as DPt, Inches as DIn, RGBColor as DRGB, Emu as DEmu, Twips
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
except ImportError:          # run `python satva_brand.py` to see what is missing
    class _Missing:
        def __getattr__(self, name): return None
    MSO_SHAPE = PP_ALIGN = MSO_ANCHOR = _Missing()

# ---------------- brand tokens ----------------
COMPANY_DECK = "Satva Solutions"
COMPANY_DOC = "Satva Technolabs Pvt. Ltd."
PHONE = "+1-415-802-0027"
EMAIL = "sales@satvasolutions.com"
# slides
BLUE = "0094D5"; BLUE_DARK = "00679A"; CARD = "EFF7FC"; MUTED = "5F6B76"; INK = "000000"; WHITE = "FFFFFF"
F_HEAD = "Lexend SemiBold"; F_BODY = "Mulish"
# documents
DOC_BLUE = "2D74B5"; DOC_BLUE2 = "3D85C6"; LINK = "1155CC"; F_DOC = "Mulish"

import tempfile, shutil, sys
_ASSET_DIR = os.path.join(tempfile.gettempdir(), "satva_brand_assets")
def asset(name):
    os.makedirs(_ASSET_DIR, exist_ok=True)
    p = os.path.join(_ASSET_DIR, name + ".png")
    if os.path.exists(p): return p
    if name in ("logo_white", "watermark_slide"):          # derived from the blue logo
        from PIL import Image as _I
        im = _I.open(asset("logo_blue")).convert("RGBA")
        px = im.load()
        for yy in range(im.height):
            for xx in range(im.width):
                r, g, b, a = px[xx, yy]
                px[xx, yy] = (255, 255, 255, a) if name == "logo_white" else (228, 228, 228, int(a * 0.4))
        if name == "watermark_slide": im = im.resize((im.width * 3, im.height * 3), _I.LANCZOS)
        im.save(p)
    else:
        open(p, "wb").write(base64.b64decode(ASSETS[name]))
    return p

def _rgb(h): return RGBColor.from_string(h)

# =============================== SLIDES ===============================
W, H = 10.0, 5.625

def new_deck():
    prs = Presentation(); prs.slide_width = Inches(W); prs.slide_height = Inches(H)
    return prs

def _blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])

def _bg(slide, hexcol):
    f = slide.background.fill; f.solid(); f.fore_color.rgb = _rgb(hexcol)

def _rect(slide, x, y, w, h, fill, shape=MSO_SHAPE.RECTANGLE):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid(); s.fill.fore_color.rgb = _rgb(fill); s.line.fill.background()
    st = s._element.find(qn_p("p:style"))
    if st is not None: s._element.remove(st)          # no theme shadow/outline
    return s

def text(slide, x, y, w, h, txt, size=14, color=INK, font=F_BODY, bold=False, align=PP_ALIGN.LEFT,
         anchor=MSO_ANCHOR.TOP, bullets=False):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.05); tf.margin_top = tf.margin_bottom = Inches(0.03)
    lines = txt if isinstance(txt, list) else [txt]
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run(); r.text = ("•  " + line) if bullets else line
        r.font.size = Pt(size); r.font.name = font; r.font.bold = bold; r.font.color.rgb = _rgb(color)
        if bullets: p.space_after = Pt(6)
    return tb

def _contact(slide):
    text(slide, 5.6, 5.20, 4.0, 0.3, f"{PHONE}   |   {EMAIL}", size=9, color=WHITE,
         font=F_DOC, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

def _blue_slide(prs):
    s = _blank(prs); _bg(s, BLUE)
    s.shapes.add_picture(asset("logo_white"), Inches(7.76), Inches(0.40), Inches(1.85))
    _contact(s); return s

def cover(prs, title, subtitle=""):
    s = _blue_slide(prs)
    text(s, 0.47, 1.18, 9.0, 0.85, title, 33, WHITE, F_HEAD, True, anchor=MSO_ANCHOR.MIDDLE)
    if subtitle: text(s, 0.47, 2.17, 9.16, 0.36, subtitle, 20, WHITE, F_HEAD)
    _rect(s, 0.53, 2.98, 0.78, 0.025, WHITE); return s

def section(prs, title):
    s = _blue_slide(prs)
    text(s, 0.47, 2.30, 9.0, 0.6, title, 20, WHITE, F_HEAD, anchor=MSO_ANCHOR.MIDDLE)
    _rect(s, 0.53, 3.05, 0.78, 0.025, WHITE); return s

def closing(prs, title="Thank You !"):
    s = _blue_slide(prs)
    text(s, 0.47, 2.00, 9.0, 0.9, title, 40, WHITE, F_HEAD, True, anchor=MSO_ANCHOR.MIDDLE); return s

def content(prs, title):
    s = _blank(prs); _bg(s, WHITE)
    s.shapes.add_picture(asset("watermark_slide"), Inches(1.4), Inches(1.75), Inches(7.6))   # faint centre mark
    _rect(s, 0.0, 0.185, 0.107, 0.36, BLUE)                                                  # left tab
    s.shapes.add_picture(asset("logo_blue"), Inches(8.35), Inches(0.23), Inches(1.41))        # top-right logo
    s.shapes.add_picture(asset("swoosh"), Inches(8.242), Inches(5.40), Inches(1.758), Inches(0.225))
    text(s, 0.40, 0.20, 7.8, 0.42, title, 19, BLUE, F_HEAD, True, anchor=MSO_ANCHOR.MIDDLE)
    return s

def stat_cards(slide, cards, y=1.64):
    """cards: list of (big, label, desc) — up to 4, laid out like 'Where We Are'."""
    n = len(cards); gap = 0.21; w = (9.2 - gap * (n - 1)) / n
    for i, (big, label, desc) in enumerate(cards):
        x = 0.40 + i * (w + gap)
        _rect(slide, x, y, w, 2.84, CARD)
        _rect(slide, x + 0.15, y + 0.22, 0.5, 0.5, BLUE, MSO_SHAPE.OVAL)
        text(slide, x + 0.15, y + 0.85, w - 0.3, 0.57, big, 28, BLUE, F_HEAD, True)
        text(slide, x + 0.15, y + 1.44, w - 0.3, 0.47, label, 11, BLUE, F_HEAD, True)
        text(slide, x + 0.15, y + 1.96, w - 0.3, 0.79, desc, 8.5, MUTED, F_BODY)

def table(slide, rows, x=0.40, y=1.0, w=9.2, col_w=None, size=9.5):
    """rows[0] = header. Header: brand blue + white bold; first column bold dark blue."""
    nr, nc = len(rows), len(rows[0])
    gt = slide.shapes.add_table(nr, nc, Inches(x), Inches(y), Inches(w), Inches(0.3 * nr)).table
    if col_w:
        for i, cw in enumerate(col_w): gt.columns[i].width = Inches(cw)
    for r in range(nr):
        for c in range(nc):
            cell = gt.cell(r, c); cell.text = str(rows[r][c])
            cell.fill.solid(); cell.fill.fore_color.rgb = _rgb(BLUE if r == 0 else (WHITE if r % 2 else CARD))
            cell.margin_left = cell.margin_right = Inches(0.06); cell.margin_top = cell.margin_bottom = Inches(0.03)
            for p in cell.text_frame.paragraphs:
                for run in p.runs:
                    run.font.name = F_BODY; run.font.size = Pt(size)
                    run.font.bold = (r == 0 or c == 0)
                    run.font.color.rgb = _rgb(WHITE if r == 0 else (BLUE_DARK if c == 0 else INK))
    return gt

def image(slide, path, x=0.40, y=0.95, w=None, h=None):
    """Screenshots/visuals sit inside the safe area x 0.40–9.60, y 0.85–5.25."""
    from PIL import Image as _I
    iw, ih = _I.open(path).size; maxw, maxh = w or 9.2, h or 4.3
    sc = min(maxw / iw, maxh / ih)
    return slide.shapes.add_picture(path, Inches(x + (maxw - iw * sc) / 2), Inches(y), Inches(iw * sc), Inches(ih * sc))

# =============================== DOCUMENTS ===============================
def _set_font(run, size=12, bold=None, color=None, font=F_DOC, italic=None):
    run.font.name = font; run.font.size = DPt(size)
    rpr = run._element.get_or_add_rPr(); rf = rpr.find(qn("w:rFonts"))
    if rf is None: rf = OxmlElement("w:rFonts"); rpr.insert(0, rf)
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"): rf.set(qn(a), font)
    if bold is not None: run.font.bold = bold
    if italic is not None: run.font.italic = italic
    if color: run.font.color.rgb = DRGB.from_string(color)

def _style(doc, name, size, bold, color, before, after, indent=0):
    st = doc.styles[name]; st.font.name = F_DOC; st.font.size = DPt(size); st.font.bold = bold
    st.font.color.rgb = DRGB.from_string(color); st.font.italic = False
    rpr = st.element.get_or_add_rPr(); rf = rpr.find(qn("w:rFonts"))
    if rf is None: rf = OxmlElement("w:rFonts"); rpr.insert(0, rf)
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"): rf.set(qn(a), F_DOC)
    for t in ("w:asciiTheme", "w:hAnsiTheme", "w:cstheme", "w:eastAsiaTheme"):
        if rf.get(qn(t)) is not None: del rf.attrib[qn(t)]
    pf = st.paragraph_format; pf.space_before = DPt(before); pf.space_after = DPt(after)
    pf.left_indent = DIn(indent); pf.keep_with_next = True

def new_doc():
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Twips(11909); sec.page_height = Twips(16834)            # A4
    sec.left_margin = sec.right_margin = DIn(1.0)
    sec.top_margin = DIn(0.6); sec.bottom_margin = DIn(0.66)
    sec.header_distance = DIn(0.2); sec.footer_distance = DIn(0.2)
    sec.different_first_page_header_footer = True                            # cover page: no header/footer
    n = doc.styles["Normal"]; n.font.name = F_DOC; n.font.size = DPt(12)
    n.element.get_or_add_rPr()
    _style(doc, "Normal", 12, False, "000000", 0, 6)
    doc.styles["Normal"].paragraph_format.keep_with_next = False
    doc.styles["Normal"].paragraph_format.line_spacing = 1.15
    _style(doc, "Title", 21, True, DOC_BLUE, 6, 12)
    _style(doc, "Heading 1", 16, True, DOC_BLUE, 24, 10)
    _style(doc, "Heading 2", 14, True, DOC_BLUE2, 18, 8, 0.5)
    _style(doc, "Heading 3", 12, True, "000000", 12, 4)
    # Title style in default template has a bottom border — remove it
    ppr = doc.styles["Title"].element.find(qn("w:pPr"))
    if ppr is not None and ppr.find(qn("w:pBdr")) is not None: ppr.remove(ppr.find(qn("w:pBdr")))
    _footer(doc, sec)
    doc._h1 = 0; doc._h2 = 0
    return doc

def _anchor_xml(rid, cx, cy, off_y):
    return (
        '<wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        f'distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="0" behindDoc="1" locked="1" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        '<wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH>'
        f'<wp:positionV relativeFrom="page"><wp:posOffset>{off_y}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
        '<wp:docPr id="901" name="SatvaWatermark"/><wp:cNvGraphicFramePr/>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>'
        '<pic:nvPicPr><pic:cNvPr id="901" name="watermark.png"/><pic:cNvPicPr/></pic:nvPicPr>'
        f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
        '</pic:pic></a:graphicData></a:graphic></wp:anchor>')

def _footer(doc, sec):
    """Pages 2+: faint full-width logo watermark at page bottom + centred page number."""
    from docx.oxml import parse_xml
    ftr = sec.footer; p = ftr.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    rid, _ = ftr.part.get_or_add_image(asset("doc_watermark"))
    cx = int(DIn(6.27)); cy = int(cx * 60 / 204)
    off_y = int(Twips(16834)) - cy - int(DIn(0.35))
    drawing = OxmlElement("w:drawing"); drawing.append(parse_xml(_anchor_xml(rid, cx, cy, off_y)))
    run._r.append(drawing)
    r2 = p.add_run(); _set_font(r2, 10, color="000000")
    for t, v in (("begin", None), ("instr", "PAGE"), ("end", None)):
        if t == "instr":
            it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = v; r2._r.append(it)
        else:
            fc = OxmlElement("w:fldChar"); fc.set(qn("w:fldCharType"), t); r2._r.append(fc)

def _borders(table, color="000000", sz=8):
    tblPr = table._tbl.tblPr; b = OxmlElement("w:tblBorders")
    for e in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{e}"); el.set(qn("w:val"), "single"); el.set(qn("w:sz"), str(sz))
        el.set(qn("w:space"), "0"); el.set(qn("w:color"), color); b.append(el)
    tblPr.append(b)

def _widths(table, widths):
    table.autofit = False
    grid = table._tbl.tblGrid
    for j, gc in enumerate(grid.findall(qn("w:gridCol"))): gc.set(qn("w:w"), str(int(widths[j] * 1440)))
    for row in table.rows:
        for j, c in enumerate(row.cells):
            if j < len(widths): c.width = DIn(widths[j])

def _cell(cell, txt, bold=False, size=11, align=None, fill=None, color="000000"):
    cell.text = ""; p = cell.paragraphs[0]
    if align: p.alignment = align
    p.paragraph_format.space_after = DPt(0)
    if isinstance(txt, tuple):           # ("Label:", " value") -> bold label + normal value
        r = p.add_run(txt[0]); _set_font(r, size, True, color); r = p.add_run(txt[1]); _set_font(r, size, False, color)
    else:
        r = p.add_run(str(txt)); _set_font(r, size, bold, color)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    if fill:
        tcPr = cell._tc.get_or_add_tcPr(); shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear"); shd.set(qn("w:fill"), fill); tcPr.append(shd)

def doc_cover(doc, doc_title, info_rows, amendments, logo=True):
    """info_rows: [[("Name of The Project:", " X")], [("For:", " Y"), ("Prepared By:", " Z")], [("Created On:", " date")]]
    amendments: [["1.0","Name","04/11/2026","Description"], ...] (2 blank rows are appended)."""
    if logo:
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = DPt(24); p.add_run().add_picture(asset("doc_logo"), width=DIn(2.2))
    t = doc.add_paragraph(doc_title, style="Title"); t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tbl = doc.add_table(rows=len(info_rows), cols=2); _borders(tbl); tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, row in enumerate(info_rows):
        if len(row) == 1:
            c = tbl.cell(i, 0).merge(tbl.cell(i, 1)); _cell(c, row[0])
        else:
            _cell(tbl.cell(i, 0), row[0]); _cell(tbl.cell(i, 1), row[1])
    doc.add_paragraph()
    a = doc.add_table(rows=2 + len(amendments) + 2, cols=4); _borders(a); a.alignment = WD_TABLE_ALIGNMENT.CENTER
    _widths(a, [0.75, 1.1, 1.05, 3.4]); _widths(tbl, [3.6, 2.7])
    hc = a.cell(0, 0).merge(a.cell(0, 3)); _cell(hc, "Amendment History", True, 16, WD_ALIGN_PARAGRAPH.CENTER)
    for j, h in enumerate(["Version", "Prepared By", "Date", "Description"]):
        _cell(a.cell(1, j), h, True, 11, WD_ALIGN_PARAGRAPH.CENTER)
    for i, r in enumerate(amendments):
        for j, v in enumerate(r): _cell(a.cell(2 + i, j), v, False, 11, WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_page_break()

def h1(doc, txt):
    doc._h1 += 1; doc._h2 = 0
    return doc.add_paragraph(f"{doc._h1}.  {txt}", style="Heading 1")

def h2(doc, txt):
    doc._h2 += 1
    return doc.add_paragraph(f"{doc._h2}.  {txt}", style="Heading 2")

def para(doc, txt, bold=False):
    p = doc.add_paragraph(); r = p.add_run(txt); _set_font(r, 12, bold); return p

def arrow_head(doc, label):
    """Bold run-in sub-heading: '→ Invoice Payment:'"""
    p = doc.add_paragraph(); p.paragraph_format.space_before = DPt(10); p.paragraph_format.space_after = DPt(2)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(f"→ {label}"); _set_font(r, 12, True); return p

def arrow(doc, txt):
    """'→ statement' paragraphs used under Heading 2 sections."""
    p = doc.add_paragraph(); r = p.add_run(f"→ {txt}"); _set_font(r, 12); return p

def bullet(doc, txt, level=0):
    p = doc.add_paragraph(); pf = p.paragraph_format
    pf.left_indent = DIn(0.5 + 0.25 * level); pf.first_line_indent = DIn(-0.25); pf.space_after = DPt(2)
    r = p.add_run(("\u2022" if level == 0 else "\u2013") + "\t" + txt); _set_font(r, 12)   # Mulish has these; ● would print oversized
    pf.tab_stops.add_tab_stop(DIn(0.5 + 0.25 * level)); return p

def figure(doc, path, width_in=6.0, caption=None):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(path, width=DIn(width_in))
    if caption:
        c = doc.add_paragraph(); c.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = c.add_run(caption); _set_font(r, 10, italic=True, color="666666")

def doc_table(doc, rows, widths=None, header_fill=None):
    """Body tables: black 1pt grid, bold centred header row (header_fill optional, e.g. 'DEEAF6')."""
    t = doc.add_table(rows=len(rows), cols=len(rows[0])); _borders(t); t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            _cell(t.cell(i, j), v, i == 0, 11, WD_ALIGN_PARAGRAPH.CENTER if i == 0 else None, header_fill if i == 0 else None)
    if widths: _widths(t, widths)
    doc.add_paragraph(); return t

# =============================== SPREADSHEETS ===============================
def style_sheet(ws, header_row=1, freeze=True):
    """Satva look for an openpyxl worksheet: blue Mulish header, tinted rows, thin borders, sized columns."""
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
    from openpyxl.utils import get_column_letter
    thin = Side(style="thin", color="DFE9F1"); border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for r, row in enumerate(ws.iter_rows(min_row=header_row, max_row=ws.max_row), start=header_row):
        for c in row:
            c.border = border
            if r == header_row:
                c.font = Font(name=F_DOC, size=11, bold=True, color=WHITE)
                c.fill = PatternFill("solid", fgColor=BLUE)
                c.alignment = Alignment(vertical="center", wrap_text=True)
            else:
                c.font = Font(name=F_DOC, size=10, bold=(c.column == 1), color=BLUE_DARK if c.column == 1 else INK)
                if (r - header_row) % 2 == 0: c.fill = PatternFill("solid", fgColor=CARD)
                c.alignment = Alignment(vertical="top", wrap_text=True)
    for col in range(1, ws.max_column + 1):
        width = max((len(str(ws.cell(row=r, column=col).value or "")) for r in range(header_row, ws.max_row + 1)), default=8)
        ws.column_dimensions[get_column_letter(col)].width = min(max(10, width + 2), 50)
    ws.row_dimensions[header_row].height = 22
    if freeze: ws.freeze_panes = ws.cell(row=header_row + 1, column=1)
    ws.sheet_view.showGridLines = False

# =============================== PDF ===============================
def _find_soffice():
    for name in ("soffice", "libreoffice"):
        p = shutil.which(name)
        if p: return p
    for p in ("/Applications/LibreOffice.app/Contents/MacOS/soffice",
              r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
        if os.path.exists(p): return p
    return None

def to_pdf(path, outdir=None):
    """Make a PDF from a branded .docx or .pptx so the PDF carries the same theme/header/footer.
    Uses LibreOffice if installed; otherwise Microsoft Word/PowerPoint via docx2pdf (docx only)."""
    outdir = outdir or os.path.dirname(os.path.abspath(path))
    out = os.path.join(outdir, os.path.splitext(os.path.basename(path))[0] + ".pdf")
    soffice = _find_soffice()
    if soffice:
        subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", outdir, path],
                       check=True, capture_output=True, timeout=300)
        return out
    if path.lower().endswith(".docx"):
        try:
            from docx2pdf import convert            # needs Microsoft Word (Windows / macOS)
            convert(path, out); return out
        except ImportError:
            pass
    raise RuntimeError("No PDF converter found. Install LibreOffice (free, libreoffice.org), "
                       "or for Word files run: pip install docx2pdf (uses Microsoft Word). "
                       "The .docx/.pptx is already saved and can be exported to PDF from Word/PowerPoint.")

# =============================== ARTIFACTS / HTML ===============================
def _b64(name):
    return base64.b64encode(open(asset(name), "rb").read()).decode()

BRAND_CSS = """
@import url('https://fonts.googleapis.com/css2?family=Lexend:wght@600&family=Mulish:ital,wght@0,400;0,700;1,400&display=swap');
:root{--blue:#0094D5;--blue-dark:#00679A;--card:#EFF7FC;--muted:#5F6B76;--doc-blue:#2D74B5;--doc-blue2:#3D85C6}
*{box-sizing:border-box}
body{margin:0;font-family:Mulish,Arial,sans-serif;font-size:16px;line-height:1.55;color:#000;background:#fff;display:flex;flex-direction:column;min-height:100vh}
.sv-header{display:flex;align-items:center;justify-content:space-between;padding:18px 40px 14px 0;border-bottom:1px solid #e6eef4}
.sv-header .sv-title{display:flex;align-items:center;gap:28px;font-family:'Lexend',Mulish,sans-serif;font-weight:600;font-size:26px;color:var(--blue)}
.sv-header .sv-title:before{content:"";width:10px;height:34px;background:var(--blue)}
.sv-header img{height:34px}
.sv-sub{color:var(--muted);font-size:14px;margin:4px 0 0 38px}
main{flex:1;width:100%;max-width:1000px;margin:0 auto;padding:28px 40px 60px}
h1,h2,h3{font-family:'Lexend',Mulish,sans-serif;font-weight:600;color:var(--blue);line-height:1.25}
h1{font-size:28px}h2{font-size:22px;margin-top:32px}h3{font-size:17px;color:var(--blue-dark)}
a{color:#1155CC}
table{border-collapse:collapse;width:100%;margin:16px 0;font-size:14px}
th{background:var(--blue);color:#fff;text-align:left;font-weight:700;padding:8px 10px}
td{padding:8px 10px;border-bottom:1px solid #dfe9f1}
tr:nth-child(even) td{background:var(--card)}
td:first-child{font-weight:700;color:var(--blue-dark)}
.sv-cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:16px;margin:20px 0}
.sv-card{background:var(--card);padding:20px}
.sv-card .num{font-family:'Lexend',sans-serif;font-weight:600;font-size:30px;color:var(--blue)}
.sv-card .lbl{font-family:'Lexend',sans-serif;font-weight:600;font-size:13px;color:var(--blue)}
.sv-card .txt{font-size:13px;color:var(--muted);margin-top:8px}
.sv-footer{position:relative;background:var(--blue);color:#fff;display:flex;justify-content:space-between;align-items:center;padding:16px 40px;font-size:13px}
.sv-footer img{height:26px}
"""

def html_page(title, body_html, subtitle=""):
    """Branded self-contained HTML page for artifacts: Satva header (tab + title + logo) and footer (contact strip)."""
    sub = f'<div class="sv-sub">{subtitle}</div>' if subtitle else ""
    return f"""<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{title}</title><style>{BRAND_CSS}</style></head><body>
<header><div class="sv-header"><div class="sv-title">{title}</div><img alt="Satva Solutions" src="data:image/png;base64,{_b64('logo_blue')}"></div>{sub}</header>
<main>{body_html}</main>
<footer class="sv-footer"><img alt="Satva Solutions" src="data:image/png;base64,{_b64('logo_white')}"><span>{PHONE} &nbsp;|&nbsp; {EMAIL}</span></footer>
</body></html>"""

# =============================== SETUP CHECK ===============================
def check():
    """Print what is installed. Run: python satva_brand.py"""
    ok = True
    for mod, pipname in (("docx", "python-docx"), ("pptx", "python-pptx"), ("openpyxl", "openpyxl"), ("PIL", "Pillow")):
        try: __import__(mod); print(f"OK       {pipname}")
        except ImportError: ok = False; print(f"MISSING  {pipname}")
    print(("OK       PDF converter: " + _find_soffice()) if _find_soffice() else
          "OPTIONAL PDF converter: install LibreOffice (or pip install docx2pdf to use Microsoft Word)")
    print("Fonts: install Mulish and Lexend (free, fonts.google.com) so files and PDFs show the right fonts.")
    return ok

# ---- embedded brand images (base64 PNG) ----
ASSETS = {
    "logo_blue": "iVBORw0KGgoAAAANSUhEUgAAASAAAABJCAMAAACq0QuUAAAAYFBMVEUAlNUAlNUAlNUAlNUAlNUAlNUAlNUAlNUAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABuIXF2AAAAIHRSTlMA/q2QblHTLAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABjjZPgAAATESURBVHja7ZzpjuMgDICJD/z+b7yHulMC+AJGO6riX51pSsyHsY1BlNIKQUioOPJ6Tv8mLfL+WFKCX7/DuJajiFwJQY/P67nJV9ei8PtjzfCxfvf6f4AQQ05bB9AbNh8DVOj9mRKA3j0DFZ14jdSstg4gqyvLgMrtj6g0XesHi8Ltpfk4gMB6ch2Q1Dwh1jXBpnE4zMcGJF3HTgFqwWMQUFQRjkE+A+j+KJwDVNIm1MwiMczcbg8PA4LLfPUOoJKMZKSPUz9tIDTgtezLaJCcmwwlNpiYimDV1ZKWQs6C1HTSlACUjGQpT+i3gt/DJwg+2PEEIdHHiCZaVveVfHaCUc6jBh9OTDJUFcBpgs57mmXDKjfGBOcAxVtlfYBa3yTuOJ4E1LxZ2r/kHKBwJNNN7c4YPOIHAXVGbWQhG2oEWwXVMqTzjt6kPRjEevPHxCSLj1O9mWk6gjUrFuy0rN8cxca0K5EMJQw50qpuFqNvAtsNNd2iYxHMd5QbgAKFD73PNOIlm3hmqcGxgZV8zEm5wur5adH91Oy3diTLrsUQPQ9dp/rwOUDFW0PpJgaOluhVJ+ILR52zxLKRDUBOEgq6G5/PBrEjWb4gxPEEDrJpS3ZFM0ayJk55ESym5QkbUp0NB9dkOUAcq+pi2PKcSFZ3CRkKS2ySJfNV44WsWgOFCtRTZ8m4VTITa0BjkSyb0KuRzFh+WtGv+gk6VE9UQLbTDC3A0yserVVU+WBQy/V0kJVXoE0fI346DUjRJRTB0I6MeGI5mkncIFO3WygdNMNiRDAveaonSs9TQK6TiUSyPKDpkNfIPhjtFVKy6Ypkun+Vg2WpsVVRxyoQTDlTnon3I3SqwI9kK4CqUT6gBfOo13792QQU8qgn63aXWtopqQh2MJINr8EYdXfFsQSob9U96eIkY7wfyYZ+1HSJ+DpY+b0bhu5D8ucu6iFAEF2te1WuNUC3VvUKL0TnDl6prRhLofSsvewYsbh30EaeSATjsrxZtJIHcQa4WZVa3lyJbOUmEhy59iJZxxczrZFJcxWQ+NUqzOQ3Ez9NYYFL3fjKbyue2p6DRCkGciZJGyd2cCU3t55fn/p9pUZSUzuwclwGBBsHvvgcILZPaaTT4yGSHTCgYFYlhrIbwYOso2Kcz45RPdeW4iMrWWcg0myFDiuCXSvtrQHCbqryikfFc4AayyTjKNDqHnEWT83U4sO5/FZ+BoGNZlgzSUkCAuCt6pK64thLYMGNYNfmOYxHHnnkkUceeeSRRx555JFHHvlRwu0VEOjc7wBf6/nxnofiXBGB7dOzE8n9fRTzZszbMIzbJviuutZX6i7h4LtOWPwLN/DnASL/l397SuIA4uGWkkDnsVecfxqg4Zeg9t0BRDNAErhA5UaCvgfQfCoEAaGmb28c5AHCoWMfAYiigP69WQUklv/lzwbE8uc75OAUY5bZpPtkQC9vRaoKMvFl+CUwn2+fBOifRYiqAt54jLP04wH9Pg//uvMLA9d7lYjH/pwoxje/6wKi/waIzgOCBKBXQPMAyd+X9TmorCeKXobnp+M7gNzWb+0ihBLFERAHEnleWgtFGlkF1DeOfrscAoTj2sxdCmLx17OrjSwDuvcj1C77gP7g+AU19yYlUnUkKQAAAABJRU5ErkJggg==",
    "swoosh": "iVBORw0KGgoAAAANSUhEUgAAAWgAAAAuCAMAAAAyTXMiAAAAYFBMVEUAlNUAlNUAlNUAlNUAlNUAlNUAlNUAlNUAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABuIXF2AAAAIHRSTlP+CDxL05SrawAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALA1kbMAAADJSURBVHja7djLDcIADAXBkPDpv2M6IBdrZaRxCaO9PB+PiTt/3/twB+f/gebcQHNeAX0hHoK+CfqJeAZazw20nhvoU9AJ9I3zB+8MtJ4baD3vgLZUhqBvnF9sZ6D13EDruYHW8wroyyIcgra8G2hDpYHWcwNteTfQlncDrecGWs87oC3CIWjODTTnBppzA+2TtALaJ2kK2vJuoPXcQOu5gba8G2jLu4HWcwOt5x3QFIegLe8GmnMDzbmB5rwD2iIcgubcQHtBD90Xw28U76wApYoAAAAASUVORK5CYII=",
    "doc_logo": "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",
    "doc_watermark": "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",
}

if __name__ == "__main__":
    sys.exit(0 if check() else 1)
