"""
MedAI Suite — Professional PPTX Renderer
Phase 2: Server-side python-pptx with real gradients, Slide Masters, precise layout
"""

import json
import io
import base64
import traceback
from pptx import Presentation
from pptx.util import Inches, Pt, Emu, Cm
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn, nsmap
from pptx.oxml import parse_xml
from lxml import etree
import copy


# ── CONSTANTS ─────────────────────────────────────────────────────────────

W  = Inches(13.33)   # Slide width
H  = Inches(7.5)     # Slide height
MX = Inches(0.35)    # Margin X
MY = Inches(0.94)    # Content top (below header)
CW = Inches(12.63)   # Content width

# Colors
C = {
    'navy':     RGBColor(0x0D, 0x2B, 0x4E),
    'navyMid':  RGBColor(0x1A, 0x40, 0x72),
    'navyLt':   RGBColor(0x2E, 0x6D, 0xB4),
    'teal':     RGBColor(0x00, 0xB4, 0xD8),
    'tealDk':   RGBColor(0x00, 0x96, 0xC7),
    'tealLt':   RGBColor(0x90, 0xE0, 0xEF),
    'green':    RGBColor(0x2D, 0x6A, 0x4F),
    'greenLt':  RGBColor(0xD8, 0xF3, 0xDC),
    'greenAcc': RGBColor(0x40, 0x91, 0x6C),
    'amber':    RGBColor(0xE9, 0xC4, 0x6A),
    'amberDk':  RGBColor(0xF4, 0xA2, 0x61),
    'amberBg':  RGBColor(0xFF, 0xF8, 0xE7),
    'red':      RGBColor(0xC1, 0x12, 0x1F),
    'orange':   RGBColor(0xE7, 0x6F, 0x51),
    'white':    RGBColor(0xFF, 0xFF, 0xFF),
    'offWhite': RGBColor(0xF8, 0xFA, 0xFB),
    'slate':    RGBColor(0xE2, 0xE8, 0xF0),
    'stone':    RGBColor(0x94, 0xA3, 0xB8),
    'ink':      RGBColor(0x1E, 0x29, 0x3B),
    'ghost':    RGBColor(0xE8, 0xEF, 0xF8),
}

FONT_TITLE = 'Calibri Light'
FONT_BODY  = 'Calibri'


# ── XML HELPERS ───────────────────────────────────────────────────────────

def rgb_hex(color: RGBColor) -> str:
    return f'{color[0]:02X}{color[1]:02X}{color[2]:02X}'

def solid_fill_xml(color: RGBColor) -> str:
    h = rgb_hex(color)
    return f'<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:srgbClr val="{h}"/></a:solidFill>'

def grad_fill_xml(c1: RGBColor, c2: RGBColor, angle_deg: int = 90) -> str:
    """Linear gradient from c1 (start) to c2 (end). Angle in degrees × 60000."""
    h1, h2 = rgb_hex(c1), rgb_hex(c2)
    ang = angle_deg * 60000
    return f'''<a:gradFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" rotWithShape="1">
  <a:gsLst>
    <a:gs pos="0"><a:srgbClr val="{h1}"/></a:gs>
    <a:gs pos="100000"><a:srgbClr val="{h2}"/></a:gs>
  </a:gsLst>
  <a:lin ang="{ang}" scaled="0"/>
</a:gradFill>'''

def no_line_xml() -> str:
    return '<a:ln xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:noFill/></a:ln>'

def solid_line_xml(color: RGBColor, width_pt: float = 0.75) -> str:
    h = rgb_hex(color)
    w = int(width_pt * 12700)  # pt to EMU
    return f'<a:ln xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" w="{w}"><a:solidFill><a:srgbClr val="{h}"/></a:solidFill></a:ln>'

def set_fill(shape, fill_xml: str):
    """Replace shape fill with given XML and neutralize theme style override."""
    sp = shape._element
    spPr = sp.find(qn('p:spPr'))
    if spPr is None:
        return
    # Remove existing fills
    for tag in ['a:solidFill', 'a:gradFill', 'a:noFill', 'a:pattFill', 'a:blipFill']:
        el = spPr.find(qn(tag))
        if el is not None:
            spPr.remove(el)
    new_fill = parse_xml(fill_xml)
    spPr.insert(0, new_fill)
    # CRITICAL: Neutralize ALL p:style theme overrides for PowerPoint
    # fillRef idx=0 → no fill override
    # lnRef idx=0   → no border override  
    # effectRef idx=0 → no glow/shadow override
    # fontRef → use dk1 (dark) not lt1 (white) so text colors are not forced white
    style = sp.find(qn('p:style'))
    if style is not None:
        for ref_tag in ['a:fillRef', 'a:lnRef', 'a:effectRef']:
            ref = style.find(qn(ref_tag))
            if ref is not None:
                ref.set('idx', '0')
                for child in list(ref):
                    ref.remove(child)
        # fontRef: keep idx="minor" but change lt1 to dk1 so text isn't forced white
        fontRef = style.find(qn('a:fontRef'))
        if fontRef is not None:
            schemeClr = fontRef.find(qn('a:schemeClr'))
            if schemeClr is not None:
                schemeClr.set('val', 'dk1')

def set_line(shape, line_xml: str):
    """Replace shape line/border with given XML."""
    sp = shape._element
    spPr = sp.find(qn('p:spPr'))
    if spPr is None:
        return
    ln = spPr.find(qn('a:ln'))
    if ln is not None:
        spPr.remove(ln)
    new_ln = parse_xml(line_xml)
    spPr.append(new_ln)

def set_round_corners(shape, radius_pct: int = 10):
    """Add rounded corners to a shape (prstGeom or custGeom)."""
    sp = shape._element
    spPr = sp.find(qn('p:spPr'))
    if spPr is None:
        return
    # Replace prstGeom with rounded rect
    for geom in spPr.findall(qn('a:prstGeom')):
        spPr.remove(geom)
    prstGeom = parse_xml(
        f'<a:prstGeom xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" prst="roundRect">'
        f'<a:avLst><a:gd name="adj" fmla="val {radius_pct * 1000}"/></a:avLst>'
        f'</a:prstGeom>'
    )
    spPr.insert(0, prstGeom)


# ── SHAPE BUILDERS ────────────────────────────────────────────────────────

def add_rect(slide, x, y, w, h, fill_color=None, fill_xml=None,
             line_color=None, line_width=0.75, line_xml=None, radius=0):
    """Add a rectangle or rounded rectangle."""
    shape = slide.shapes.add_shape(1, x, y, w, h)  # MSO_SHAPE_TYPE.RECTANGLE = 1

    # Fill
    if fill_xml:
        set_fill(shape, fill_xml)
    elif fill_color:
        set_fill(shape, solid_fill_xml(fill_color))
    else:
        set_fill(shape, '<a:noFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>')

    # Line
    if line_xml:
        set_line(shape, line_xml)
    elif line_color:
        set_line(shape, solid_line_xml(line_color, line_width))
    else:
        set_line(shape, no_line_xml())

    # Rounded corners
    if radius > 0:
        set_round_corners(shape, radius)

    return shape


def add_text(slide, text, x, y, w, h,
             font_size=10, font_face=None, bold=False, italic=False,
             color=None, align='left', valign='top',
             fill_color=None, wrap=True, line_spacing=1.15):
    """Add a text box with full formatting control."""
    txBox = slide.shapes.add_textbox(x, y, w, h)
    tf = txBox.text_frame
    tf.word_wrap = wrap

    # Vertical alignment
    if valign == 'middle':
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    elif valign == 'bottom':
        tf.vertical_anchor = MSO_ANCHOR.BOTTOM
    else:
        tf.vertical_anchor = MSO_ANCHOR.TOP

    # Paragraph
    p = tf.paragraphs[0]

    # Alignment
    align_map = {'left': PP_ALIGN.LEFT, 'center': PP_ALIGN.CENTER,
                 'right': PP_ALIGN.RIGHT, 'justify': PP_ALIGN.JUSTIFY}
    p.alignment = align_map.get(align, PP_ALIGN.LEFT)

    # Line spacing
    from pptx.oxml.ns import qn as _qn
    pPr = p._pPr
    if pPr is None:
        pPr = p._p.get_or_add_pPr()
    lnSpc = etree.SubElement(pPr, _qn('a:lnSpc'))
    spcPct = etree.SubElement(lnSpc, _qn('a:spcPct'))
    spcPct.set('val', str(int(line_spacing * 100000)))

    # Run
    if text:
        run = p.add_run()
        run.text = str(text)
        font = run.font
        font.size = Pt(font_size)
        font.bold = bold
        font.italic = italic
        font.name = font_face or FONT_BODY
        if color:
            font.color.rgb = color

    # Background fill
    if fill_color:
        set_fill(txBox, solid_fill_xml(fill_color))
        set_line(txBox, no_line_xml())

    return txBox


def add_line(slide, x1, y1, x2, y2, color, width_pt=0.75, dash=None):
    """Add a line shape."""
    from pptx.util import Pt as _Pt
    connector = slide.shapes.add_connector(1, x1, y1, x2, y2)
    ln = connector.line
    ln.color.rgb = color
    ln.width = _Pt(width_pt)
    if dash:
        ln.dash_style = dash
    return connector


# ── HEADER & FOOTER ───────────────────────────────────────────────────────

def add_header(slide, heading, sub=None):
    """Premium gradient header bar."""
    # Background gradient
    bg = add_rect(slide, 0, 0, W, Inches(0.88),
                  fill_xml=grad_fill_xml(C['navy'], C['navyMid'], angle_deg=0))
    # Left teal accent
    add_rect(slide, 0, 0, Inches(0.22), Inches(0.88),
             fill_xml=grad_fill_xml(C['teal'], C['tealDk'], angle_deg=90))
    # Title
    add_text(slide, heading, Inches(0.32), 0, Inches(9.20), Inches(0.88),
             font_size=15, font_face=FONT_TITLE, bold=True,
             color=C['white'], valign='middle', line_spacing=1.0)
    # Subtitle
    if sub:
        add_text(slide, sub, Inches(9.55), 0, Inches(3.60), Inches(0.88),
                 font_size=8.5, color=C['tealLt'], align='right', valign='middle')


def add_home_button(slide):
    """Teal home button — top right."""
    btn = add_rect(slide, Inches(12.46), Inches(0.04), Inches(0.82), Inches(0.36),
                   fill_color=C['teal'], radius=6)
    add_text(slide, '⌂ Home', Inches(12.46), Inches(0.04), Inches(0.82), Inches(0.36),
             font_size=8, bold=True, color=C['white'], align='center', valign='middle')


def add_badge(slide, badge_type='green'):
    """Colored badge strip below header."""
    y = Inches(0.92)
    if badge_type == 'green':
        bg = C['greenLt']
        bd = C['greenAcc']
        fg = C['green']
        txt = '● with audited References — data sourced from publications, registries and regulatory documents'
    else:  # yellow
        bg = C['amberBg']
        bd = C['amberDk']
        fg = RGBColor(0x92, 0x40, 0x0E)
        txt = '⚠  DeepResearch Suggested Content — for internal strategic planning only'

    badge = add_rect(slide, MX, y, CW, Inches(0.24),
                     fill_color=bg, line_color=bd, line_width=0.75, radius=4)
    add_text(slide, txt, MX + Inches(0.10), y, CW - Inches(0.14), Inches(0.24),
             font_size=8, bold=True, color=fg, valign='middle')


def add_footer(slide, section_num=None, slide_idx=1, refs=None):
    """Footer with optional footnotes."""
    y = Inches(7.25)
    # Separator line
    add_line(slide, MX, y - Inches(0.04), MX + CW, y - Inches(0.04), C['slate'], 0.5)

    if refs:
        fn_text = '   '.join(f'[{i+1}] {_trunc(str(r), 100)}' for i, r in enumerate(refs[:4]))
        add_text(slide, fn_text, MX, y, Inches(11.60), Inches(0.20),
                 font_size=6.5, color=C['stone'])
    else:
        add_text(slide,
                 f'CORPORATE PRESENTATION · NOT FOR PROMOTION',
                 MX, y, Inches(11.60), Inches(0.20),
                 font_size=6.5, color=C['stone'])

    if section_num is not None:
        add_text(slide, f'{section_num}.{slide_idx}',
                 Inches(12.68), y, Inches(0.55), Inches(0.20),
                 font_size=6.5, color=C['stone'], align='right')


# ── UTILITIES ─────────────────────────────────────────────────────────────

def _trunc(s, n):
    if not s:
        return ''
    s = str(s)
    return s[:n-1] + '…' if len(s) > n else s

def _strip_refs(text):
    """Remove inline citations from text."""
    import re
    if not text:
        return ''
    text = str(text)
    # Remove long parenthetical refs: (Author et al 2020, Journal...)
    text = re.sub(r'\s*\([A-Z][^)]{15,}\)', '', text)
    # Remove bracket refs: [Author 2020]
    text = re.sub(r'\s*\[[A-Z][^\]]{5,}\]', '', text)
    return text.strip()

def _parse_orr(s):
    """Parse ORR string, return float 0-100."""
    import re
    if not s:
        return 0.0
    m = re.match(r'^[^0-9]*([0-9]+\.?[0-9]*)', str(s))
    if m:
        return min(100.0, max(0.0, float(m.group(1))))
    return 0.0


# ── TABLE HELPER ──────────────────────────────────────────────────────────

def add_table(slide, headers, rows, x, y, w, col_widths, row_h,
              header_fill=None, alt_fill=None):
    """
    Add a formatted table.
    headers: list of str
    rows: list of list of {text, bold, color, fill, align}
    col_widths: list of floats in inches
    """
    from pptx.util import Inches as I
    n_cols = len(headers)
    n_rows = len(rows) + 1  # +1 for header
    table = slide.shapes.add_table(n_rows, n_cols, x, y, w, I(row_h * n_rows)).table

    # Column widths
    for i, cw in enumerate(col_widths):
        table.columns[i].width = I(cw)

    # Header row — supports both plain strings and hcell() dicts
    for j, h in enumerate(headers):
        cell = table.cell(0, j)
        if isinstance(h, dict):
            # hcell() dict format
            h_text  = _trunc(str(h.get('text', '')), 120)
            h_bold  = h.get('bold', True)
            h_color = h.get('color', C['white'])
            h_fill  = h.get('fill', header_fill or C['navy'])
            h_align = h.get('align', 'center')
            h_sz    = h.get('sz', 9.5)
        else:
            h_text  = _trunc(str(h), 120)
            h_bold  = True
            h_color = C['white']
            h_fill  = header_fill or C['navy']
            h_align = 'center'
            h_sz    = 9.5
        cell.text = h_text
        p = cell.text_frame.paragraphs[0]
        align_map = {'left': PP_ALIGN.LEFT, 'center': PP_ALIGN.CENTER, 'right': PP_ALIGN.RIGHT}
        p.alignment = align_map.get(h_align, PP_ALIGN.CENTER)
        run = p.runs[0] if p.runs else p.add_run()
        run.font.size = Pt(h_sz)
        run.font.name = FONT_TITLE
        run.font.bold = h_bold
        run.font.color.rgb = h_color
        # Cell fill
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        fill_el = tcPr.find(qn('a:solidFill'))
        if fill_el is not None:
            tcPr.remove(fill_el)
        tcPr.insert(0, parse_xml(solid_fill_xml(h_fill)))

    # Data rows
    alt = alt_fill or C['ghost']
    for i, row in enumerate(rows):
        bg = alt if i % 2 == 1 else C['white']
        for j, cell_data in enumerate(row):
            cell = table.cell(i + 1, j)
            if isinstance(cell_data, dict):
                txt   = _trunc(str(cell_data.get('text', '—')), 400)
                bold  = cell_data.get('bold', False)
                color = cell_data.get('color', C['ink'])
                fill  = cell_data.get('fill', bg)
                align = cell_data.get('align', 'left')
                sz    = cell_data.get('sz', 9)
            else:
                txt, bold, color, fill, align, sz = str(cell_data), False, C['ink'], bg, 'left', 9

            cell.text = txt
            p = cell.text_frame.paragraphs[0]
            align_map = {'left': PP_ALIGN.LEFT, 'center': PP_ALIGN.CENTER, 'right': PP_ALIGN.RIGHT}
            p.alignment = align_map.get(align, PP_ALIGN.LEFT)
            run = p.runs[0] if p.runs else p.add_run()
            run.font.size = Pt(sz)
            run.font.name = FONT_BODY
            run.font.bold = bold
            run.font.color.rgb = color
            # Cell fill
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            f_el = tcPr.find(qn('a:solidFill'))
            if f_el is not None:
                tcPr.remove(f_el)
            # fill=None means use alternating row color (already set as bg above)
            effective_fill = fill if fill is not None else bg
            tcPr.insert(0, parse_xml(solid_fill_xml(effective_fill)))

    return table


def hcell(txt, align='center', bold=True, color=None, fill=None, sz=9.5):
    return {'text': txt, 'bold': bold, 'color': color or C['white'],
            'fill': fill or C['navy'], 'align': align, 'sz': sz}

def dcell(txt, align='left', bold=False, color=None, fill=None, sz=9.0):
    return {'text': txt, 'bold': bold, 'color': color or C['ink'],
            'fill': fill, 'align': align, 'sz': sz}  # fill=None → alternating applied by table


# ── SLIDE RENDERERS ───────────────────────────────────────────────────────

def render_title(prs, meta):
    sl = prs.slides.add_slide(prs.slide_layouts[6])

    # Full gradient background
    add_rect(sl, 0, 0, W, H,
             fill_xml=grad_fill_xml(C['ink'], C['navy'], angle_deg=135))
    # Left teal bar
    add_rect(sl, 0, 0, Inches(0.28), H,
             fill_xml=grad_fill_xml(C['teal'], C['tealDk'], angle_deg=90))
    # Bottom bar
    add_rect(sl, 0, Inches(6.72), W, Inches(0.78), fill_color=RGBColor(0x09, 0x1E, 0x3A))

    # MAP type badge
    scope = meta.get('scope', 'GLOBAL MAP').upper()
    badge_box = add_rect(sl, Inches(0.42), Inches(0.60), Inches(3.2), Inches(0.38),
                         fill_color=C['teal'], radius=4)
    add_text(sl, scope, Inches(0.42), Inches(0.60), Inches(3.2), Inches(0.38),
             font_size=9, font_face=FONT_TITLE, bold=True,
             color=C['white'], align='center', valign='middle')

    # Eyebrow
    add_text(sl, 'MEDICAL AFFAIRS PLAN',
             Inches(0.42), Inches(1.10), Inches(12.0), Inches(0.42),
             font_size=11, font_face=FONT_TITLE, color=C['tealLt'])

    # Title
    add_text(sl, meta.get('title', 'Medical Affairs Plan'),
             Inches(0.42), Inches(1.52), Inches(11.80), Inches(2.0),
             font_size=34, font_face=FONT_TITLE, bold=True,
             color=C['white'], valign='bottom', line_spacing=1.1)

    # Accent line
    add_rect(sl, Inches(0.42), Inches(3.62), Inches(6.0), Inches(0.05),
             fill_xml=grad_fill_xml(C['teal'], C['navyMid'], angle_deg=0))

    # Drug + indication
    add_text(sl, f"{meta.get('drug', '')}  ·  {meta.get('indication', '')}",
             Inches(0.42), Inches(3.72), Inches(10.0), Inches(0.58),
             font_size=17, font_face=FONT_TITLE,
             color=RGBColor(0xB8, 0xD4, 0xF5), valign='middle')

    # Year + scope label
    scope_label = meta.get('scope_label', '')
    company = meta.get('company', '')
    sub_line = f"{meta.get('year', 2027)}  ·  {scope_label} Medical Affairs Plan"
    if company:
        sub_line += f"  ·  {company}"
    add_text(sl, sub_line,
             Inches(0.42), Inches(4.38), Inches(10.0), Inches(0.34),
             font_size=10.5, font_face=FONT_BODY, bold=True, color=C['teal'])

    # Disclaimer
    add_text(sl, 'CORPORATE PRESENTATION — NOT FOR PROMOTION',
             Inches(0.42), Inches(6.86), Inches(12.0), Inches(0.30),
             font_size=7.5, font_face=FONT_BODY,
             color=RGBColor(0x60, 0x70, 0x90))

    return sl


def render_toc(prs, sections, meta):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    sl.background.fill.solid()
    sl.background.fill.fore_color.rgb = C['offWhite']

    add_header(sl, 'Contents',
               f"{meta.get('drug')} · {meta.get('indication')} · {meta.get('year')}")
    add_home_button(sl)

    half = (len(sections) + 1) // 2
    left = sections[:half]
    right = sections[half:]

    # Each section = divider + content = 2 slides minimum
    # Slide 1=Title, 2=ToC, then pairs
    def draw_col(items, start_idx, x_base):
        for i, sec in enumerate(items):
            y = Inches(1.04 + i * 0.50)
            slide_target = 2 + (start_idx + i) * 2 + 1  # divider slide

            # Chevron shape
            chev = add_rect(sl,
                Inches(x_base), y + Inches(0.05),
                Inches(0.34), Inches(0.28),
                fill_xml=grad_fill_xml(C['teal'], C['navyLt'], angle_deg=0))

            # Label
            add_text(sl, f"{sec.get('icon','')}  {sec.get('label','')}",
                     Inches(x_base + 0.44), y, Inches(5.46), Inches(0.38),
                     font_size=10.5, color=C['ink'], valign='middle')

            # Dotted separator
            add_line(sl,
                     Inches(x_base + 0.44), y + Inches(0.42),
                     Inches(x_base + 5.90), y + Inches(0.42),
                     C['slate'], 0.5)

    draw_col(left, 0, 0.35)
    draw_col(right, len(left), 6.98)
    add_footer(sl)
    return sl


def render_divider(prs, n, label, icon=''):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(sl, 0, 0, W, H,
             fill_xml=grad_fill_xml(C['navy'], C['ink'], angle_deg=135))
    add_rect(sl, 0, 0, Inches(0.22), H,
             fill_xml=grad_fill_xml(C['teal'], C['tealDk'], angle_deg=90))
    add_rect(sl, Inches(11.20), 0, Inches(2.13), H, fill_color=RGBColor(0x09, 0x1E, 0x3A))

    # Giant number
    add_text(sl, str(n), Inches(0.35), Inches(0.15), Inches(3.0), Inches(3.5),
             font_size=120, font_face=FONT_TITLE, bold=True,
             color=RGBColor(0x1A, 0x34, 0x60), valign='top')

    # Right panel — large section number echo
    add_text(sl, str(n), Inches(11.30), Inches(1.60), Inches(1.80), Inches(2.20),
             font_size=88, font_face=FONT_TITLE, bold=True,
             color=C['teal'], align='center', valign='middle')

    # Section label
    add_text(sl, label, Inches(0.42), Inches(3.0), Inches(10.50), Inches(1.60),
             font_size=34, font_face=FONT_TITLE, bold=True,
             color=C['white'], valign='middle', line_spacing=1.1)

    # Teal underline
    add_rect(sl, Inches(0.42), Inches(4.72), Inches(5.0), Inches(0.06),
             fill_xml=grad_fill_xml(C['teal'], C['navyMid'], angle_deg=0))

    add_text(sl, 'CORPORATE PRESENTATION — NOT FOR PROMOTION',
             Inches(0.42), Inches(6.84), Inches(12.0), Inches(0.30),
             font_size=7.5, color=RGBColor(0x60, 0x70, 0x90))
    add_home_button(sl)
    return sl


def render_executive_summary(prs, data, meta, sec_n):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    sl.background.fill.solid()
    sl.background.fill.fore_color.rgb = C['offWhite']
    add_header(sl, f"MAP {meta.get('year')}: Executive Summary",
               f"{meta.get('drug')} · {meta.get('indication')}")
    add_home_button(sl)
    add_badge(sl, 'yellow')

    rows_data = data.get('rows', [])[:6]
    avail_h = 5.88
    row_h = min(1.02, avail_h / max(len(rows_data), 1))
    icons_list = ['01', '02', '03', '04', '05', '06']

    for i, row in enumerate(rows_data):
        y = Inches(1.22 + i * (row_h + 0.03))
        bg = C['white'] if i % 2 == 0 else C['ghost']
        # Card
        add_rect(sl, MX, y, CW, Inches(row_h), fill_color=bg,
                 line_color=C['slate'], line_width=0.5, radius=6)
        # Left accent
        add_rect(sl, MX, y, Inches(0.08), Inches(row_h), fill_color=C['teal'], radius=3)
        # Icon — colored circle with number (reliable cross-platform)
        icon_lbl = icons_list[i] if i < len(icons_list) else str(i+1).zfill(2)
        ic = add_rect(sl, MX + Inches(0.14), y + Inches((row_h-0.36)/2),
                      Inches(0.36), Inches(0.36),
                      fill_xml=grad_fill_xml(C['teal'], C['tealDk'], angle_deg=90), radius=50)
        add_text(sl, icon_lbl, MX + Inches(0.14), y + Inches((row_h-0.36)/2),
                 Inches(0.36), Inches(0.36),
                 font_size=8, bold=True, color=C['white'], align='center', valign='middle')
        # Topic
        add_text(sl, row.get('topic', ''),
                 MX + Inches(0.66), y + Inches(0.05),
                 Inches(2.28), Inches(row_h - 0.10),
                 font_size=9.5, font_face=FONT_TITLE, bold=True,
                 color=C['navy'], valign='middle')
        # Separator
        add_line(sl, MX + Inches(3.04), y + Inches(0.08),
                 MX + Inches(3.04), y + Inches(row_h - 0.08),
                 C['slate'], 0.5)
        # Summary text (no refs)
        summary = _strip_refs(row.get('summary', ''))
        add_text(sl, _trunc(summary, 400),
                 MX + Inches(3.16), y + Inches(0.05),
                 CW - Inches(3.24), Inches(row_h - 0.10),
                 font_size=9.5, color=C['ink'], valign='middle', line_spacing=1.15)

    add_footer(sl, sec_n, 1)
    return sl


def render_prevalence(prs, data, meta, sec_n):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    sl.background.fill.solid()
    sl.background.fill.fore_color.rgb = C['offWhite']
    add_header(sl, f"Prevalence & Epidemiology: {meta.get('indication')} — {meta.get('scope_label')}",
               f"{meta.get('drug')} · {meta.get('year')}")
    add_home_button(sl)
    add_badge(sl, 'green')

    refs = []
    kpis = data.get('kpis', [])[:4]
    scope_label = meta.get('scope_label', '')
    flag = meta.get('flag', '🌍')

    # LEFT: country panel
    add_rect(sl, MX, Inches(1.22), Inches(3.42), Inches(5.0),
             fill_xml=grad_fill_xml(C['navy'], C['navyMid'], angle_deg=135), radius=10)
    add_rect(sl, MX, Inches(1.22), Inches(3.42), Inches(0.08), fill_color=C['teal'])

    add_text(sl, flag, MX, Inches(1.30), Inches(3.42), Inches(1.62),
             font_size=54, align='center', valign='middle')
    add_text(sl, scope_label, MX, Inches(2.96), Inches(3.42), Inches(0.48),
             font_size=14, font_face=FONT_TITLE, bold=True,
             color=C['white'], align='center', valign='middle')
    add_text(sl, meta.get('indication', ''), MX, Inches(3.48), Inches(3.42), Inches(0.36),
             font_size=8.5, color=C['tealLt'], align='center', valign='middle', italic=True)
    add_rect(sl, MX + Inches(0.40), Inches(3.88), Inches(2.62), Inches(0.05),
             fill_color=C['teal'])
    add_text(sl, f"{meta.get('year')} Epidemiology Data",
             MX, Inches(3.96), Inches(3.42), Inches(0.30),
             font_size=8.5, font_face=FONT_BODY, bold=True, color=C['teal'], align='center')

    # RIGHT: 2×2 KPI grid
    kpi_colors = [C['teal'], C['navyMid'], RGBColor(0x2E, 0x86, 0xAB), C['orange']]
    kpi_icons  = ['EP', 'PT', 'RX', 'OS']

    for i, kpi in enumerate(kpis):
        col = i % 2
        row = i // 2
        kw, kh = 4.34, 2.38
        x = Inches(3.96 + col * (kw + 0.12))
        y = Inches(1.22 + row * (kh + 0.08))
        kc = kpi_colors[i] if i < len(kpi_colors) else C['navyMid']

        # Card
        add_rect(sl, x, y, Inches(kw), Inches(kh),
                 fill_color=C['white'], line_color=C['slate'], line_width=0.75, radius=10)
        # Header bar
        add_rect(sl, x, y, Inches(kw), Inches(0.34),
                 fill_xml=grad_fill_xml(kc, C['navy'], angle_deg=0), radius=6)
        # Icon + label
        kpi_code = kpi_icons[i] if i < len(kpi_icons) else '●'
        label_txt = f"[ {kpi_code} ]  {_trunc(kpi.get('label',''), 42)}"
        add_text(sl, label_txt, x + Inches(0.10), y, Inches(kw - 0.14), Inches(0.34),
                 font_size=8, bold=True, color=C['white'], valign='middle')
        # Big value
        add_text(sl, kpi.get('value', '—'),
                 x + Inches(0.10), y + Inches(0.38), Inches(kw - 0.20), Inches(1.22),
                 font_size=24, font_face=FONT_TITLE, bold=True,
                 color=kc, align='center', valign='middle')
        # Source ref
        if kpi.get('source'):
            refs.append(kpi['source'])
            add_rect(sl, x, y + Inches(kh - 0.30), Inches(kw), Inches(0.30),
                     fill_color=C['ghost'])
            add_text(sl, _trunc(_strip_refs(kpi['source']), 80),
                     x + Inches(0.10), y + Inches(kh - 0.28), Inches(kw - 0.36), Inches(0.26),
                     font_size=6.5, color=C['stone'], italic=True, valign='middle')
            add_text(sl, f'[{len(refs)}]',
                     x + Inches(kw - 0.30), y + Inches(kh - 0.28), Inches(0.26), Inches(0.26),
                     font_size=7, bold=True, color=kc, align='right', valign='middle')

    # Bottom epidemiology table
    ctx_rows = data.get('context_table', [])[:4]
    if ctx_rows:
        tbl_rows = []
        for row in ctx_rows:
            if row.get('source'):
                refs.append(row['source'])
            tbl_rows.append([
                dcell(row.get('metric', ''), bold=True),
                dcell(row.get('value', '—'), align='center'),
                dcell(f'[{len(refs)}]' if row.get('source') else '', align='center', color=C['stone'], sz=8.5),
            ])
        add_table(sl,
                  [hcell('Parameter', align='left'), hcell(f'{scope_label} Data'), hcell('Ref')],
                  tbl_rows,
                  MX, Inches(6.22), CW, [5.0, 6.78, 0.85], 0.34)

    if data.get('second_line_detail'):
        add_text(sl, f"Patient flow: {_trunc(_strip_refs(data['second_line_detail']), 220)}",
                 MX, Inches(7.02), CW, Inches(0.20),
                 font_size=8.5, bold=True, color=C['teal'])

    add_footer(sl, sec_n, 1, refs)
    return sl


def render_unmet_needs(prs, data, meta, sec_n):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    sl.background.fill.solid()
    sl.background.fill.fore_color.rgb = C['offWhite']
    add_header(sl, data.get('heading') or f"Unmet Medical Needs: {meta.get('indication')} — {meta.get('scope_label')}",
               f"{meta.get('drug')} · {meta.get('year')}")
    add_home_button(sl)
    add_badge(sl, 'green')

    if data.get('intro'):
        add_text(sl, _strip_refs(data['intro']),
                 MX, Inches(1.20), CW, Inches(0.28),
                 font_size=9.5, color=C['stone'], italic=True, valign='middle')

    needs = data.get('needs', [])[:5]
    start_y = 1.52 if data.get('intro') else 1.22
    avail_h = 7.10 - start_y
    nh = min(1.14, avail_h / max(len(needs), 1))
    refs = []
    mag_colors = {'HIGH': C['red'], 'MEDIUM': C['amberDk'], 'LOW': C['navyLt']}

    for i, need in enumerate(needs):
        y = Inches(start_y + i * (nh + 0.04))
        mc = mag_colors.get(need.get('magnitude', ''), C['navyLt'])
        bg = C['white'] if i % 2 == 0 else C['ghost']

        add_rect(sl, MX, y, CW, Inches(nh), fill_color=bg,
                 line_color=mc, line_width=0.75, radius=8)
        add_rect(sl, MX, y, Inches(0.14), Inches(nh), fill_color=mc, radius=3)

        # Magnitude pill
        add_rect(sl, MX + Inches(0.22), y + Inches(0.06), Inches(0.80), Inches(0.22),
                 fill_color=mc, radius=6)
        add_text(sl, need.get('magnitude', ''),
                 MX + Inches(0.22), y + Inches(0.06), Inches(0.80), Inches(0.22),
                 font_size=7.5, bold=True, color=C['white'], align='center', valign='middle')

        # Title
        add_text(sl, _strip_refs(need.get('title', '')),
                 MX + Inches(0.22), y + Inches(0.32), Inches(2.88), Inches(nh - 0.40),
                 font_size=9.5, font_face=FONT_TITLE, bold=True,
                 color=mc, valign='top')

        # Divider
        add_line(sl, MX + Inches(3.56), y + Inches(0.07),
                 MX + Inches(3.56), y + Inches(nh - 0.07), C['slate'], 0.5)

        # Detail
        add_text(sl, _strip_refs(need.get('detail', '')),
                 MX + Inches(3.68), y + Inches(0.06), CW - Inches(3.76), Inches(nh - 0.12),
                 font_size=9.5, color=C['ink'], valign='middle', line_spacing=1.15)

        if need.get('source'):
            refs.append(need['source'])
            add_text(sl, f'[{len(refs)}]',
                     MX + CW - Inches(0.30), y + Inches(nh - 0.24), Inches(0.26), Inches(0.20),
                     font_size=6.5, color=C['stone'], italic=True, align='right')

    add_footer(sl, sec_n, 1, refs)
    return sl


def render_guidelines(prs, data, meta, sec_n):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    sl.background.fill.solid()
    sl.background.fill.fore_color.rgb = C['offWhite']
    add_header(sl, data.get('heading') or f"{meta.get('indication')} Guidelines — {meta.get('drug')}",
               f"{meta.get('scope_label')} · {meta.get('year')}")
    add_home_button(sl)
    add_badge(sl, 'green')

    if data.get('scope_note'):
        add_text(sl, f"Scope: {_trunc(data['scope_note'], 250)}",
                 MX, Inches(1.20), CW, Inches(0.22),
                 font_size=9, color=C['stone'], italic=True, valign='middle')

    rows_data = data.get('rows', [])[:8]
    refs = []
    start_y = 1.46 if data.get('scope_note') else 1.20
    tbl_rows = []
    for row in rows_data:
        if row.get('source'):
            refs.append(row['source'])
        import re as _re
        grade = ''
        rec = row.get('recommendation', '')
        m = _re.search(r'Level [IVX]+|Cat\. [12][AB]|Grade [A-C]|[IV]+[,\s][AB]', rec)
        if m:
            grade = m.group(0)
        tbl_rows.append([
            dcell(_trunc(row.get('guideline', ''), 50), bold=True),
            dcell(_trunc(row.get('line', ''), 20), align='center'),
            dcell(_trunc(_strip_refs(row.get('refractory_status', '')), 100), color=C['stone']),
            dcell(_trunc(_strip_refs(rec), 180), bold=True, color=C['teal']),
            dcell(_trunc(grade, 30), align='center'),
            dcell(f'[{len(refs)}]' if row.get('source') else '', align='center', color=C['stone'], sz=8.5),
        ])

    if tbl_rows:
        add_table(sl,
                  [hcell('Guideline (Year)', align='left'), hcell('Line'),
                   hcell('Patient Type', align='left'), hcell('Recommendation', align='left'),
                   hcell('Grade'), hcell('Ref')],
                  tbl_rows, MX, Inches(start_y), CW,
                  [2.0, 0.80, 2.48, 4.96, 1.06, 1.33], 0.54)

    if data.get('key_message'):
        add_rect(sl, MX, Inches(6.82), CW, Inches(0.38),
                 fill_color=C['ghost'], line_color=C['teal'], line_width=1.0, radius=6)
        add_text(sl, f"Key message: {_trunc(_strip_refs(data['key_message']), 300)}",
                 MX + Inches(0.12), Inches(6.82), CW - Inches(0.18), Inches(0.38),
                 font_size=9.5, font_face=FONT_TITLE, bold=True, color=C['teal'], valign='middle')

    add_footer(sl, sec_n, 1, refs)
    return sl


def render_pivotal(prs, data, meta, sec_n):
    """One slide per trial."""
    trials = [t for t in data.get('trials', [])
              if t and t.get('name') and '[DATA PENDING' not in t.get('name', '')]
    if not trials:
        sl = prs.slides.add_slide(prs.slide_layouts[6])
        sl.background.fill.solid()
        sl.background.fill.fore_color.rgb = C['offWhite']
        add_header(sl, f"Pivotal Clinical Evidence: {meta.get('drug')}", f"Key Trial Data · {meta.get('year')}")
        add_home_button(sl)
        add_badge(sl, 'green')
        add_text(sl, 'No pivotal trial data — verify at EMA EPAR or PubMed',
                 MX, Inches(2.5), CW, Inches(0.5), font_size=11, color=C['stone'], align='center')
        add_footer(sl, sec_n, 1)
        return [sl]

    slides = []
    for ti, trial in enumerate(trials):
        sl = _render_one_trial(prs, trial, meta, sec_n, ti + 1, len(trials))
        slides.append(sl)
    return slides


def _render_one_trial(prs, trial, meta, sec_n, idx, total):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    sl.background.fill.solid()
    sl.background.fill.fore_color.rgb = C['offWhite']

    subtitle = f'Trial {idx}/{total} · {meta.get("year")}' if total > 1 else f'Pivotal Evidence · {meta.get("year")}'
    add_header(sl, f"Pivotal Evidence: {_trunc(trial.get('name',''), 55)}", subtitle)
    add_home_button(sl)
    add_badge(sl, 'green')

    eff  = trial.get('efficacy', {})
    refs = []
    if trial.get('source'):
        refs.append(trial['source'])

    # Trial banner
    add_rect(sl, MX, Inches(1.20), CW, Inches(0.72),
             fill_xml=grad_fill_xml(C['navy'], C['navyMid'], angle_deg=0), radius=8)
    add_text(sl, f"{_trunc(trial.get('name',''),70)}  |  {trial.get('phase','Phase 3')}  |  N = {trial.get('n_total','?')}  |  NCT: {trial.get('nct','?')}",
             MX + Inches(0.14), Inches(1.20), CW - Inches(0.20), Inches(0.42),
             font_size=11, font_face=FONT_TITLE, bold=True,
             color=C['white'], align='center', valign='middle')
    add_text(sl, _trunc(f"{trial.get('design','')} · {trial.get('population','')}", 280),
             MX + Inches(0.14), Inches(1.62), CW - Inches(0.20), Inches(0.28),
             font_size=8.5, color=C['tealLt'], align='center', valign='middle')

    # 3 KPI boxes left
    kpis = [
        (eff.get('mpfs_drug', '?'), f"mPFS — {meta.get('drug','')}", eff.get('mpfs_control'), C['teal']),
        (eff.get('hr_pfs', '?'),    'HR (PFS)',                        None,                    C['navyMid']),
        (eff.get('orr_drug', '?'),  f"ORR — {meta.get('drug','')}",   eff.get('orr_control'),  C['navyLt']),
    ]
    for i, (val, label, sub, kc) in enumerate(kpis):
        y = Inches(2.00 + i * 1.02)
        add_rect(sl, MX, y, Inches(3.44), Inches(0.90),
                 fill_xml=grad_fill_xml(kc, C['navy'], angle_deg=135), radius=8)
        add_text(sl, _trunc(str(val), 28), MX + Inches(0.08), y + Inches(0.02),
                 Inches(3.28), Inches(0.54), font_size=22, font_face=FONT_TITLE, bold=True,
                 color=C['white'], align='center', valign='middle')
        add_text(sl, label, MX + Inches(0.08), y + Inches(0.58), Inches(3.28), Inches(0.22),
                 font_size=8, color=C['tealLt'], align='center')
        if sub:
            add_text(sl, _trunc(str(sub), 45), MX + Inches(0.08), y + Inches(0.76),
                     Inches(3.28), Inches(0.14), font_size=7, color=RGBColor(0xAA,0xCC,0xDD),
                     align='center', italic=True)

    # ORR bar chart
    orrD = _parse_orr(eff.get('orr_drug', '0'))
    orrC = _parse_orr(eff.get('orr_control', '0'))
    cX, cY, cW, cH = Inches(3.90), Inches(2.00), Inches(4.10), Inches(2.88)

    add_rect(sl, cX, cY, cW, cH + Inches(0.48),
             fill_color=C['white'], line_color=C['slate'], line_width=0.75, radius=8)
    add_text(sl, 'Overall Response Rate (ORR)', cX, cY + Inches(0.06), cW, Inches(0.26),
             font_size=9, font_face=FONT_TITLE, bold=True, color=C['navy'], align='center')

    innerH = Emu(cH.emu - Inches(0.12).emu)
    bW = Inches(1.20)
    bGap = Inches(0.50)

    if orrD > 0:
        dh  = Emu(int((orrD / 100) * (cH.emu - Inches(0.12).emu)))
        dby = Emu(cY.emu + Inches(0.36).emu + innerH.emu - dh.emu)
        add_rect(sl, cX + Inches(0.54), dby, bW, dh, fill_color=C['teal'], radius=6)
        add_text(sl, f'{orrD:.0f}%',
                 cX + Inches(0.54), Emu(max(cY.emu + Inches(0.10).emu, dby.emu - Inches(0.26).emu)), bW, Inches(0.24),
                 font_size=12, font_face=FONT_TITLE, bold=True, color=C['teal'], align='center')
        add_text(sl, _trunc(meta.get('drug', ''), 14),
                 cX + Inches(0.54), cY + Inches(0.36) + innerH + Inches(0.04), bW, Inches(0.22),
                 font_size=8, bold=True, color=C['navy'], align='center')

    if orrC > 0:
        ch  = Emu(int((orrC / 100) * (cH.emu - Inches(0.12).emu)))
        cby = Emu(cY.emu + Inches(0.36).emu + innerH.emu - ch.emu)
        cx2 = cX + Inches(0.54) + bW + bGap
        add_rect(sl, cx2, cby, bW, ch, fill_color=C['navyMid'], radius=6)
        add_text(sl, f'{orrC:.0f}%', cx2, Emu(max(cY.emu + Inches(0.10).emu, cby.emu - Inches(0.26).emu)),
                 bW, Inches(0.24), font_size=12, font_face=FONT_TITLE, bold=True,
                 color=C['navyMid'], align='center')
        add_text(sl, 'Control', cx2, cY + Inches(0.36) + innerH + Inches(0.04),
                 bW, Inches(0.22), font_size=8, color=C['stone'], align='center')

    # Baseline
    add_rect(sl, cX + Inches(0.18), Emu(cY.emu + Inches(0.36).emu + innerH.emu), cW - Inches(0.36), Inches(0.02),
             fill_color=C['navy'])

    # OS data
    rX, rW = Inches(8.18), Inches(4.78)
    add_rect(sl, rX, Inches(2.00), rW, Inches(0.32),
             fill_color=C['navy'], radius=6)
    add_text(sl, 'Overall Survival', rX, Inches(2.00), rW, Inches(0.32),
             font_size=9.5, font_face=FONT_TITLE, bold=True,
             color=C['white'], align='center', valign='middle')

    os_items = [
        (f"mOS — {meta.get('drug','')}", eff.get('mos_drug', '—')),
        ('mOS — Control',                eff.get('mos_control', '—')),
        ('HR (OS)',                        eff.get('hr_os', '—')),
    ]
    for i, (lbl, val) in enumerate(os_items):
        yy = Inches(2.36 + i * 0.42)
        bg = C['white'] if i % 2 == 0 else C['ghost']
        add_rect(sl, rX, yy, rW, Inches(0.38), fill_color=bg,
                 line_color=C['slate'], line_width=0.5, radius=4)
        add_text(sl, lbl, rX + Inches(0.12), yy, Inches(rW/Inches(1) * 0.58), Inches(0.38),
                 font_size=9, color=C['stone'], valign='middle')
        add_text(sl, _trunc(str(val), 25),
                 rX + Inches(2.80), yy, Inches(1.88), Inches(0.38),
                 font_size=10, font_face=FONT_TITLE, bold=True,
                 color=C['navy'], align='right', valign='middle')

    # Safety section
    safeY = Inches(4.90)
    add_rect(sl, MX, safeY, CW, Inches(0.28),
             fill_color=C['navy'], radius=6)
    add_text(sl, 'Key Safety Data — Grade 3/4 Adverse Events',
             MX + Inches(0.14), safeY, CW - Inches(0.20), Inches(0.28),
             font_size=9.5, font_face=FONT_TITLE, bold=True,
             color=C['white'], align='center', valign='middle')

    safety = trial.get('safety', {})
    heme  = [_trunc(_strip_refs(a), 70) for a in safety.get('grade34_heme', [])[:4]]
    non_h = [_trunc(_strip_refs(a), 70) for a in safety.get('grade34_nonheme', [])[:4]]

    for panel_aes, panel_lbl, px in [(heme, 'Haematological', MX), (non_h, 'Non-Haematological', Inches(6.60))]:
        pw = Inches(6.10) if px == MX else Inches(6.38)
        add_rect(sl, px, safeY + Inches(0.30), pw, Inches(0.24),
                 fill_color=C['navyMid'], radius=6)
        add_text(sl, panel_lbl, px, safeY + Inches(0.30), pw, Inches(0.24),
                 font_size=9, font_face=FONT_TITLE, bold=True,
                 color=C['white'], align='center', valign='middle')
        for i, ae in enumerate(panel_aes):
            bg = C['white'] if i % 2 == 0 else C['ghost']
            add_rect(sl, px, safeY + Inches(0.56 + i * 0.28), pw, Inches(0.26),
                     fill_color=bg, line_color=C['slate'], line_width=0.3, radius=4)
            add_text(sl, f'• {ae}', px + Inches(0.12), safeY + Inches(0.56 + i * 0.28),
                     pw - Inches(0.18), Inches(0.26), font_size=9, color=C['ink'], valign='middle')

    add_footer(sl, sec_n, idx, refs)
    return sl


def render_swot(prs, data, meta, sec_n):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    sl.background.fill.solid()
    sl.background.fill.fore_color.rgb = C['offWhite']
    add_header(sl, 'SWOT Analysis',
               f"{meta.get('drug')} · {meta.get('indication')} · {meta.get('scope_label')} {meta.get('year')}")
    add_home_button(sl)
    add_badge(sl, 'yellow')

    boxes = [
        ('strengths',    'Strengths',    C['teal'],    MX,         Inches(1.24)),
        ('weaknesses',   'Weaknesses',   C['red'],     Inches(6.74), Inches(1.24)),
        ('opportunities','Opportunities',C['navyMid'], MX,         Inches(4.18)),
        ('threats',      'Threats',      C['amberDk'], Inches(6.74), Inches(4.18)),
    ]

    for key, lbl, col, bx, by in boxes:
        bH = Inches(2.78)
        items = (data.get(key) or [])[:5]

        # Card
        add_rect(sl, bx, by, Inches(6.20), bH,
                 fill_color=C['white'], line_color=col, line_width=1.0, radius=10)
        # Header
        add_rect(sl, bx, by, Inches(6.20), Inches(0.44),
                 fill_xml=grad_fill_xml(col, C['navy'], angle_deg=0), radius=8)
        # Label
        add_text(sl, lbl, bx + Inches(0.54), by + Inches(0.06), Inches(5.58), Inches(0.32),
                 font_size=12, font_face=FONT_TITLE, bold=True,
                 color=C['white'], valign='middle')

        # Bullets — each item as separate text box for reliability
        if items:
            y_offset = by + Inches(0.50)
            item_h = (bH - Inches(0.56)) / max(len(items), 1)
            for item in items:
                # Get text robustly
                if isinstance(item, dict):
                    txt = str(item.get('text') or item.get('bullet') or item.get('content') or '')
                else:
                    txt = str(item)
                txt = _trunc(txt.strip(), 200)
                if txt:
                    add_text(sl, f'• {txt}', bx + Inches(0.16), y_offset,
                             Inches(5.90), item_h - Inches(0.02),
                             font_size=9.5, color=C['ink'], valign='top', line_spacing=1.2)
                y_offset += item_h

    add_footer(sl, sec_n, 1)
    return sl


def render_generic(prs, data, label, icon, meta, sec_n):
    """Generic slide for Disease Intro and others."""
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    sl.background.fill.solid()
    sl.background.fill.fore_color.rgb = C['offWhite']
    add_header(sl, f'{icon}  {label}', f"{meta.get('drug')} · {meta.get('scope_label')}")
    add_home_button(sl)
    add_badge(sl, 'green')

    items = (data.get('slides', [{}])[0].get('items', []) if data.get('slides') else [])[:7]
    refs = []
    if not items:
        add_text(sl, f'No data available for {label}',
                 MX, Inches(2.5), CW, Inches(0.5),
                 font_size=11, color=C['stone'], align='center')
        add_footer(sl, sec_n, 1)
        return sl

    avail_h = 5.88
    ih = min(0.94, avail_h / max(len(items), 1))

    icon_list = ['A','B','C','D','E','F','G']
    for i, item in enumerate(items):
        y   = Inches(1.22 + i * (ih + 0.04))
        col = C['teal'] if i % 2 == 0 else C['navyMid']
        bg  = C['white'] if i % 2 == 0 else C['ghost']

        add_rect(sl, MX, y, CW, Inches(ih), fill_color=bg,
                 line_color=col, line_width=0.75, radius=8)
        add_rect(sl, MX, y, Inches(0.12), Inches(ih), fill_color=col, radius=3)
        _ic_lbl = icon_list[i % len(icon_list)]
        add_rect(sl, MX + Inches(0.18), y + Inches((ih-0.36)/2),
                 Inches(0.36), Inches(0.36), fill_color=col, radius=50)
        add_text(sl, _ic_lbl, MX + Inches(0.18), y + Inches((ih-0.36)/2),
                 Inches(0.36), Inches(0.36),
                 font_size=9, bold=True, color=C['white'], align='center', valign='middle')
        add_text(sl, item.get('label', ''),
                 MX + Inches(0.70), y + Inches(0.06), Inches(2.74), Inches(ih - 0.12),
                 font_size=10, font_face=FONT_TITLE, bold=True, color=C['navy'], valign='middle')
        add_line(sl, MX + Inches(3.52), y + Inches(0.08),
                 MX + Inches(3.52), y + Inches(ih - 0.08), C['slate'], 0.5)
        add_text(sl, _strip_refs(item.get('text', '')),
                 MX + Inches(3.64), y + Inches(0.06), CW - Inches(3.72), Inches(ih - 0.12),
                 font_size=9.5, color=C['ink'], valign='middle', line_spacing=1.15)

        if item.get('source'):
            refs.append(item['source'])
            add_text(sl, f'[{len(refs)}]',
                     MX + CW - Inches(0.32), y + Inches(ih - 0.26), Inches(0.28), Inches(0.22),
                     font_size=6.5, color=C['stone'], italic=True, align='right')

    add_footer(sl, sec_n, 1, refs)
    return sl


def render_tactics(prs, data, meta, sec_n):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    sl.background.fill.solid()
    sl.background.fill.fore_color.rgb = C['offWhite']
    add_header(sl, data.get('heading') or f"Medical Affairs Tactical Plan {meta.get('year')}",
               f"{meta.get('drug')} · {meta.get('scope_label')}")
    add_home_button(sl)
    add_badge(sl, 'yellow')

    rows_data = data.get('rows', [])[:8]
    tbl_rows = []
    for row in rows_data:
        tbl_rows.append([
            dcell(row.get('type', ''), bold=True, color=C['navy']),
            dcell(row.get('si', ''), align='center', color=C['stone'], sz=8.5),
            dcell(row.get('tactic', ''), bold=True),
            dcell(_strip_refs(row.get('description', '')), sz=9.5),
        ])
    if tbl_rows:
        add_table(sl,
                  [hcell('Type', align='left'), hcell('SI'), hcell('Tactic', align='left'), hcell('Description', align='left')],
                  tbl_rows, MX, Inches(1.22), CW, [1.60, 0.72, 2.68, 7.63], 0.60)

    add_footer(sl, sec_n, 1)
    return sl


def render_imperatives(prs, data, meta, sec_n):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    sl.background.fill.solid()
    sl.background.fill.fore_color.rgb = C['offWhite']
    add_header(sl, 'Strategic Imperatives',
               f"{meta.get('drug')} · {meta.get('scope_label')} · {meta.get('year')}")
    add_home_button(sl)
    add_badge(sl, 'yellow')

    if data.get('overarching'):
        add_rect(sl, MX, Inches(1.20), CW, Inches(0.44),
                 fill_xml=grad_fill_xml(C['navy'], C['navyMid'], angle_deg=0), radius=8)
        add_text(sl, _trunc(_strip_refs(data['overarching']), 300),
                 MX + Inches(0.14), Inches(1.20), CW - Inches(0.20), Inches(0.44),
                 font_size=10, font_face=FONT_TITLE, bold=True, italic=True,
                 color=C['white'], align='center', valign='middle')

    pillars = data.get('pillars', [])[:4]
    pW = Inches(3.0)
    gap = Inches(0.14)
    start_y = Inches(1.70) if data.get('overarching') else Inches(1.22)
    box_h = Inches(7.08) - start_y - Inches(0.24)
    p_colors = [C['teal'], C['navyMid'], C['navyLt'], C['red']]

    for i, pillar in enumerate(pillars):
        x   = MX + i * (pW + gap)
        col = p_colors[i] if i < len(p_colors) else C['navy']

        add_rect(sl, x, start_y, pW, box_h,
                 fill_color=C['white'], line_color=col, line_width=0.75, radius=10)
        add_rect(sl, x, start_y, pW, Inches(0.52),
                 fill_xml=grad_fill_xml(col, C['navy'], angle_deg=135), radius=8)
        add_text(sl, _trunc(pillar.get('title', ''), 50),
                 x + Inches(0.06), start_y, pW - Inches(0.12), Inches(0.52),
                 font_size=10, font_face=FONT_TITLE, bold=True,
                 color=C['white'], align='center', valign='middle')

        if pillar.get('rationale'):
            add_text(sl, _trunc(_strip_refs(pillar['rationale']), 120),
                     x + Inches(0.08), start_y + Inches(0.56), pW - Inches(0.16), Inches(0.36),
                     font_size=7.5, color=C['stone'], align='center', italic=True)

        add_line(sl, x + Inches(0.08), start_y + Inches(0.96),
                 x + pW - Inches(0.08), start_y + Inches(0.96), C['slate'], 0.5)

        objs = [str(o) for o in (pillar.get('objectives') or [])[:6]]
        y_off = start_y + Inches(1.04)
        item_h = (box_h - Inches(1.10)) / max(len(objs), 1)
        for obj in objs:
            add_text(sl, f'• {_trunc(_strip_refs(obj), 200)}',
                     x + Inches(0.12), y_off, pW - Inches(0.20), item_h - Inches(0.02),
                     font_size=9.5, color=C['ink'], valign='top', line_spacing=1.2)
            y_off += item_h

    add_footer(sl, sec_n, 1)
    return sl


def render_narrative(prs, data, meta, sec_n):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    sl.background.fill.solid()
    sl.background.fill.fore_color.rgb = C['offWhite']
    add_header(sl, 'Scientific Narrative & Key Messaging', f"{meta.get('drug')} · {meta.get('year')}")
    add_home_button(sl)
    add_badge(sl, 'yellow')
    refs = []

    # Primary message
    add_rect(sl, MX, Inches(1.20), CW, Inches(0.64),
             fill_xml=grad_fill_xml(C['navy'], C['navyMid'], angle_deg=0), radius=10)
    add_rect(sl, MX, Inches(1.20), Inches(0.14), Inches(0.64), fill_color=C['teal'], radius=3)
    add_text(sl, 'Primary Message', MX + Inches(0.22), Inches(1.20), Inches(2.2), Inches(0.26),
             font_size=8, bold=True, color=C['teal'], valign='middle')
    add_text(sl, _strip_refs(data.get('primary_message', '')),
             MX + Inches(0.22), Inches(1.46), CW - Inches(0.30), Inches(0.36),
             font_size=9.5, font_face=FONT_TITLE, italic=True, color=C['white'], valign='middle')

    if data.get('key_evidence_statement'):
        add_rect(sl, MX, Inches(1.88), CW, Inches(0.30),
                 fill_color=C['ghost'], line_color=C['teal'], line_width=1.0, radius=6)
        add_text(sl, f"Key Evidence: {_strip_refs(data['key_evidence_statement'])}",
                 MX + Inches(0.12), Inches(1.88), CW - Inches(0.16), Inches(0.30),
                 font_size=9.5, font_face=FONT_TITLE, bold=True, color=C['teal'], valign='middle')

    tps = data.get('talking_points', [])[:4]
    tp_start_y = 2.24 if data.get('key_evidence_statement') else 2.04
    tpW_in = 12.52 / max(len(tps), 1) - 0.10
    tpH = Inches(7.08 - tp_start_y - 0.22)
    tp_colors = [C['navyMid'], C['teal'], C['navyLt'], C['red']]

    for i, tp in enumerate(tps):
        x   = MX + Inches(i * (tpW_in + 0.10))
        col = tp_colors[i] if i < len(tp_colors) else C['navy']

        add_rect(sl, x, Inches(tp_start_y), Inches(tpW_in), Inches(0.36),
                 fill_xml=grad_fill_xml(col, C['navy'], angle_deg=0), radius=6)
        add_text(sl, tp.get('focus', ''), x, Inches(tp_start_y), Inches(tpW_in), Inches(0.36),
                 font_size=9.5, font_face=FONT_TITLE, bold=True,
                 color=C['white'], align='center', valign='middle')

        add_rect(sl, x, Inches(tp_start_y + 0.36), Inches(tpW_in), tpH,
                 fill_color=C['white'], line_color=col, line_width=0.75, radius=8)

        pts = [str(p) for p in (tp.get('points') or [])[:5]]
        y_off = Inches(tp_start_y + 0.44)
        pt_h  = tpH / max(len(pts), 1)
        for pt in pts:
            add_text(sl, f'• {_strip_refs(pt)}',
                     x + Inches(0.10), y_off, Inches(tpW_in - 0.20), pt_h - Inches(0.02),
                     font_size=9.5, color=C['ink'], valign='top', line_spacing=1.2)
            y_off += pt_h

    if data.get('source_statement'):
        refs.append(data['source_statement'])
    add_footer(sl, sec_n, 1, refs)
    return sl


def render_competitive(prs, data, meta, sec_n):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    sl.background.fill.solid()
    sl.background.fill.fore_color.rgb = C['offWhite']
    add_header(sl, data.get('heading') or f"Competitive Landscape: {meta.get('indication')}",
               f"{meta.get('scope_label')} · {meta.get('year')}")
    add_home_button(sl)
    add_badge(sl, 'green')

    rows_data = data.get('rows', [])[:9]
    refs = []
    tbl_rows = []
    for row in rows_data:
        if row.get('notes'):
            refs.append(row['notes'])
        is_focus = bool(row.get('is_focus'))
        fg = C['teal'] if is_focus else C['ink']
        bg = C['ghost'] if is_focus else None  # None = alternating
        tbl_rows.append([
            dcell(_strip_refs(row.get('drug_trial', '')), bold=is_focus, color=fg, fill=bg),
            dcell(row.get('prior_lot', '—'), align='center', color=fg, fill=bg),
            dcell(row.get('mpfs', '—'), align='center', bold=is_focus, color=fg, fill=bg),
            dcell(row.get('hr_pfs', '—'), align='center', color=fg, fill=bg),
            dcell(row.get('orr', '—'), align='center', color=fg, fill=bg),
            dcell(row.get('mos', '—'), align='center', color=fg, fill=bg),
            dcell((f'[{len(refs)}] ' + _trunc(row['notes'], 80)) if row.get('notes') else '—',
                  color=C['stone'], sz=8.5, fill=bg),
        ])

    if tbl_rows:
        add_table(sl,
                  [hcell('Drug + Trial', align='left'), hcell('LOT'), hcell('mPFS'),
                   hcell('HR PFS'), hcell('ORR'), hcell('mOS'), hcell('Reference', align='left')],
                  tbl_rows, MX, Inches(1.20), CW,
                  [2.78, 0.74, 0.82, 0.80, 0.68, 0.80, 6.01], 0.50)

    if data.get('footnote'):
        add_text(sl, _trunc(_strip_refs(data['footnote']), 280),
                 MX, Inches(7.02), CW, Inches(0.22),
                 font_size=7.5, color=C['stone'], italic=True)

    add_footer(sl, sec_n, 1, refs)
    return sl


def render_iep(prs, data, meta, sec_n):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    sl.background.fill.solid()
    sl.background.fill.fore_color.rgb = C['offWhite']
    add_header(sl, data.get('heading') or f"Integrated Evidence Plan: {meta.get('drug')} {meta.get('year')}",
               f"{meta.get('scope_label')} · {meta.get('indication')}")
    add_home_button(sl)
    add_badge(sl, 'green')

    if data.get('scope_note'):
        add_text(sl, _trunc(data['scope_note'], 260), MX, Inches(1.20), CW, Inches(0.22),
                 font_size=9, color=C['stone'], italic=True, valign='middle')

    gaps = data.get('gaps', [])[:7]
    refs = []
    st_icon = {'active': '✓', 'planned': '▲', 'pending': '○'}
    st_col  = {'active': C['teal'], 'planned': C['amberDk'], 'pending': C['stone']}
    tbl_rows = []
    for gap in gaps:
        if gap.get('source'):
            refs.append(gap['source'])
        act = _trunc(_strip_refs(gap.get('activity', '')), 160)
        if gap.get('source'):
            act += f'  [{len(refs)}]'
        status = gap.get('status', 'pending')
        tbl_rows.append([
            dcell(_trunc(_strip_refs(gap.get('gap', '')), 100), bold=True),
            dcell(act),
            dcell(_trunc(gap.get('responsible', 'GMA'), 65), color=C['stone'], sz=9),
            dcell(f"{st_icon.get(status,'○')} {status}",
                  color=st_col.get(status, C['stone']),
                  bold=(status == 'active'), sz=9),
            dcell(_trunc(gap.get('timeline', str(meta.get('year', ''))), 20),
                  align='center', sz=9),
        ])

    start_y = Inches(1.46) if data.get('scope_note') else Inches(1.22)
    if tbl_rows:
        add_table(sl,
                  [hcell('Evidence Gap', align='left'), hcell('Activity & Reference', align='left'),
                   hcell('Responsible', align='left'), hcell('Status'), hcell('Timeline')],
                  tbl_rows, MX, start_y, CW, [2.80, 5.10, 1.96, 1.08, 1.69], 0.66)

    rk_key = f"key_readouts_{meta.get('year', '')}"
    if data.get(rk_key):
        add_text(sl, f"Key readouts {meta.get('year')}: {_trunc(_strip_refs(data[rk_key]), 360)}",
                 MX, Inches(7.02), CW, Inches(0.22),
                 font_size=9.5, bold=True, color=C['teal'])

    add_footer(sl, sec_n, 1, refs)
    return sl


def render_timeline(prs, data, meta, sec_n):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    sl.background.fill.solid()
    sl.background.fill.fore_color.rgb = C['offWhite']
    add_header(sl, f"Medical Affairs Timeline {data.get('year') or meta.get('year')}",
               f"{meta.get('drug')} · {meta.get('scope_label')}")
    add_home_button(sl)
    add_badge(sl, 'yellow')

    lineY = Inches(3.68)
    events = data.get('events', [])[:8]

    # Timeline bar
    add_rect(sl, Inches(0.36), lineY, Inches(12.60), Inches(0.14),
             fill_xml=grad_fill_xml(C['teal'], C['navyMid'], angle_deg=0))

    qPos = {'Q1': 1.74, 'Q2': 5.10, 'Q3': 8.46, 'Q4': 11.82}
    yr   = data.get('year') or meta.get('year', '')
    for q, xv in qPos.items():
        xv_in = Inches(xv)
        # Circle marker
        add_rect(sl, xv_in - Inches(0.18), lineY - Inches(0.16),
                 Inches(0.36), Inches(0.36), fill_color=C['teal'], radius=50)
        add_text(sl, q, xv_in - Inches(0.32), lineY + Inches(0.22),
                 Inches(0.64), Inches(0.24), font_size=9.5, font_face=FONT_TITLE,
                 bold=True, color=C['navy'], align='center')
        add_text(sl, str(yr), xv_in - Inches(0.32), lineY + Inches(0.46),
                 Inches(0.64), Inches(0.20), font_size=7, color=C['stone'], align='center')

    for evt in events:
        xv    = Inches(qPos.get(evt.get('quarter', 'Q1'), 1.74))
        above = evt.get('position', '') == 'above'
        bW, bH = Inches(2.62), Inches(1.04)
        bY = lineY - Inches(2.22) if above else lineY + Inches(0.80)
        # Connector line
        lY1 = bY + bH if above else lineY + Inches(0.14)
        lY2 = lineY if above else bY
        add_line(sl, xv, lY1, xv, lY2, C['teal'], 1.0)

        # Event card
        add_rect(sl, xv - bW/2, bY, bW, bH,
                 fill_color=C['white'], line_color=C['navyMid'], line_width=0.75, radius=8)
        add_rect(sl, xv - bW/2, bY, Inches(0.08), bH, fill_color=C['teal'], radius=3)
        add_text(sl, _trunc(evt.get('event', ''), 60),
                 xv - bW/2 + Inches(0.14), bY + Inches(0.06),
                 bW - Inches(0.22), Inches(0.38),
                 font_size=9, font_face=FONT_TITLE, bold=True, color=C['ink'], line_spacing=1.1)
        if evt.get('detail'):
            add_text(sl, _trunc(evt['detail'], 110),
                     xv - bW/2 + Inches(0.14), bY + Inches(0.46),
                     bW - Inches(0.22), Inches(0.52),
                     font_size=7.5, color=C['stone'], valign='top', line_spacing=1.1)

    if data.get('spark_bar'):
        add_rect(sl, MX, Inches(6.90), CW, Inches(0.26), fill_color=C['navy'], radius=6)
        add_text(sl, _trunc(data['spark_bar'], 300),
                 MX + Inches(0.14), Inches(6.90), CW - Inches(0.20), Inches(0.26),
                 font_size=8.5, color=RGBColor(0xC8, 0xDC, 0xF8), align='center', valign='middle')

    add_footer(sl, sec_n, 1)
    return sl


def render_summary(prs, data, meta, sec_n):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    sl.background.fill.solid()
    sl.background.fill.fore_color.rgb = C['offWhite']
    add_header(sl, 'Summary & Strategic Next Steps',
               f"{meta.get('drug')} · {meta.get('scope_label')} · {meta.get('year')}")
    add_home_button(sl)
    add_badge(sl, 'green')

    msgs = data.get('key_messages', [])[:5]
    mH   = min(1.02, 5.70 / max(len(msgs), 1))
    refs = []
    m_colors = [C['teal'], C['navyMid'], C['navyLt'], C['red'], C['green']]

    for i, m in enumerate(msgs):
        y   = Inches(1.22 + i * (mH + 0.04))
        col = m_colors[i] if i < len(m_colors) else C['navy']
        bg  = C['white'] if i % 2 == 0 else C['ghost']

        add_rect(sl, MX, y, CW, Inches(mH), fill_color=bg,
                 line_color=col, line_width=0.75, radius=8)
        add_rect(sl, MX, y, Inches(0.12), Inches(mH), fill_color=col, radius=3)

        add_text(sl, _strip_refs(m.get('message', '')),
                 MX + Inches(0.72), y + Inches(0.05), CW - Inches(0.80), Inches(mH - 0.12),
                 font_size=10, font_face=FONT_TITLE, bold=True,
                 color=C['ink'], valign='middle', line_spacing=1.1)

        if m.get('source'):
            refs.append(m['source'])

    if data.get('critical_milestones'):
        add_rect(sl, MX, Inches(6.32), CW, Inches(0.34), fill_color=C['navy'], radius=6)
        add_text(sl, f"Key {meta.get('year')} Milestones: {_trunc(_strip_refs(data['critical_milestones']), 300)}",
                 MX + Inches(0.14), Inches(6.32), CW - Inches(0.20), Inches(0.34),
                 font_size=9.5, font_face=FONT_TITLE, bold=True,
                 color=C['white'], valign='middle')

    if data.get('call_to_action'):
        add_rect(sl, MX, Inches(6.72), CW, Inches(0.50),
                 fill_xml=grad_fill_xml(C['teal'], C['tealDk'], angle_deg=0), radius=8)
        add_text(sl, f"→ {_trunc(_strip_refs(data['call_to_action']), 300)}",
                 MX + Inches(0.14), Inches(6.72), CW - Inches(0.20), Inches(0.50),
                 font_size=10, font_face=FONT_TITLE, bold=True,
                 color=C['white'], valign='middle')

    add_footer(sl, sec_n, 1, refs)
    return sl


# ── DISPATCHER ────────────────────────────────────────────────────────────

SECTION_DISPATCH = {
    'executive_summary': render_executive_summary,
    'prevalence_kpi':    render_prevalence,
    'unmet_needs':       render_unmet_needs,
    'guidelines':        render_guidelines,
    'pivotal_table':     render_pivotal,
    'swot':              render_swot,
    'tactics':           render_tactics,
    'imperatives':       render_imperatives,
    'narrative':         render_narrative,
    'competitor_table':  render_competitive,
    'iep':               render_iep,
    'timeline':          render_timeline,
    'summary':           render_summary,
}


def build_pptx(payload: dict) -> bytes:
    meta     = payload.get('meta', {})
    sections = payload.get('sections', [])

    prs = Presentation()
    prs.slide_width  = W
    prs.slide_height = H

    # Title + ToC
    render_title(prs, meta)
    render_toc(prs, sections, meta)

    # Content sections
    for i, sec in enumerate(sections):
        n       = i + 1
        label   = sec.get('label', '')
        icon    = sec.get('icon', '')
        data    = sec.get('data', {})
        sec_type = data.get('section_type', sec.get('id', ''))

        # Divider
        render_divider(prs, n, label, icon)

        # Content
        render_fn = SECTION_DISPATCH.get(sec_type)
        if render_fn:
            result = render_fn(prs, data, meta, n)
            # pivotal returns a list, others return a single slide
            # (already added to prs inside each function)
        else:
            render_generic(prs, data, label, icon, meta, n)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.read()


# ── NETLIFY HANDLER ───────────────────────────────────────────────────────

def handler(event, context):
    try:
        import time as _time; t0 = _time.time()
        if event.get('httpMethod') == 'OPTIONS':
            return {
                'statusCode': 200,
                'headers': {'Access-Control-Allow-Origin': '*',
                            'Access-Control-Allow-Headers': 'Content-Type'},
                'body': ''
            }

        body = event.get('body', '{}')
        if isinstance(body, str):
            payload = json.loads(body)
        else:
            payload = body

        pptx_bytes = build_pptx(payload)
        elapsed_ms = int((time.time() - t0) * 1000)
        b64 = base64.b64encode(pptx_bytes).decode('utf-8')

        return {
            'statusCode': 200,
            'headers': {
                'Content-Type': 'application/json',
                'Access-Control-Allow-Origin': '*',
            },
            'body': json.dumps({'pptx': b64, 'size': len(pptx_bytes), 'elapsed_ms': elapsed_ms})
        }

    except Exception as e:
        tb = traceback.format_exc()
        return {
            'statusCode': 500,
            'headers': {'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*'},
            'body': json.dumps({'error': str(e), 'traceback': tb})
        }
