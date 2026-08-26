"""
RIOS — Gerador de POP em Word (.docx)
Segue o "Prompt Mestre POP Padrão" (Chef Marco À Souza):
capa técnica, histórico de revisão, marca-d'água, cabeçalho institucional
repetido nas páginas internas, seções numeradas, tabelas padronizadas,
fluxograma em página própria, campos de paginação automática.
"""
import io
import os
import re
import unicodedata

from docx import Document
from docx.shared import Cm, Pt, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from PIL import Image, ImageEnhance

from rios_flow_png import generate_flow_png, hex_to_rgb as _hex_to_rgb_tuple

DEFAULT_COLOR = "5E4778"
FONT = "Arial"


# ── Helpers de cor / texto ──────────────────────────────────────────────
def norm_color(hex_color):
    h = (hex_color or DEFAULT_COLOR).lstrip('#').strip()
    if len(h) != 6 or any(c not in '0123456789abcdefABCDEF' for c in h):
        h = DEFAULT_COLOR
    return h.upper()


def rgb_of(hex_color):
    h = norm_color(hex_color)
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def slugify(text, maxlen=40):
    text = str(text or "POP")
    text = unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode('ascii')
    text = re.sub(r'[^A-Za-z0-9]+', '_', text).strip('_')
    return (text or "POP")[:maxlen]


# ── Helpers OOXML de baixo nível ────────────────────────────────────────
def set_cell_bg(cell, hex_color):
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), norm_color(hex_color))
    cell._tc.get_or_add_tcPr().append(shd)


def set_cell_borders(cell, sz=6, color="000000", edges=('top', 'left', 'bottom', 'right')):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = tcPr.find(qn('w:tcBorders'))
    if borders is None:
        borders = OxmlElement('w:tcBorders')
        tcPr.append(borders)
    for edge in edges:
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), str(sz))
        el.set(qn('w:space'), '0')
        el.set(qn('w:color'), color)
        borders.append(el)


def set_cell_margins(cell, top=60, bottom=60, left=100, right=100):
    tcPr = cell._tc.get_or_add_tcPr()
    mar = OxmlElement('w:tcMar')
    for tag, val in (('top', top), ('bottom', bottom), ('start', left), ('end', right)):
        el = OxmlElement(f'w:{tag}')
        el.set(qn('w:w'), str(val))
        el.set(qn('w:type'), 'dxa')
        mar.append(el)
    tcPr.append(mar)


def set_cell_vcenter(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    va = OxmlElement('w:vAlign')
    va.set(qn('w:val'), 'center')
    tcPr.append(va)


def no_table_autofit(table):
    tblPr = table._tbl.tblPr
    layout = OxmlElement('w:tblLayout')
    layout.set(qn('w:type'), 'fixed')
    tblPr.append(layout)


def set_col_widths(table, widths_cm):
    no_table_autofit(table)
    table.autofit = False
    for row in table.rows:
        for idx, w in enumerate(widths_cm):
            if idx < len(row.cells):
                row.cells[idx].width = Cm(w)
    grid = table._tbl.find(qn('w:tblGrid'))
    if grid is not None:
        for idx, gridcol in enumerate(grid.findall(qn('w:gridCol'))):
            if idx < len(widths_cm):
                gridcol.set(qn('w:w'), str(int(widths_cm[idx] * 567)))


def add_field(paragraph, field_code):
    """Insere um campo de Word (ex.: PAGE, NUMPAGES) — atualiza sozinho no Word."""
    run = paragraph.add_run()
    fld_begin = OxmlElement('w:fldChar')
    fld_begin.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve')
    instr.text = f' {field_code} '
    fld_sep = OxmlElement('w:fldChar')
    fld_sep.set(qn('w:fldCharType'), 'separate')
    txt = OxmlElement('w:t')
    txt.text = '1'
    fld_end = OxmlElement('w:fldChar')
    fld_end.set(qn('w:fldCharType'), 'end')
    run._r.append(fld_begin)
    run2 = paragraph.add_run(); run2._r.append(instr)
    run3 = paragraph.add_run(); run3._r.append(fld_sep)
    run3._r.append(txt)
    run4 = paragraph.add_run(); run4._r.append(fld_end)


def add_right_tab(paragraph, section):
    usable = section.page_width - section.left_margin - section.right_margin
    paragraph.paragraph_format.tab_stops.add_tab_stop(usable, WD_TAB_ALIGNMENT.RIGHT)


def add_bottom_border(paragraph, sz=16, color="000000", space=6):
    pPr = paragraph._p.get_or_add_pPr()
    pbdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), str(sz))
    bottom.set(qn('w:space'), str(space))
    bottom.set(qn('w:color'), color)
    pbdr.append(bottom)
    pPr.append(pbdr)


def set_keep_with_next(paragraph):
    pPr = paragraph._p.get_or_add_pPr()
    el = OxmlElement('w:keepNext')
    pPr.append(el)


def add_anchor_picture(paragraph, image_path, width_emu, height_emu, page_w_emu, page_h_emu,
                        offset_y_frac=0.62, behind=True):
    """Insere imagem flutuante (atrás do texto), centralizada horizontalmente, na
    fração vertical indicada da página (0 = topo, 1 = base)."""
    run = paragraph.add_run()
    run.add_picture(image_path, width=Emu(width_emu), height=Emu(height_emu))
    drawing = run._r.find(qn('w:drawing'))
    inline = drawing.find(qn('wp:inline'))
    extent = inline.find(qn('wp:extent'))
    docpr = inline.find(qn('wp:docPr'))
    graphic = inline.find(qn('a:graphic'))

    anchor = OxmlElement('wp:anchor')
    anchor.set('distT', '0'); anchor.set('distB', '0'); anchor.set('distL', '0'); anchor.set('distR', '0')
    anchor.set('simplePos', '0')
    anchor.set('relativeHeight', '2')
    anchor.set('behindDoc', '1' if behind else '0')
    anchor.set('locked', '0')
    anchor.set('layoutInCell', '1')
    anchor.set('allowOverlap', '1')

    simplepos = OxmlElement('wp:simplePos'); simplepos.set('x', '0'); simplepos.set('y', '0')
    posh = OxmlElement('wp:positionH'); posh.set('relativeFrom', 'page')
    align_h = OxmlElement('wp:align'); align_h.text = 'center'
    posh.append(align_h)
    posv = OxmlElement('wp:positionV'); posv.set('relativeFrom', 'page')
    offv = OxmlElement('wp:posOffset')
    offv.text = str(int(page_h_emu * offset_y_frac - height_emu / 2))
    posv.append(offv)

    wrap = OxmlElement('wp:wrapNone')

    anchor.append(simplepos)
    anchor.append(posh)
    anchor.append(posv)
    anchor.append(extent)
    wrap_ok = inline.find(qn('wp:wrapNone'))
    anchor.append(wrap)
    anchor.append(docpr)
    cnv = OxmlElement('wp:cNvGraphicFramePr')
    anchor.append(cnv)
    anchor.append(graphic)

    drawing.remove(inline)
    drawing.append(anchor)


def make_watermark_image(src_path, out_path, opacity=0.10):
    """Clareia o logo para uso como marca-d'água (compõe sobre branco)."""
    img = Image.open(src_path).convert("RGBA")
    white_bg = Image.new("RGBA", img.size, (255, 255, 255, 255))
    faded = Image.blend(white_bg, img, opacity)
    faded = faded.convert("RGB")
    faded.save(out_path)
    return out_path


def make_placeholder_logo(out_path, text="LOGOTIPO\nDO CLIENTE", size=(600, 300)):
    img = Image.new("RGB", size, "white")
    from PIL import ImageDraw
    d = ImageDraw.Draw(img)
    d.rectangle([4, 4, size[0]-4, size[1]-4], outline=(150, 150, 150), width=3)
    from rios_flow_png import _load_font
    font = _load_font(30, bold=True)
    lines = text.split("\n")
    total_h = sum([font.size + 8 for _ in lines])
    y = (size[1]-total_h)//2
    for line in lines:
        bbox = d.textbbox((0, 0), line, font=font)
        w = bbox[2]-bbox[0]
        d.text(((size[0]-w)//2, y), line, font=font, fill=(150, 150, 150))
        y += font.size + 8
    img.save(out_path)
    return out_path


# ── Construção do documento ─────────────────────────────────────────────
def build_pop_docx(pop, cliente, out_path, logo_path=None, color_hex=None, tmp_dir=None):
    tmp_dir = tmp_dir or os.path.dirname(out_path) or "."
    color_hex = norm_color(color_hex or (cliente or {}).get('corHex') or DEFAULT_COLOR)
    cliente_nome = (cliente or {}).get('nome') or pop.get('cliente') or 'Cliente'
    sistema_nome = (cliente or {}).get('sistemaGestao') or f"Sistema de Gestão {cliente_nome}"

    revisoes = pop.get('revisoes') or []
    ultima_rev = revisoes[-1] if revisoes else {}
    rev_num = str(ultima_rev.get('numero', 0))
    rev_data = ultima_rev.get('data') or pop.get('atualizadoEm', '')[:10] or ''
    aprov_nome = ultima_rev.get('aprovadoPor') or '—'

    codigo = pop.get('codigo') or 'POP-0000'
    titulo = pop.get('titulo') or 'Procedimento Operacional Padrão'
    setor = pop.get('setor') or '—'

    # logo real ou placeholder
    logo_final = logo_path
    if not logo_final or not os.path.exists(logo_final):
        logo_final = os.path.join(tmp_dir, "_placeholder_logo.png")
        make_placeholder_logo(logo_final)

    watermark_path = os.path.join(tmp_dir, "_watermark.png")
    make_watermark_image(logo_final, watermark_path, opacity=0.10)

    doc = Document()
    style = doc.styles['Normal']
    style.font.name = FONT
    style.font.size = Pt(11)
    style.font.color.rgb = RGBColor(0x11, 0x11, 0x11)
    style.paragraph_format.line_spacing = 1.12
    style.paragraph_format.space_after = Pt(4)

    # ── Seção 0 — CAPA ──────────────────────────────────────────────
    sec0 = doc.sections[0]
    sec0.page_height = Cm(29.7)
    sec0.page_width = Cm(21.0)
    sec0.left_margin = Cm(1.2)
    sec0.right_margin = Cm(1.2)
    sec0.top_margin = Cm(1.6)
    sec0.bottom_margin = Cm(1.5)

    page_w_emu = int(sec0.page_width)
    page_h_emu = int(sec0.page_height)

    # tabela técnica (3 colunas)
    tbl = doc.add_table(rows=1, cols=3)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    usable_cm = 21.0 - 1.2 - 1.2
    w_logo = usable_cm * 0.23
    w_mid = usable_cm * 0.50
    w_meta = usable_cm - w_logo - w_mid
    set_col_widths(tbl, [w_logo, w_mid, w_meta])

    c_logo, c_mid, c_meta = tbl.rows[0].cells
    for c in (c_logo, c_mid, c_meta):
        set_cell_borders(c, sz=10, color="000000")
        set_cell_vcenter(c)
        set_cell_margins(c)

    p_logo = c_logo.paragraphs[0]
    p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_logo = p_logo.add_run()
    with Image.open(logo_final) as im:
        iw, ih = im.size
    max_w_cm, max_h_cm = 4.0, 2.2
    ratio = min(max_w_cm / (iw / 96 * 2.54), max_h_cm / (ih / 96 * 2.54)) if iw and ih else 1
    # calcula proporcional simples via razão da imagem
    disp_w = max_w_cm
    disp_h = max_w_cm * ih / iw if iw else max_h_cm
    if disp_h > max_h_cm:
        disp_h = max_h_cm
        disp_w = max_h_cm * iw / ih if ih else max_w_cm
    r_logo.add_picture(logo_final, width=Cm(disp_w), height=Cm(disp_h))

    p1 = c_mid.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run1 = p1.add_run(sistema_nome)
    run1.bold = True
    run1.font.size = Pt(14)
    run1.font.name = FONT
    p2 = c_mid.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run2 = p2.add_run(titulo.upper())
    run2.bold = False
    run2.font.size = Pt(16)
    run2.font.name = FONT

    meta_lines = [
        ("Cód.: ", codigo),
        ("Rev.: ", rev_num),
        ("Data: ", rev_data or '—'),
    ]
    first = True
    for label, value in meta_lines:
        pm = c_meta.paragraphs[0] if first else c_meta.add_paragraph()
        first = False
        rl = pm.add_run(label); rl.bold = True; rl.font.size = Pt(10); rl.font.name = FONT
        rv = pm.add_run(value); rv.font.size = Pt(10); rv.font.name = FONT
    pm = c_meta.add_paragraph()
    rl = pm.add_run("Páginas: "); rl.bold = True; rl.font.size = Pt(10); rl.font.name = FONT
    rv = pm.add_run("1 / "); rv.font.size = Pt(10); rv.font.name = FONT
    add_field(pm, "NUMPAGES")
    for r in pm.runs[-3:]:
        r.font.size = Pt(10); r.font.name = FONT

    doc.add_paragraph()

    # histórico de revisão
    ph = doc.add_paragraph()
    rh = ph.add_run("Histórico de Revisão")
    rh.bold = True; rh.font.size = Pt(12); rh.font.name = FONT
    rh.font.color.rgb = rgb_of(color_hex)

    rev_tbl = doc.add_table(rows=1, cols=5)
    set_col_widths(rev_tbl, [1.8, usable_cm - 1.8 - 4.2 - 4.2 - 3.0, 4.2, 4.2, 3.0])
    hdr = rev_tbl.rows[0].cells
    headers = ["Revisão", "Descrição", "Revisado por", "Aprovado por", "Data"]
    for i, htext in enumerate(headers):
        set_cell_bg(hdr[i], color_hex)
        set_cell_borders(hdr[i], sz=6, color="666666")
        set_cell_margins(hdr[i])
        set_cell_vcenter(hdr[i])
        pp = hdr[i].paragraphs[0]
        pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        rr = pp.add_run(htext); rr.bold = True; rr.font.color.rgb = RGBColor(255, 255, 255)
        rr.font.size = Pt(10); rr.font.name = FONT

    rows_data = revisoes if revisoes else [{"numero": 0, "descricao": "Revisão inicial",
                                             "revisadoPor": "", "aprovadoPor": "", "data": ""}]
    for rdata in rows_data:
        row = rev_tbl.add_row().cells
        vals = [str(rdata.get('numero', 0)).zfill(2), rdata.get('descricao', ''),
                rdata.get('revisadoPor', ''), rdata.get('aprovadoPor', ''), rdata.get('data', '')]
        for i, v in enumerate(vals):
            set_cell_borders(row[i], sz=6, color="999999")
            set_cell_margins(row[i])
            set_cell_vcenter(row[i])
            pp = row[i].paragraphs[0]
            rr = pp.add_run(v); rr.font.size = Pt(9.5); rr.font.name = FONT

    # marca-d'água (atrás do texto, centro-inferior da capa)
    wm_p = doc.add_paragraph()
    wm_w_cm = usable_cm * 0.5
    with Image.open(watermark_path) as im:
        iw, ih = im.size
    wm_h_cm = wm_w_cm * ih / iw if iw else wm_w_cm
    add_anchor_picture(wm_p, watermark_path, Emu(int(Cm(wm_w_cm))), Emu(int(Cm(wm_h_cm))),
                        page_w_emu, page_h_emu, offset_y_frac=0.70, behind=True)

    # quebra de página real
    doc.add_page_break()

    # ── Seção 1 — corpo (cabeçalho institucional repetido) ─────────
    new_sec = doc.add_section(WD_SECTION.NEW_PAGE)
    new_sec.page_height = Cm(29.7)
    new_sec.page_width = Cm(21.0)
    new_sec.left_margin = Cm(1.2)
    new_sec.right_margin = Cm(1.2)
    new_sec.top_margin = Cm(3.0)     # espaço reservado para o cabeçalho não sobrepor o corpo
    new_sec.bottom_margin = Cm(1.5)
    new_sec.header_distance = Cm(0.6)

    new_sec.header.is_linked_to_previous = False
    header = new_sec.header
    for hp in header.paragraphs:
        hp.text = ""

    htbl = header.add_table(rows=1, cols=2, width=Cm(usable_cm))
    htbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_col_widths(htbl, [usable_cm * 0.72, usable_cm * 0.28])
    for cell in htbl.rows[0].cells:
        set_cell_borders(cell, sz=0, color="FFFFFF", edges=())
        set_cell_margins(cell, top=0, bottom=0, left=0, right=0)
        set_cell_vcenter(cell)
    hc_left, hc_right = htbl.rows[0].cells
    p_set = hc_left.paragraphs[0]
    r_set = p_set.add_run(setor.upper())
    r_set.bold = True; r_set.font.size = Pt(12); r_set.font.name = FONT
    p_sub = hc_left.add_paragraph()
    r_sub = p_sub.add_run("Procedimento Operacional")
    r_sub.font.size = Pt(10); r_sub.font.name = FONT

    p_logo2 = hc_right.paragraphs[0]
    p_logo2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_logo2 = p_logo2.add_run()
    disp_w2 = 2.5
    disp_h2 = disp_w2 * ih / iw if iw else 1.2
    with Image.open(logo_final) as im2:
        iw2, ih2 = im2.size
    disp_h2 = disp_w2 * ih2 / iw2 if iw2 else 1.2
    if disp_h2 > 1.4:
        disp_h2 = 1.4
        disp_w2 = disp_h2 * iw2 / ih2 if ih2 else 2.5
    r_logo2.add_picture(logo_final, width=Cm(disp_w2), height=Cm(disp_h2))

    p_rule = header.add_paragraph()
    add_bottom_border(p_rule, sz=16, color=color_hex, space=4)

    p_info = header.add_paragraph()
    left_txt = f"{codigo}    REV.{rev_num}    Data: {rev_data or '—'}"
    r_left = p_info.add_run(left_txt)
    r_left.font.size = Pt(8); r_left.font.color.rgb = RGBColor(0xA6, 0xA6, 0xA6); r_left.font.name = FONT
    add_right_tab(p_info, new_sec)
    p_info.add_run("\t")
    r_right = p_info.add_run("Aprovação: ")
    r_right.bold = True; r_right.font.size = Pt(8); r_right.font.color.rgb = RGBColor(0xA6, 0xA6, 0xA6)
    r_right.font.name = FONT
    r_right2 = p_info.add_run(aprov_nome)
    r_right2.font.size = Pt(8); r_right2.font.color.rgb = RGBColor(0xA6, 0xA6, 0xA6); r_right2.font.name = FONT

    new_sec.footer.is_linked_to_previous = False
    for fp in new_sec.footer.paragraphs:
        fp.text = ""

    # ── Corpo: seções numeradas ──────────────────────────────────────
    counter = {"n": 0}

    def add_heading(title):
        counter["n"] += 1
        p = doc.add_paragraph()
        set_keep_with_next(p)
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(4)
        r = p.add_run(f"{counter['n']}. {title}")
        r.bold = True; r.font.size = Pt(14); r.font.name = FONT
        r.font.color.rgb = RGBColor(0x11, 0x11, 0x11)
        return p

    def add_body_text(text, justify=True):
        p = doc.add_paragraph()
        if justify:
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r = p.add_run(text or "Não aplicável.")
        r.font.size = Pt(11); r.font.name = FONT
        return p

    def add_bullets(items, bold_prefix=None):
        items = items or []
        if not items:
            add_body_text("Não aplicável.", justify=False)
            return
        for it in items:
            p = doc.add_paragraph(style='List Bullet')
            r = p.add_run(it)
            r.font.size = Pt(11); r.font.name = FONT

    def add_std_table(headers, rows, col_widths_cm):
        t = doc.add_table(rows=1, cols=len(headers))
        set_col_widths(t, col_widths_cm)
        hdr_cells = t.rows[0].cells
        for i, htext in enumerate(headers):
            set_cell_bg(hdr_cells[i], color_hex)
            set_cell_borders(hdr_cells[i], sz=6, color="666666")
            set_cell_margins(hdr_cells[i], top=40, bottom=40, left=80, right=80)
            set_cell_vcenter(hdr_cells[i])
            pp = hdr_cells[i].paragraphs[0]
            pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            rr = pp.add_run(htext); rr.bold = True; rr.font.color.rgb = RGBColor(255, 255, 255)
            rr.font.size = Pt(9.5); rr.font.name = FONT
        # repetir cabeçalho em novas páginas
        trPr = t.rows[0]._tr.get_or_add_trPr()
        th = OxmlElement('w:tblHeader'); th.set(qn('w:val'), 'true')
        trPr.append(th)
        if not rows:
            row = t.add_row().cells
            row[0].merge(row[-1])
            set_cell_borders(row[0], sz=6, color="999999")
            pp = row[0].paragraphs[0]
            rr = pp.add_run("Não aplicável"); rr.font.size = Pt(9.5); rr.font.name = FONT
            return t
        for rdata in rows:
            row = t.add_row().cells
            for i, v in enumerate(rdata):
                set_cell_borders(row[i], sz=6, color="999999")
                set_cell_margins(row[i], top=40, bottom=40, left=80, right=80)
                set_cell_vcenter(row[i])
                pp = row[i].paragraphs[0]
                rr = pp.add_run(str(v)); rr.font.size = Pt(9.5); rr.font.name = FONT
        return t

    usable_body_cm = 21.0 - 1.2 - 1.2

    # 1. Objetivo
    add_heading("Objetivo")
    add_body_text(pop.get('objetivo'))

    # 2. Campo de aplicação
    add_heading("Campo de aplicação")
    add_body_text(pop.get('aplicacao'))

    # 3. Referências
    add_heading("Referências")
    add_bullets(pop.get('referencias'))

    # 4. Definições
    add_heading("Definições")
    defs = pop.get('definicoes') or []
    if not defs:
        add_body_text("Não aplicável.", justify=False)
    else:
        for d in defs:
            p = doc.add_paragraph(style='List Bullet')
            rb = p.add_run(f"{d.get('termo','')}: "); rb.bold = True; rb.font.size = Pt(11); rb.font.name = FONT
            rv = p.add_run(d.get('definicao', '')); rv.font.size = Pt(11); rv.font.name = FONT

    # 5. Responsabilidades
    add_heading("Responsabilidades")
    resp = pop.get('responsaveis') or []
    if not resp:
        add_body_text("Não aplicável.", justify=False)
    else:
        for r in resp:
            p = doc.add_paragraph(style='List Bullet')
            rb = p.add_run(f"{r.get('papel','')}: "); rb.bold = True; rb.font.size = Pt(11); rb.font.name = FONT
            rv = p.add_run('; '.join(r.get('tarefas') or [])); rv.font.size = Pt(11); rv.font.name = FONT

    # 6. Materiais, utensílios, equipamentos e EPIs
    add_heading("Materiais, utensílios, equipamentos e EPIs")
    add_bullets(pop.get('materiais'))

    # 7. Descrição do procedimento
    add_heading("Descrição do procedimento")
    if pop.get('frequencia'):
        p = doc.add_paragraph()
        rb = p.add_run("Frequência e horário: "); rb.bold = True; rb.font.size = Pt(11); rb.font.name = FONT
        rv = p.add_run(pop['frequencia']); rv.font.size = Pt(11); rv.font.name = FONT
    etapas = pop.get('etapas') or []
    if not etapas:
        add_body_text("Não aplicável.", justify=False)
    else:
        for i, e in enumerate(etapas):
            p = doc.add_paragraph()
            set_keep_with_next(p)
            p.paragraph_format.space_before = Pt(6)
            destaque = e.get('destaque')
            tag = {"critico": " [PONTO CRÍTICO]", "regra": " [REGRA PRINCIPAL]",
                   "financeiro": " [PCC FINANCEIRO]"}.get(destaque, "")
            rb = p.add_run(f"{i+1}. {e.get('titulo','')}{tag}")
            rb.bold = True; rb.font.size = Pt(11); rb.font.name = FONT
            if tag:
                rb.font.color.rgb = RGBColor(0xB0, 0x00, 0x20) if destaque == 'critico' else \
                                     RGBColor(0xC9, 0x74, 0x0A) if destaque == 'regra' else RGBColor(0xA6, 0x8B, 0x00)
            for item in (e.get('itens') or []):
                pi = doc.add_paragraph(style='List Bullet 2')
                ri = pi.add_run(item); ri.font.size = Pt(10.5); ri.font.name = FONT

    # 8. Pontos críticos de controle (PCC)
    add_heading("Pontos críticos de controle (PCC)")
    pcc = pop.get('pcc') or []
    rows = [[p.get('etapa', ''), p.get('risco', ''), p.get('controle', '')] for p in pcc]
    add_std_table(["Etapa", "Risco", "Controle"], rows,
                  [usable_body_cm*0.28, usable_body_cm*0.36, usable_body_cm*0.36])

    # 9. Registros obrigatórios
    add_heading("Registros obrigatórios")
    add_bullets(pop.get('registros'))

    # 10. Indicadores, rendimento ou critérios de aceitação
    add_heading("Indicadores, rendimento ou critérios de aceitação")
    add_bullets(pop.get('indicadores'))

    if pop.get('obs'):
        p = doc.add_paragraph()
        rb = p.add_run("Observações complementares: "); rb.bold = True; rb.font.size = Pt(11); rb.font.name = FONT
        add_body_text(pop['obs'])

    # 11. Fluxo do processo — página própria
    doc.add_page_break()
    add_heading("Fluxo do processo")
    if etapas:
        flow_path = os.path.join(tmp_dir, "_flow.png")
        generate_flow_png(etapas, color_hex, flow_path, width=1600)
        with Image.open(flow_path) as im:
            fw, fh = im.size
        max_w_cm = usable_body_cm
        disp_w = max_w_cm
        disp_h = disp_w * fh / fw
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run()
        r.add_picture(flow_path, width=Cm(disp_w), height=Cm(disp_h))
    else:
        add_body_text("Não aplicável — nenhuma etapa cadastrada.", justify=False)

    doc.save(out_path)
    return out_path


def docx_filename(pop):
    revisoes = pop.get('revisoes') or []
    rev_num = str(revisoes[-1].get('numero', 0)) if revisoes else '0'
    codigo = slugify(pop.get('codigo') or 'POP-0000', 20)
    titulo_curto = slugify(' '.join((pop.get('titulo') or 'POP').split()[:4]), 30)
    return f"POP_{codigo}_{titulo_curto}_REV_{rev_num}.docx"
