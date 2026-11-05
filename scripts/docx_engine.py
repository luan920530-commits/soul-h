# -*- coding: utf-8 -*-
"""
Motor de composición para el informe técnico-jurídico (python-docx).
Define estilos, numeración multinivel, tablas con sombreado, cajas de énfasis,
encabezado/pie con campos, numeración de páginas y utilidades de contenido.
"""
import os
import datetime
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml.ns import qn, nsmap
from docx.oxml import OxmlElement

# ----------------------------------------------------------------------------
# Paleta institucional
# ----------------------------------------------------------------------------
AZUL = RGBColor(0x1F, 0x33, 0x55)        # azul tinta (títulos)
VINO = RGBColor(0x8C, 0x1C, 0x13)        # rojo institucional / alertas
GRIS = RGBColor(0x44, 0x4B, 0x55)        # texto secundario
GRIS_CLARO = RGBColor(0x6B, 0x72, 0x7C)
NEGRO = RGBColor(0x1A, 0x1A, 0x1A)
BLANCO = RGBColor(0xFF, 0xFF, 0xFF)

HEX_AZUL = "1F3355"
HEX_AZUL_MED = "2E4A6B"
HEX_CELESTE = "DCE4EE"
HEX_CELESTE2 = "EDF1F6"
HEX_ARENA = "F5F2EA"
HEX_VINO = "8C1C13"
HEX_VERDE = "1E5B3F"
HEX_AMBAR = "8A6A17"

FUENTE = "Times New Roman"
FUENTE_SANS = "Arial"


# ----------------------------------------------------------------------------
# Helpers XML
# ----------------------------------------------------------------------------
ORDEN_TCPR = ['w:cnfStyle', 'w:tcW', 'w:gridSpan', 'w:hMerge', 'w:vMerge', 'w:tcBorders',
              'w:shd', 'w:noWrap', 'w:tcMar', 'w:textDirection', 'w:tcFitText', 'w:vAlign',
              'w:hideMark', 'w:headers', 'w:cellIns', 'w:cellDel', 'w:cellMerge', 'w:tcPrChange']
ORDEN_TRPR = ['w:cnfStyle', 'w:divId', 'w:gridBefore', 'w:gridAfter', 'w:wBefore', 'w:wAfter',
              'w:cantSplit', 'w:trHeight', 'w:tblHeader', 'w:tblCellSpacing', 'w:jc', 'w:hidden',
              'w:ins', 'w:del', 'w:trPrChange']
ORDEN_TBLPR = ['w:tblStyle', 'w:tblpPr', 'w:tblOverlap', 'w:bidiVisual', 'w:tblStyleRowBandSize',
               'w:tblStyleColBandSize', 'w:tblW', 'w:jc', 'w:tblCellSpacing', 'w:tblInd',
               'w:tblBorders', 'w:shd', 'w:tblLayout', 'w:tblCellMar', 'w:tblLook',
               'w:tblCaption', 'w:tblDescription', 'w:tblPrChange']
ORDEN_PPR = ['w:pStyle', 'w:keepNext', 'w:keepLines', 'w:pageBreakBefore', 'w:framePr',
             'w:widowControl', 'w:numPr', 'w:suppressLineNumbers', 'w:pBdr', 'w:shd', 'w:tabs',
             'w:suppressAutoHyphens', 'w:kinsoku', 'w:wordWrap', 'w:overflowPunct',
             'w:topLinePunct', 'w:autoSpaceDE', 'w:autoSpaceDN', 'w:bidi', 'w:adjustRightInd',
             'w:snapToGrid', 'w:spacing', 'w:ind', 'w:contextualSpacing', 'w:mirrorIndents',
             'w:suppressOverlap', 'w:jc', 'w:textDirection', 'w:textAlignment',
             'w:textboxTightWrap', 'w:outlineLvl', 'w:divId', 'w:cnfStyle', 'w:rPr',
             'w:sectPr', 'w:pPrChange']
ORDEN_RPR = ['w:rStyle', 'w:rFonts', 'w:b', 'w:bCs', 'w:i', 'w:iCs', 'w:caps', 'w:smallCaps',
             'w:strike', 'w:dstrike', 'w:outline', 'w:shadow', 'w:emboss', 'w:imprint',
             'w:noProof', 'w:snapToGrid', 'w:vanish', 'w:webHidden', 'w:color', 'w:spacing',
             'w:w', 'w:kern', 'w:position', 'w:sz', 'w:szCs', 'w:highlight', 'w:u', 'w:effect',
             'w:bdr', 'w:shd', 'w:fitText', 'w:vertAlign', 'w:rtl', 'w:cs', 'w:em', 'w:lang',
             'w:eastAsianLayout', 'w:specVanish', 'w:oMath']


def _ins(parent, child, orden):
    """Inserta `child` respetando el orden del esquema OOXML."""
    tag = child.tag
    if tag not in orden:
        parent.append(child)
        return child
    idx = orden.index(tag)
    for existente in list(parent):
        if existente.tag in orden and orden.index(existente.tag) > idx:
            existente.addprevious(child)
            return child
    parent.append(child)
    return child


def _quitar(parent, tag):
    for old in parent.findall(qn(tag)):
        parent.remove(old)


def _el(tag, **attrs):
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(qn(k), v)
    return e


def sombrear(celda, hex_color):
    tcPr = celda._tc.get_or_add_tcPr()
    _quitar(tcPr, 'w:shd')
    _ins(tcPr, _el('w:shd', **{'w:val': 'clear', 'w:color': 'auto', 'w:fill': hex_color}), ORDEN_TCPR)


def bordes_tabla(tabla, sz=4, color="B7C0CC", interior=True):
    tbl = tabla._tbl
    tblPr = tbl.tblPr
    _quitar(tblPr, 'w:tblBorders')
    b = OxmlElement('w:tblBorders')
    for lado in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        if not interior and lado.startswith('inside'):
            continue
        b.append(_el('w:' + lado, **{'w:val': 'single', 'w:sz': str(sz),
                                     'w:space': '0', 'w:color': color}))
    _ins(tblPr, b, ORDEN_TBLPR)


def bordes_celda(celda, color="B7C0CC", sz=4, lados=('top', 'left', 'bottom', 'right'), val='single'):
    tcPr = celda._tc.get_or_add_tcPr()
    _quitar(tcPr, 'w:tcBorders')
    b = OxmlElement('w:tcBorders')
    for lado in lados:
        b.append(_el('w:' + lado, **{'w:val': val, 'w:sz': str(sz),
                                     'w:space': '0', 'w:color': color}))
    _ins(tcPr, b, ORDEN_TCPR)


def margenes_celda(celda, top=80, bottom=80, left=110, right=110):
    tcPr = celda._tc.get_or_add_tcPr()
    _quitar(tcPr, 'w:tcMar')
    m = OxmlElement('w:tcMar')
    for k, v in (('top', top), ('left', left), ('bottom', bottom), ('right', right)):
        m.append(_el('w:' + k, **{'w:w': str(v), 'w:type': 'dxa'}))
    _ins(tcPr, m, ORDEN_TCPR)


def no_dividir_fila(fila):
    trPr = fila._tr.get_or_add_trPr()
    _quitar(trPr, 'w:cantSplit')
    _ins(trPr, _el('w:cantSplit'), ORDEN_TRPR)


def repetir_encabezado(fila):
    trPr = fila._tr.get_or_add_trPr()
    _quitar(trPr, 'w:tblHeader')
    _ins(trPr, _el('w:tblHeader'), ORDEN_TRPR)


def ancho_fijo(tabla):
    tblPr = tabla._tbl.tblPr
    _quitar(tblPr, 'w:tblLayout')
    _ins(tblPr, _el('w:tblLayout', **{'w:type': 'fixed'}), ORDEN_TBLPR)


def set_col_widths(tabla, anchos_cm):
    tabla.autofit = False
    ancho_fijo(tabla)
    for row in tabla.rows:
        for idx, w in enumerate(anchos_cm):
            if idx < len(row.cells):
                row.cells[idx].width = Cm(w)
    # grid
    tblGrid = tabla._tbl.find(qn('w:tblGrid'))
    if tblGrid is not None:
        for i, gc in enumerate(tblGrid.findall(qn('w:gridCol'))):
            if i < len(anchos_cm):
                gc.set(qn('w:w'), str(int(anchos_cm[i] * 567)))


# ----------------------------------------------------------------------------
# Numeración multinivel (1 / 1.1 / 1.1.1)
# ----------------------------------------------------------------------------
def _crear_numeracion(doc, abstract_id=90, num_id=90):
    numbering = doc.part.numbering_part.element
    # abstractNum
    abs_num = _el('w:abstractNum', **{'w:abstractNumId': str(abstract_id)})
    abs_num.append(_el('w:multiLevelType', **{'w:val': 'multilevel'}))
    for ilvl in range(1, 5):
        lvl = _el('w:lvl', **{'w:ilvl': str(ilvl - 1)})
        lvl.append(_el('w:start', **{'w:val': '1'}))
        lvl.append(_el('w:numFmt', **{'w:val': 'decimal'}))
        lvl.append(_el('w:pStyle', **{'w:val': 'Heading%d' % ilvl}))
        lvl.append(_el('w:suff', **{'w:val': 'space'}))
        txt = '.'.join('%%%d' % i for i in range(1, ilvl + 1))
        lvl.append(_el('w:lvlText', **{'w:val': txt}))
        lvl.append(_el('w:lvlJc', **{'w:val': 'left'}))
        pPr_lvl = OxmlElement('w:pPr')
        pPr_lvl.append(_el('w:ind', **{'w:left': '0', 'w:firstLine': '0'}))
        lvl.append(pPr_lvl)
        abs_num.append(lvl)
    first = numbering.find(qn('w:num'))
    if first is not None:
        numbering.insert(0, abs_num)
    else:
        numbering.append(abs_num)
    num = _el('w:num', **{'w:numId': str(num_id)})
    num.append(_el('w:abstractNumId', **{'w:val': str(abstract_id)}))
    numbering.append(num)


def aplicar_numeracion(par, num_id=90, ilvl=0):
    pPr = par._p.get_or_add_pPr()
    _quitar(pPr, 'w:numPr')
    numPr = OxmlElement('w:numPr')
    numPr.append(_el('w:ilvl', **{'w:val': str(ilvl)}))
    numPr.append(_el('w:numId', **{'w:val': str(num_id)}))
    _ins(pPr, numPr, ORDEN_PPR)


# ----------------------------------------------------------------------------
# Campos (PAGE, NUMPAGES, TOC)
# ----------------------------------------------------------------------------
def add_field(par, instr, placeholder="1", bold=False, size=9, color=None, font=FUENTE):
    run = par.add_run()
    fld = _el('w:fldChar', **{'w:fldCharType': 'begin'})
    run._r.append(fld)
    run2 = par.add_run()
    it = OxmlElement('w:instrText')
    it.set(qn('xml:space'), 'preserve')
    it.text = ' %s ' % instr
    run2._r.append(it)
    run3 = par.add_run()
    run3._r.append(_el('w:fldChar', **{'w:fldCharType': 'separate'}))
    r4 = par.add_run(placeholder)
    run5 = par.add_run()
    run5._r.append(_el('w:fldChar', **{'w:fldCharType': 'end'}))
    for r in (run, run2, run3, r4, run5):
        r.font.name = font
        r.font.size = Pt(size)
        r.bold = bold
        if color is not None:
            r.font.color.rgb = color
    return r4


def add_tabla_contenido(doc, niveles='1-3'):
    par = doc.add_paragraph()
    run = par.add_run()
    run._r.append(_el('w:fldChar', **{'w:fldCharType': 'begin'}))
    run2 = par.add_run()
    it = OxmlElement('w:instrText')
    it.set(qn('xml:space'), 'preserve')
    it.text = ' TOC \\o "%s" \\h \\z \\u ' % niveles
    run2._r.append(it)
    run3 = par.add_run()
    run3._r.append(_el('w:fldChar', **{'w:fldCharType': 'separate'}))
    run4 = par.add_run(
        "Tabla de contenido automática. En Word: clic derecho sobre este texto → "
        "«Actualizar campo» → «Actualizar toda la tabla» (o Ctrl+E, F9) para "
        "generar los títulos y sus números de página."
    )
    run4.italic = True
    run4.font.size = Pt(9)
    run4.font.color.rgb = GRIS_CLARO
    run5 = par.add_run()
    run5._r.append(_el('w:fldChar', **{'w:fldCharType': 'end'}))


def marcar_entrada_indice(par, nivel=1, texto=None):
    """Agrega un campo TC para permitir índices especiales."""
    run = par.add_run()
    run._r.append(_el('w:fldChar', **{'w:fldCharType': 'begin'}))
    r2 = par.add_run()
    it = OxmlElement('w:instrText')
    it.set(qn('xml:space'), 'preserve')
    it.text = ' TC "  %s  " \\l "%d" ' % (texto or par.text, nivel)
    r2._r.append(it)
    r3 = par.add_run()
    r3._r.append(_el('w:fldChar', **{'w:fldCharType': 'end'}))


# ----------------------------------------------------------------------------
# Estilos
# ----------------------------------------------------------------------------
def _set_font(style, name=FUENTE, size=12, bold=False, color=None, italic=False,
              all_caps=False, spacing_pt=None):
    style.font.name = name
    style.font.size = Pt(size)
    style.font.bold = bold
    style.font.italic = italic
    if color is not None:
        style.font.color.rgb = color
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    for attr in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
        rFonts.set(qn(attr), name)
    if all_caps:
        style.font.all_caps = True
    if spacing_pt is not None:
        sp = OxmlElement('w:spacing')
        sp.set(qn('w:val'), str(int(spacing_pt * 20)))
        _ins(rPr, sp, ORDEN_RPR)


def construir_documento():
    doc = Document()

    # --- página: Carta, márgenes 3/2/2/2 ---
    sec = doc.sections[0]
    sec.page_width = Cm(21.59)
    sec.page_height = Cm(27.94)
    sec.top_margin = Cm(2.6)
    sec.bottom_margin = Cm(2.4)
    sec.left_margin = Cm(3.0)
    sec.right_margin = Cm(2.5)
    sec.header_distance = Cm(1.2)
    sec.footer_distance = Cm(1.1)

    styles = doc.styles

    # Normal
    normal = styles['Normal']
    _set_font(normal, size=11.5, color=NEGRO)
    pf = normal.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.space_before = Pt(0)
    pf.space_after = Pt(6)
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = 1.25
    pf.widow_control = True

    # Títulos 1-4 (numerados)
    cfg = {
        1: dict(size=15, color=AZUL, antes=20, despues=8, caps=True),
        2: dict(size=12.5, color=AZUL, antes=14, despues=5, caps=False),
        3: dict(size=11.5, color=AZUL_MED if False else RGBColor(0x2E, 0x4A, 0x6B), antes=10, despues=4, caps=False),
        4: dict(size=11.5, color=NEGRO, antes=8, despues=3, caps=False),
    }
    for lvl, c in cfg.items():
        st = styles['Heading %d' % lvl]
        _set_font(st, size=c['size'], bold=True, color=c['color'],
                  italic=(lvl == 4), all_caps=c['caps'])
        p = st.paragraph_format
        p.space_before = Pt(c['antes'])
        p.space_after = Pt(c['despues'])
        p.keep_with_next = True
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.line_spacing_rule = WD_LINE_SPACING.SINGLE
        p.line_spacing = 1.0
        p.left_indent = Cm(0)
        p.first_line_indent = Cm(0)

    # Título de documento / sección especial
    for nombre, size, color, caps, align in (
            ('TituloDoc', 20, AZUL, True, WD_ALIGN_PARAGRAPH.CENTER),
            ('SubtituloDoc', 13, GRIS, False, WD_ALIGN_PARAGRAPH.CENTER)):
        st = styles.add_style(nombre, 1)  # paragraph
        st.base_style = styles['Normal']
        _set_font(st, size=size, bold=(nombre == 'TituloDoc'), color=color, all_caps=caps)
        st.paragraph_format.alignment = align
        st.paragraph_format.space_before = Pt(6)
        st.paragraph_format.space_after = Pt(6)
        st.paragraph_format.line_spacing = 1.0
        st.paragraph_format.keep_with_next = True

    # Citas largas (sangría)
    cita = styles.add_style('CitaLegal', 1)
    cita.base_style = styles['Normal']
    _set_font(cita, size=10.5, color=GRIS)
    cita.paragraph_format.left_indent = Cm(1.2)
    cita.paragraph_format.right_indent = Cm(0.8)
    cita.paragraph_format.space_before = Pt(4)
    cita.paragraph_format.space_after = Pt(8)
    cita.paragraph_format.line_spacing = 1.1

    # Pie de tabla / fuente
    fuente = styles.add_style('Fuente', 1)
    fuente.base_style = styles['Normal']
    _set_font(fuente, size=8.5, italic=True, color=GRIS_CLARO)
    fuente.paragraph_format.space_before = Pt(2)
    fuente.paragraph_format.space_after = Pt(10)
    fuente.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    fuente.paragraph_format.line_spacing = 1.0

    # Texto de tabla
    tt = styles.add_style('Texto_Tabla', 1)
    tt.base_style = styles['Normal']
    _set_font(tt, size=9, color=NEGRO)
    tt.paragraph_format.space_before = Pt(1)
    tt.paragraph_format.space_after = Pt(1)
    tt.paragraph_format.line_spacing = 1.0
    tt.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT

    ttc = styles.add_style('Texto_Tabla_C', 1)
    ttc.base_style = tt
    ttc.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Viñetas
    for nm in ('List Bullet', 'List Bullet 2', 'List Number'):
        try:
            st = styles[nm]
            st.font.name = FUENTE
            st.font.size = Pt(11.5)
            st.paragraph_format.space_after = Pt(4)
            st.paragraph_format.line_spacing = 1.2
        except KeyError:
            pass

    _crear_numeracion(doc)
    return doc


# ----------------------------------------------------------------------------
# Encabezado y pie
# ----------------------------------------------------------------------------
def encabezado_pie(doc, texto_izq, texto_der, autor_pie):
    sec = doc.sections[0]
    sec.different_first_page_header_footer = True

    # Encabezado (a partir de la página 2)
    hp = sec.header.paragraphs[0]
    hp.text = ''
    hp.paragraph_format.tab_stops.add_tab_stop(Cm(16.0), WD_TAB_ALIGNMENT.RIGHT)
    hp.paragraph_format.space_after = Pt(2)
    r = hp.add_run(texto_izq)
    r.font.name = FUENTE_SANS
    r.font.size = Pt(7.5)
    r.font.color.rgb = GRIS_CLARO
    r.font.all_caps = True
    r2 = hp.add_run('\t' + texto_der)
    r2.font.name = FUENTE_SANS
    r2.font.size = Pt(7.5)
    r2.font.color.rgb = GRIS_CLARO
    # línea inferior del encabezado
    pPr = hp._p.get_or_add_pPr()
    pbdr = OxmlElement('w:pBdr')
    pbdr.append(_el('w:bottom', **{'w:val': 'single', 'w:sz': '4', 'w:space': '2', 'w:color': 'C3CAD4'}))
    _ins(pPr, pbdr, ORDEN_PPR)

    # Pie
    fp = sec.footer.paragraphs[0]
    fp.text = ''
    fp.paragraph_format.tab_stops.add_tab_stop(Cm(8.0), WD_TAB_ALIGNMENT.CENTER)
    fp.paragraph_format.tab_stops.add_tab_stop(Cm(16.0), WD_TAB_ALIGNMENT.RIGHT)
    r = fp.add_run(autor_pie)
    r.font.name = FUENTE_SANS
    r.font.size = Pt(7.5)
    r.font.color.rgb = GRIS_CLARO
    fp.add_run('\t')
    add_field(fp, 'PAGE', '1', size=8, color=GRIS)
    r = fp.add_run(' / ')
    r.font.size = Pt(8)
    r.font.name = FUENTE_SANS
    r.font.color.rgb = GRIS_CLARO
    add_field(fp, 'NUMPAGES', '1', size=8, color=GRIS)
    fp.add_run('\t')
    r = fp.add_run('Documento de trabajo académico')
    r.font.name = FUENTE_SANS
    r.font.size = Pt(7.5)
    r.font.color.rgb = GRIS_CLARO

    # Primera página: sin encabezado, pie simple
    fpf = sec.first_page_footer.paragraphs[0]
    fpf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = fpf.add_run(autor_pie)
    r.font.name = FUENTE_SANS
    r.font.size = Pt(7.5)
    r.font.color.rgb = GRIS_CLARO


# ----------------------------------------------------------------------------
# Componentes de contenido
# ----------------------------------------------------------------------------
def p(doc, texto='', style=None, size=None, bold=False, italic=False, color=None,
      align=None, space_before=None, space_after=None, indent_first=None,
      keep_with_next=False, line=None, style_name=None):
    par = doc.add_paragraph(style=style_name or style)
    if texto:
        r = par.add_run(texto)
        if size:
            r.font.size = Pt(size)
        r.bold = bold
        r.italic = italic
        if color is not None:
            r.font.color.rgb = color
    if align is not None:
        par.alignment = align
    if space_before is not None:
        par.paragraph_format.space_before = Pt(space_before)
    if space_after is not None:
        par.paragraph_format.space_after = Pt(space_after)
    if indent_first is not None:
        par.paragraph_format.first_line_indent = Cm(indent_first)
    if keep_with_next:
        par.paragraph_format.keep_with_next = True
    if line is not None:
        par.paragraph_format.line_spacing = line
    return par


def rich(doc, partes, style=None, align=None, space_after=None, space_before=None,
         indent_first=None, size=None, style_name=None):
    """partes: lista de (texto, dict_format); el dict admite b, i, c (tamaño), col y caps."""
    par = doc.add_paragraph(style=style_name or style)
    for item in partes:
        if isinstance(item, str):
            txt, fmt = item, {}
        else:
            txt, fmt = item
        r = par.add_run(txt)
        r.bold = fmt.get('b', False)
        r.italic = fmt.get('i', False)
        if fmt.get('c'):
            r.font.size = Pt(fmt['c'])
        elif size:
            r.font.size = Pt(size)
        if fmt.get('col') is not None:
            r.font.color.rgb = fmt['col']
        if fmt.get('caps'):
            r.font.all_caps = True
        if fmt.get('u'):
            r.underline = True
        if fmt.get('font'):
            r.font.name = fmt['font']
    if align is not None:
        par.alignment = align
    if space_after is not None:
        par.paragraph_format.space_after = Pt(space_after)
    if space_before is not None:
        par.paragraph_format.space_before = Pt(space_before)
    if indent_first is not None:
        par.paragraph_format.first_line_indent = Cm(indent_first)
    return par


def h(doc, nivel, texto):
    par = doc.add_heading(texto, level=nivel)
    aplicar_numeracion(par, num_id=90, ilvl=nivel - 1)
    return par


def h_sin_num(doc, nivel, texto):
    return doc.add_heading(texto, level=nivel)


def bullet(doc, texto, nivel=1, size=None):
    estilo = 'List Bullet' if nivel == 1 else 'List Bullet 2'
    par = doc.add_paragraph(style=estilo)
    if isinstance(texto, str):
        r = par.add_run(texto)
        if size:
            r.font.size = Pt(size)
    else:
        for t in texto:
            if isinstance(t, str):
                txt, fmt = t, {}
            else:
                txt, fmt = t
            r = par.add_run(txt)
            r.bold = fmt.get('b', False)
            r.italic = fmt.get('i', False)
            if fmt.get('c') or size:
                r.font.size = Pt(fmt.get('c') or size)
            if fmt.get('col') is not None:
                r.font.color.rgb = fmt['col']
    par.paragraph_format.space_after = Pt(4)
    return par


def cita(doc, texto, autor=None):
    par = doc.add_paragraph(style='CitaLegal')
    r = par.add_run(texto)
    if autor:
        par2 = doc.add_paragraph(style='CitaLegal')
        par2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r2 = par2.add_run(autor)
        r2.italic = True
        r2.font.size = Pt(10)
        par2.paragraph_format.space_after = Pt(10)
    return par


def tabla(doc, encabezados, filas, anchos, titulo=None, fuente=None,
          size=9, header_bg=HEX_AZUL, zebra=True, align_centro=None,
          primera_col_bold=False, header_size=8.5, alto=None):
    if titulo:
        tp = doc.add_paragraph()
        tp.paragraph_format.space_before = Pt(10)
        tp.paragraph_format.space_after = Pt(3)
        tp.paragraph_format.keep_with_next = True
        tp.paragraph_format.line_spacing = 1.0
        rr = tp.add_run(titulo)
        rr.bold = True
        rr.font.size = Pt(9)
        rr.font.name = FUENTE_SANS
        rr.font.color.rgb = AZUL
        rr.font.all_caps = True

    # protección: ningún conjunto de anchos puede exceder el área útil (16,0 cm)
    _tot = sum(anchos)
    if _tot > 16.0:
        anchos = [round(a * 16.0 / _tot, 2) for a in anchos]
    t = doc.add_table(rows=1, cols=len(encabezados))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    bordes_tabla(t)
    hdr = t.rows[0]
    repetir_encabezado(hdr)
    no_dividir_fila(hdr)
    for i, txt in enumerate(encabezados):
        c = hdr.cells[i]
        c.text = ''
        sombrear(c, header_bg)
        margenes_celda(c)
        par = c.paragraphs[0]
        par.style = doc.styles['Texto_Tabla']
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = par.add_run(txt)
        r.bold = True
        r.font.size = Pt(header_size)
        r.font.color.rgb = BLANCO
        r.font.name = FUENTE_SANS
        r.font.all_caps = True

    for fi, fila in enumerate(filas):
        row = t.add_row()
        no_dividir_fila(row)
        for ci, val in enumerate(fila):
            c = row.cells[ci]
            c.text = ''
            margenes_celda(c)
            if zebra and fi % 2 == 1:
                sombrear(c, HEX_CELESTE2)
            par = c.paragraphs[0]
            par.style = doc.styles['Texto_Tabla']
            if align_centro and ci in align_centro:
                par.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            if isinstance(val, list):
                partes = val
            elif isinstance(val, tuple):
                partes = [val]
            else:
                partes = [(str(val), {})]
            for item in partes:
                if isinstance(item, str):
                    txt, fmt = item, {}
                else:
                    txt, fmt = item
                r = par.add_run(txt)
                r.bold = fmt.get('b', False) or (primera_col_bold and ci == 0)
                r.italic = fmt.get('i', False)
                r.font.size = Pt(fmt.get('c', size))
                if fmt.get('col') is not None:
                    r.font.color.rgb = fmt['col']
    set_col_widths(t, anchos)

    if fuente:
        fp = doc.add_paragraph(style='Fuente')
        fp.add_run(fuente)
    return t


RECUADROS = os.environ.get("SIN_RECUADROS", "1") == "1"


def _trozos(linea):
    """Normaliza una línea de caja (str / tupla (txt, fmt) / lista de ambas) en trozos."""
    if isinstance(linea, str):
        return [(linea, {})]
    if isinstance(linea, tuple):
        return [(linea[0], linea[1])]
    return [(it, {}) if isinstance(it, str) else (it[0], it[1]) for it in linea]


def _prosa_antetitulo(doc, titulo, lineas):
    """Vuelca una caja como texto corrido: entrada negrita + párrafos justificados."""
    primero = True
    for ln in lineas:
        par = doc.add_paragraph()
        par.style = doc.styles['Normal']
        par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        par.paragraph_format.space_after = Pt(6)
        if primero:
            r = par.add_run(titulo.rstrip('.') + ". ")
            r.bold = True
            r.font.size = Pt(11)
            primero = False
        for txt, fmt in _trozos(ln):
            rr = par.add_run(txt)
            rr.bold = fmt.get('b', False)
            rr.italic = fmt.get('i', False)
            rr.font.size = Pt(11)
    return None


def caja(doc, titulo, lineas, tipo='info', ancho=16.0):
    """Recuadro de énfasis. Con SIN_RECUADROS=1, el contenido se vierte como prosa:

    el antetítulo se convierte en entrada negrita y el resto en párrafos normales, sin
    color ni borde. El razonamiento no se pierde; solo deja de ir enmarcado.
    """
    if RECUADROS:
        return _prosa_antetitulo(doc, titulo, lineas)
    paleta = {
        'info': (HEX_AZUL, HEX_CELESTE),
        'alerta': (HEX_VINO, 'FBEDEA'),
        'exito': (HEX_VERDE, 'E9F2ED'),
        'nota': (HEX_AMBAR, 'FBF6E6'),
    }
    barra, fondo = paleta.get(tipo, paleta['info'])
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    c = t.rows[0].cells[0]
    c.text = ''
    sombrear(c, fondo)
    bordes_celda(c, color=barra, sz=4)
    # barra izquierda gruesa
    tcB = c._tc.get_or_add_tcPr().find(qn('w:tcBorders'))
    if tcB is not None:
        izq = tcB.find(qn('w:left'))
        if izq is not None:
            izq.set(qn('w:sz'), '24')
            izq.set(qn('w:color'), barra)
    margenes_celda(c, top=120, bottom=130, left=200, right=180)
    par = c.paragraphs[0]
    par.style = doc.styles['Texto_Tabla']
    par.paragraph_format.space_after = Pt(4)
    r = par.add_run(titulo)
    r.bold = True
    r.font.size = Pt(9)
    r.font.name = FUENTE_SANS
    r.font.all_caps = True
    r.font.color.rgb = RGBColor.from_string(barra)
    for i, ln in enumerate(lineas):
        pp = c.add_paragraph()
        pp.style = doc.styles['Texto_Tabla']
        pp.paragraph_format.space_after = Pt(3)
        pp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        if isinstance(ln, tuple):
            # una sola corrida con formato: (texto, {"b": True})
            txt, fmt = ln
            rr = pp.add_run(txt)
            rr.bold = fmt.get('b', False)
            rr.italic = fmt.get('i', False)
            rr.font.size = Pt(fmt.get('c', 9))
            if fmt.get('col') is not None:
                rr.font.color.rgb = fmt['col']
        elif isinstance(ln, list):
            for item in ln:
                txt, fmt = (item, {}) if isinstance(item, str) else item
                rr = pp.add_run(txt)
                rr.bold = fmt.get('b', False)
                rr.italic = fmt.get('i', False)
                rr.font.size = Pt(fmt.get('c', 9))
                if fmt.get('col') is not None:
                    rr.font.color.rgb = fmt['col']
        else:
            rr = pp.add_run(ln)
            rr.font.size = Pt(9)
    set_col_widths(t, [ancho])
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


def salto_pagina(doc):
    doc.add_page_break()


def linea_horizontal(doc, color=HEX_AZUL, sz=8, space_before=2, space_after=8):
    par = doc.add_paragraph()
    par.paragraph_format.space_before = Pt(space_before)
    par.paragraph_format.space_after = Pt(space_after)
    pPr = par._p.get_or_add_pPr()
    pbdr = OxmlElement('w:pBdr')
    pbdr.append(_el('w:bottom', **{'w:val': 'single', 'w:sz': str(sz),
                                    'w:space': '1', 'w:color': color}))
    _ins(pPr, pbdr, ORDEN_PPR)
    return par


def imagen(doc, ruta, ancho_cm, pie=None):
    par = doc.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    par.paragraph_format.space_before = Pt(8)
    par.paragraph_format.space_after = Pt(2)
    par.paragraph_format.keep_with_next = True
    par.add_run().add_picture(ruta, width=Cm(ancho_cm))
    if pie:
        fp = doc.add_paragraph(style='Fuente')
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fp.add_run(pie)


# ----------------------------------------------------------------------------
# Normalización final del orden de elementos (compatibilidad estricta OOXML)
# ----------------------------------------------------------------------------
_ORDENES = {
    'tcPr': ORDEN_TCPR,
    'trPr': ORDEN_TRPR,
    'tblPr': ORDEN_TBLPR,
    'pPr': ORDEN_PPR,
    'rPr': ORDEN_RPR,
    'lvl': ['w:start', 'w:numFmt', 'w:lvlRestart', 'w:pStyle', 'w:isLgl', 'w:suff',
            'w:lvlText', 'w:lvlPicBulletId', 'w:legacy', 'w:lvlJc', 'w:pPr', 'w:rPr'],
    'abstractNum': ['w:nsid', 'w:multiLevelType', 'w:tmpl', 'w:name', 'w:styleLink',
                    'w:numStyleLink', 'w:lvl'],
    'num': ['w:abstractNumId', 'w:lvlOverride'],
}


def normalizar_orden(doc):
    """Reordena los hijos de las propiedades según el esquema (evita avisos de reparación)."""
    partes = [doc.element]
    for sec in doc.sections:
        for ref in (sec.header, sec.footer, sec.first_page_header, sec.first_page_footer,
                    sec.even_page_header, sec.even_page_footer):
            try:
                partes.append(ref._element)
            except Exception:
                pass
    try:
        partes.append(doc.part.numbering_part.element)
    except Exception:
        pass
    try:
        partes.append(doc.styles.element)
    except Exception:
        pass

    for raiz in partes:
        for tag, orden in _ORDENES.items():
            for el in raiz.iter(qn('w:' + tag)):
                hijos = list(el)
                if len(hijos) < 2:
                    continue
                orden_q = [qn(t) for t in orden]

                def clave(h):
                    return orden_q.index(h.tag) if h.tag in orden_q else len(orden_q)
                ordenados = sorted(hijos, key=clave)
                if [id(x) for x in ordenados] != [id(x) for x in hijos]:
                    for h in hijos:
                        el.remove(h)
                    for h in ordenados:
                        el.append(h)
