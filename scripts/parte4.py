# -*- coding: utf-8 -*-
"""EJE IV: Insumo procesal y determinación de competencia judicial."""
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx_engine import (p, rich, h, bullet, cita, tabla, caja, salto_pagina,
                         aplicar_numeracion, AZUL, VINO, GRIS, GRIS_CLARO, NEGRO,
                         HEX_AZUL, HEX_CELESTE, sombrear, bordes_celda, set_col_widths,
                         margenes_celda, linea_horizontal, FUENTE_SANS)
from datos import SMLMV, money, IBC_TOTAL, TOPE_CUANTIA

AUTOR = "Julia Inés Buitrago García"
NIT = "39626936-8"


# ---------------------------------------------------------------- utilidades
def rotulo(doc, texto):
    par = p(doc, texto, size=9.5, bold=True, color=AZUL, space_before=10, space_after=4,
            keep_with_next=True)
    par.runs[0].font.name = FUENTE_SANS
    par.runs[0].font.all_caps = True
    return par


def parrafo_doc(doc, texto, sangria=0, just=True, size=10.5, bold=False, space_after=7,
                align=None, italic=False):
    par = doc.add_paragraph()
    r = par.add_run(texto)
    r.font.size = Pt(size)
    r.bold = bold
    r.italic = italic
    par.alignment = align if align is not None else (
        WD_ALIGN_PARAGRAPH.JUSTIFY if just else WD_ALIGN_PARAGRAPH.LEFT)
    par.paragraph_format.left_indent = Cm(sangria)
    par.paragraph_format.space_after = Pt(space_after)
    par.paragraph_format.line_spacing = 1.15
    return par


def numerado_doc(doc, num, texto, size=10.5):
    par = doc.add_paragraph()
    par.paragraph_format.left_indent = Cm(1.1)
    par.paragraph_format.first_line_indent = Cm(-1.1)
    par.paragraph_format.space_after = Pt(6)
    par.paragraph_format.line_spacing = 1.15
    par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = par.add_run("%s. " % num)
    r.bold = True
    r.font.size = Pt(size)
    r2 = par.add_run(texto)
    r2.font.size = Pt(size)
    return par


def encabezado_escrito(doc, entidad, ciudad_fecha):
    t = doc.add_table(rows=1, cols=1)
    c = t.rows[0].cells[0]
    sombrear(c, HEX_CELESTE)
    bordes_celda(c, color="B7C0CC")
    margenes_celda(c, top=110, bottom=110, left=160, right=160)
    c.text = ""
    par = c.paragraphs[0]
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = par.add_run("DOCUMENTO MODELO PARA RADICACIÓN")
    r.bold = True; r.font.size = Pt(8); r.font.name = FUENTE_SANS; r.font.color.rgb = AZUL
    r.font.all_caps = True
    par2 = c.add_paragraph(); par2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = par2.add_run(entidad)
    r.bold = True; r.font.size = Pt(10.5); r.font.name = FUENTE_SANS; r.font.color.rgb = AZUL
    par3 = c.add_paragraph(); par3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    par3.paragraph_format.space_after = Pt(0)
    r = par3.add_run(ciudad_fecha)
    r.font.size = Pt(9); r.font.name = FUENTE_SANS; r.font.color.rgb = GRIS
    set_col_widths(t, [16.0])
    p(doc, "", space_after=2)


def firma(doc):
    p(doc, "", space_after=18)
    linea_horizontal(doc, color="9AA4B1", sz=4, space_after=2)
    par = p(doc, AUTOR, size=10.5, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, space_after=0)
    p(doc, "C.C. ____________________  ·  Representante legal de EXTINTORES S C M  ·  NIT %s" % NIT,
      size=9, color=GRIS, space_after=2)
    p(doc, "Correo electrónico de notificaciones: ____________________  ·  Teléfono: ____________________",
      size=9, color=GRIS, space_after=2)


# ================================================================ EJE IV
def eje_iv(doc):
    salto_pagina(doc)
    h(doc, 1, "EJE IV · RECLAMACIÓN ADMINISTRATIVA Y COMPETENCIA JUDICIAL")

    # ---------- 4.1
    h(doc, 2, "Calificación del mecanismo jurídico aplicable")
    rich(doc, ["La guía del proyecto integrador distingue, acertadamente, entre la actuación que "
               "corresponde a una entidad pública y la que corresponde a una empresa privada. La "
               "distinción no es formal sino de fondo: la reclamación administrativa previa del "
               "artículo 6 del Código Procesal del Trabajo y de la Seguridad Social es un requisito "
               "de procedibilidad dirigido a las entidades de derecho público, y su incumplimiento "
               "(según la jurisprudencia) es insubsanable. Aplicarla a un particular implicaría "
               "trasladar al empleador una carga procesal que el legislador reservó para la "
               "Administración, con el riesgo adicional de que el escrito no produzca efecto "
               "interruptivo alguno."])
    rich(doc, ["La estructura del artículo 6 del C.P.T.S.S. lo confirma: la norma se refiere a "
               "«las acciones contenciosas contra la Nación, las entidades territoriales y "
               "cualquiera otra entidad de la administración pública», y la Corte Constitucional, "
               "al examinarla en la Sentencia C-792 de 2006, la entendió como una modalidad de "
               "«aseguramiento de la oportunidad para la autotutela administrativa»: se otorga a la "
               "Administración la posibilidad de revisar sus propias actuaciones antes de que sean "
               "sometidas a la jurisdicción. Ese presupuesto no existe frente a un empleador privado."])
    tabla(doc, ["Naturaleza del demandado", "Mecanismo exigible", "Norma aplicable", "Efecto del escrito"],
          [["Entidad pública", "Reclamación administrativa previa", "Art. 6, C.P.T.S.S.; C-792 de 2006",
            "Requisito de procedibilidad; suspende la prescripción"],
           ["Empresa privada (EXTINTORES S C M)", "Derecho de petición y solicitud de agotamiento de "
            "la vía interna", "Art. 23 C.P. y Ley 1755 de 2015", "Interrumpe la prescripción y "
            "constituye prueba de la reclamación"],
           ["Entidad del SGSSS (EPS, ARL, AFP)", "Derecho de petición y reclamación directa",
            "Art. 23 C.P.; Ley 1755 de 2015; arts. 22 y 24, Ley 100 de 1993",
            "Elemento necesario para acreditar la reclamación ante el juez (art. 11, C.P.T.S.S.)"],
           ["Cualquiera de las anteriores", "Conciliación extrajudicial, cuando el asunto sea "
            "conciliable", "Ley 640 de 2001", "Requisito de procedibilidad en los asuntos "
            "susceptibles de conciliación"]],
          [3.6, 4.4, 4.2, 3.8], size=8.4,
          titulo="Tabla 23. Determinación del mecanismo según la naturaleza del sujeto")

    rich(doc, ["Para EXTINTORES S C M la consecuencia práctica es que la reclamación se dirige al "
               "empleador mediante derecho de petición y solicitud de agotamiento de la vía interna, "
               "y que las decisiones de las entidades del sistema (EPS, ARL y AFP) se controvierten "
               "mediante derecho de petición y, agotada la respuesta, mediante demanda ante la "
               "jurisdicción laboral. Es importante precisar que la conciliación extrajudicial solo "
               "procede respecto de derechos inciertos y discutibles; los derechos ciertos e "
               "indiscutibles, como las semanas efectivamente cotizadas o las prestaciones ya "
               "reconocidas, no son conciliables."])
    caja(doc, "Error frecuente que debe evitarse",
         ["Radicar una «reclamación administrativa previa» ante una empresa privada no suspende la "
          "prescripción y puede llevar a que el juez declare la demanda mal formulada. Inversamente, "
          "demandar directamente a una entidad pública sin reclamación previa conduce al rechazo de "
          "la demanda por falta de requisito de procedibilidad. La calificación del sujeto pasivo es, "
          "por tanto, la primera decisión procesal del caso."],
         tipo='alerta')

    # ---------- 4.2 documentos
    h(doc, 2, "Insumo procesal 1: derecho de petición ante la ARL")
    rich(doc, ["Se redacta el escrito dirigido a la administradora de riesgos laborales con dos "
               "pretensiones conexas: la revisión de la clasificación del riesgo y el reconocimiento "
               "de las prestaciones derivadas del accidente AT-2026-014. El modelo incluye los "
               "elementos que la Ley 1755 de 2015 exige (designación de la autoridad, identificación del peticionario, objeto, razones y petición concreta) y añade la solicitud de copias "
               "y la advertencia sobre el silencio administrativo."])
    salto_pagina(doc)
    encabezado_escrito(doc, "ADMINISTRADORA DE RIESGOS LABORALES [NOMBRE DE LA ARL]  ·  "
                            "OFICINA DE ATENCIÓN AL USUARIO", "Soacha (Cundinamarca), noviembre de 2026")
    rotulo(doc, "Referencia")
    parrafo_doc(doc, "Asunto: derecho de petición de interés particular. Solicitud de revisión de la "
                     "clasificación del riesgo, reconocimiento de prestaciones derivadas del "
                     "accidente de trabajo AT-2026-014 y expedición de copias.", bold=True, size=10.5)
    rotulo(doc, "Identificación del peticionario")
    parrafo_doc(doc, "%s, mayor de edad, identificada con cédula de ciudadanía número __________ "
                     "expedida en Soacha, en calidad de representante legal y titular de la empresa "
                     "EXTINTORES S C M, identificada con cédula/NIT %s, con domicilio en el municipio "
                     "de Soacha (Cundinamarca), carrera 4 n.° 60-17, barrio La Despensa, "
                     "dirección __________, correo electrónico __________, teléfono __________, "
                     "actuando en nombre propio y en ejercicio del derecho fundamental de petición "
                     "consagrado en el artículo 23 de la Constitución Política y regulado por la Ley "
                     "1755 de 2015, presento la siguiente solicitud." % (AUTOR, NIT))

    rotulo(doc, "Hechos")
    hechos = [
        "La empresa EXTINTORES S C M se encuentra afiliada a esa administradora desde el __ de "
        "__________ de ____, con el número de afiliación __________, y tiene a su cargo tres (3) "
        "trabajadores dependientes, todos afiliados al Sistema General de Riesgos Laborales.",
        "La actividad económica principal registrada en el Registro Único Tributario corresponde al "
        "código CIIU 4759 (comercio al por menor de otros artículos domésticos en establecimientos "
        "especializados), que la tabla del Decreto 768 de 2022 clasifica en clase de riesgo I, "
        "subactividad 4759-01, con una tarifa inicial del 0,522 % del ingreso base de cotización. Ese "
        "código coincide con el reportado a la Secretaría de Salud de Soacha en el acta "
        "028-050164-26 del 09/09/2026.",
        "A pesar de esa clasificación, en las planillas de liquidación de aportes de los últimos ____ "
        "periodos la empresa ha cotizado con una tarifa equivalente a la clase de riesgo II, sin que "
        "obre decisión motivada de esa administradora que justifique la clasificación aplicada.",
        "El 12 de agosto de 2026 el trabajador Luis Antonio Aguilar Rodríguez, C.C. 1073694256, sufrió "
        "un accidente de trabajo durante la "
        "prueba hidrostática de un cilindro extintor, que le produjo fractura conminuta de los dedos "
        "índice y medio de la mano derecha y lesión del tendón flexor.",
        "El accidente fue reportado a esa administradora el 17 de agosto de 2026, es decir, cinco "
        "días después de su ocurrencia.",
        "A la fecha de esta solicitud han transcurrido más de ____ días desde el reporte sin que esa "
        "administradora haya notificado decisión de fondo sobre el reconocimiento de las prestaciones "
        "económicas ni sobre la calificación de la pérdida de la capacidad laboral del trabajador.",
    ]
    for i, t in enumerate(hechos, 1):
        numerado_doc(doc, i, t)

    rotulo(doc, "Fundamentos de derecho")
    for t in [
        "Artículos 23 y 29 de la Constitución Política: derecho de petición, debido proceso y "
        "respuesta de fondo, clara y precisa.",
        "Ley 1755 de 2015, artículos 5, 14 y 16: términos para resolver (quince días hábiles) y "
        "consecuencias de la falta de respuesta.",
        "Artículos 5, 7 y 13 del Decreto 1295 de 1994 y artículo 13 de la Ley 1562 de 2012: "
        "obligaciones de la administradora y del empleador en materia de prestaciones y reporte.",
        "Artículos 1, 5 y 7 de la Ley 776 de 2002: subsidio por incapacidad temporal equivalente al "
        "100 % del salario base de cotización e indemnización por incapacidad permanente parcial.",
        "Artículos 2, 3 y 4 del Decreto 768 de 2022 y artículo 13 del Decreto 1772 de 1994: "
        "clasificación de la actividad económica y tarifas de cotización por clase de riesgo.",
    ]:
        bullet(doc, t, size=10)

    rotulo(doc, "Peticiones")
    peticiones = [
        "Que se revise y, en su caso, se corrija la clase de riesgo asignada a EXTINTORES S C M, "
        "conforme al código CIIU 4759 y a la tabla del Decreto 768 de 2022, con efectos a partir del "
        "periodo en que se produjo la discrepancia.",
        "Que se reliquiden las planillas de aportes y se devuelvan o compensen las sumas pagadas en "
        "exceso, con los intereses a que haya lugar.",
        "Que se reconozcan y paguen las prestaciones asistenciales y económicas derivadas del "
        "accidente de trabajo AT-2026-014, incluido el subsidio por incapacidad temporal y, previa "
        "calificación, la indemnización por incapacidad permanente parcial.",
        "Que se practique la calificación del origen y del porcentaje de pérdida de la capacidad "
        "laboral del trabajador, con sujeción al manual único adoptado por el Decreto 1507 de 2014.",
        "Que se expida copia íntegra del expediente administrativo del evento y de la decisión de "
        "clasificación del riesgo.",
        "Que se informe el nombre y el cargo del funcionario responsable de resolver, y se notifique "
        "la respuesta al correo electrónico indicado.",
    ]
    for i, t in enumerate(peticiones, 1):
        numerado_doc(doc, i, t)

    rotulo(doc, "Pruebas y anexos")
    for t in [
        "Certificado de existencia y representación o documento de identidad del peticionario.",
        "Registro Único Tributario y certificado de cámara de comercio, si aplica.",
        "Copias de las planillas de liquidación de aportes de los periodos objeto de la solicitud.",
        "Copia del reporte del accidente de trabajo y de los certificados de incapacidad.",
        "Historia clínica ocupacional y concepto del médico laboral, cuando exista.",
    ]:
        bullet(doc, t, size=10)
    parrafo_doc(doc, "Solicito que las notificaciones se surtan en la dirección y el correo "
                     "electrónico indicados. La falta de respuesta dentro del término legal me "
                     "habilitará para acudir a la jurisdicción y para invocar las consecuencias "
                     "previstas en la Ley 1755 de 2015.", size=10.5)
    parrafo_doc(doc, "Atentamente,", size=10.5, space_after=2)
    firma(doc)

    # ---- documento 2
    salto_pagina(doc)
    h(doc, 2, "Insumo procesal 2: derecho de petición ante la AFP y solicitud de agotamiento de vía")
    rich(doc, ["El segundo escrito se dirige a la administradora de fondos de pensiones y persigue un "
               "objetivo probatorio: obtener el expediente de afiliación y el acta de doble asesoría "
               "del caso P-03, o la constancia de su inexistencia. Ese documento es el que permite "
               "decidir si se demanda la ineficacia del traslado."])
    p(doc, "", space_after=6)
    encabezado_escrito(doc, "ADMINISTRADORA DE FONDOS DE PENSIONES [NOMBRE DE LA AFP]  ·  "
                            "OFICINA DE SERVICIO AL CLIENTE", "Soacha (Cundinamarca), noviembre de 2026")
    rotulo(doc, "Referencia")
    parrafo_doc(doc, "Asunto: derecho de petición. Solicitud de copia del expediente de afiliación y "
                     "del acta de doble asesoría de Luis Antonio Aguilar Rodríguez, y "
                     "reclamación por déficit de información en el traslado de régimen pensional.",
                     bold=True, size=10.5)
    rotulo(doc, "Identificación")
    parrafo_doc(doc, "%s, identificada con cédula de ciudadanía número __________, en calidad de "
                     "representante legal de EXTINTORES S C M (NIT %s) y empleadora de Luis Antonio "
                     "Aguilar Rodríguez, identificado con cédula de ciudadanía 1073694256, presento "
                     "derecho de petición y formulo reclamación en los siguientes términos."
                     % (AUTOR, NIT))
    rotulo(doc, "Hechos")
    for i, t in enumerate([
        "Luis Antonio Aguilar Rodríguez se afilió a esa administradora en el año 1997, cuando contaba con 27 años "
        "de edad, y permanece vinculado como cotizante dependiente a través de EXTINTORES S C M.",
        "A la fecha cuenta con 1.150 semanas cotizadas y 56 años de edad, y aspira a acceder a la "
        "pensión de vejez en el año 2027.",
        "La afiliada manifiesta que la decisión de traslado se produjo tras una visita de un asesor "
        "comercial, quien le indicó que obtendría una mesada superior, sin que se le entregara "
        "simulación, proyección o documento comparativo de los regímenes.",
        "La afiliada no conserva copia del formulario de afiliación ni recuerda haber suscrito acta "
        "de doble asesoría.",
        "La empresa empleadora no es titular del deber de asesoría, que corresponde a la "
        "administradora conforme a los artículos 20, 48 y 53 de la Constitución Política y a la Ley "
        "1748 de 2014.",
    ], 1):
        numerado_doc(doc, i, t)
    rotulo(doc, "Fundamentos de derecho")
    for t in [
        "Artículo 23 de la Constitución Política y Ley 1755 de 2015.",
        "Artículos 271 y 272 de la Ley 100 de 1993: ineficacia del traslado por incumplimiento del "
        "deber de información.",
        "Ley 1748 de 2014, artículo 3: deber de doble asesoría como condición previa al traslado.",
        "Corte Constitucional, Sentencia SU-107 de 2024: estándar de información, consentimiento "
        "informado y equilibrio de la carga probatoria.",
        "Corte Suprema de Justicia, Sala de Casación Laboral, sentencias SL3179-2023 y SL1179-2024.",
    ]:
        bullet(doc, t, size=10)
    rotulo(doc, "Peticiones")
    for i, t in enumerate([
        "Que se expida copia íntegra del expediente de afiliación del solicitante, incluidos el "
        "formulario, las simulaciones, los extractos y las comunicaciones remitidas.",
        "Que se expida copia del acta de doble asesoría o, en su defecto, se certifique que no existe "
        "soporte documental de la asesoría prestada.",
        "Que se certifique el capital acumulado, los rendimientos, las comisiones y las primas "
        "descontadas desde la fecha de afiliación.",
        "Que se revise de oficio la validez del traslado y, si se constata el déficit de información, "
        "se adopten las medidas para restablecer la afiliación al Régimen de Prima Media.",
        "Que se responda de fondo dentro de los quince días hábiles siguientes.",
    ], 1):
        numerado_doc(doc, i, t)
    parrafo_doc(doc, "La presente solicitud constituye la reclamación que se tendrá como surtida para "
                     "efectos de la competencia territorial prevista en el artículo 11 del Código "
                     "Procesal del Trabajo y de la Seguridad Social.", size=10.5)
    parrafo_doc(doc, "Atentamente,", size=10.5, space_after=2)
    firma(doc)

    # ---- documento 3
    salto_pagina(doc)
    h(doc, 2, "Insumo procesal 3: solicitud de agotamiento de la vía interna ante el empleador")
    rich(doc, ["El tercer modelo se dirige al propio empleador y es el que la guía exige para las "
               "empresas privadas. Su función es doble: permite corregir el error antes de litigar y "
               "documentar la reclamación para efectos de la interrupción de la prescripción y de la "
               "prueba del agotamiento de la vía interna."])
    p(doc, "", space_after=6)
    encabezado_escrito(doc, "EXTINTORES S C M  ·  GERENCIA GENERAL", "Soacha (Cundinamarca), noviembre de 2026")
    rotulo(doc, "Referencia")
    parrafo_doc(doc, "Asunto: solicitud de agotamiento de la vía interna y reclamación por "
                     "diferencias en el reconocimiento de prestaciones económicas del Sistema de "
                     "Seguridad Social Integral.", bold=True, size=10.5)
    rotulo(doc, "Hechos")
    for i, t in enumerate([
        "La suscrita es titular y representante legal de la empresa EXTINTORES S C M, empleadora de "
        "diez (10) trabajadores dependientes.",
        "En la auditoría interna practicada en octubre de 2026 se identificaron diferencias en la "
        "liquidación de una licencia de paternidad y en el reporte de novedades de uno de los "
        "trabajadores.",
        "Las diferencias ascienden a %s y afectan el ingreso de Felipe Martínez (C.C. 8686834), por la "
        "licencia de paternidad, y las semanas reportadas de los tres trabajadores de la planta."
        % money(412000),
        "La empresa reconoce la procedencia de la corrección y requiere formalizar el trámite para "
        "que produzca efectos frente a la EPS y al fondo de pensiones.",
    ], 1):
        numerado_doc(doc, i, t)
    rotulo(doc, "Solicitud")
    for i, t in enumerate([
        "Que se reconozca y ordene el pago de la diferencia dejada de percibir por el trabajador "
        "Felipe Martínez, con los ajustes a que haya lugar.",
        "Que se corrija la planilla integrada de liquidación de aportes y se reporten las semanas "
        "omitidas en la historia laboral de los trabajadores de la planta.",
        "Que se expida constancia escrita de la respuesta, con indicación de la fecha y del "
        "responsable de la decisión.",
    ], 1):
        numerado_doc(doc, i, t)
    parrafo_doc(doc, "La presente solicitud se presenta dentro del término legal y se entiende como "
                     "agotamiento de la vía interna para todos los efectos. Su respuesta deberá "
                     "producirse dentro de los quince (15) días hábiles siguientes, conforme al "
                     "artículo 23 de la Constitución Política y a la Ley 1755 de 2015.", size=10.5)
    parrafo_doc(doc, "Atentamente,", size=10.5, space_after=2)
    firma(doc)

    # ---------- 4.3 competencia
    salto_pagina(doc)
    h(doc, 2, "Fijación de la competencia judicial")
    rich(doc, ["Determinar el juez competente exige aplicar, en orden, los factores que el Código "
               "Procesal del Trabajo y de la Seguridad Social establece: el factor objetivo o por "
               "razón de la materia, el factor territorial, el factor por razón de la cuantía y el "
               "factor funcional. Un error en cualquiera de ellos conduce a la remisión del expediente "
               "y a la pérdida de tiempo procesal."])

    h(doc, 3, "Competencia por razón de la materia")
    rich(doc, ["El artículo 2 del C.P.T.S.S. atribuye a la jurisdicción ordinaria, en su especialidad "
               "laboral, el conocimiento de «las controversias referentes al sistema de seguridad "
               "social integral que se susciten entre los afiliados, beneficiarios o usuarios, los "
               "empleadores y las entidades administradoras o prestadoras, cualquiera que sea la "
               "naturaleza de la relación jurídica y de los actos jurídicos que se controviertan». "
               "Así, tanto la demanda de ineficacia del traslado como la de reconocimiento de "
               "prestaciones de la ARL o la de reliquidación de aportes corresponden al juez laboral, "
               "y no al juez administrativo ni al civil."])

    h(doc, 3, "Competencia territorial (artículo 11 del C.P.T.S.S.)")
    rich(doc, ["El artículo 11 del C.P.T.S.S., modificado por el artículo 8 de la Ley 712 de 2001, "
               "dispone que en los procesos contra las entidades que conforman el sistema de "
               "seguridad social integral ",
               ("«será competente el juez laboral del circuito del lugar del domicilio de la entidad "
                "de seguridad social demandada o el del lugar donde se haya surtido la reclamación "
                "del respectivo derecho, a elección del demandante»", {"i": True}), ". La norma "
               "consagra, por tanto, un fuero concurrente con dos opciones y deja la elección en "
               "manos del demandante, lo que constituye una garantía de acceso a la administración de "
               "justicia. Para los procesos contra particulares rige la regla general del artículo 5 "
               "(lugar de prestación del servicio o domicilio del demandado) y, en los procesos "
               "contra la Nación, el artículo 7."])
    tabla(doc, ["Escenario", "Demandado", "Opciones del demandante", "Juez competente"],
          [["Ineficacia del traslado (caso P-03)", "AFP con domicilio en Bogotá D.C.",
            "Lugar de prestación del servicio (Soacha) o domicilio de la AFP (Bogotá D.C.)",
            "A elección del demandante: Juzgado Laboral del Circuito de Soacha o de Bogotá D.C."],
           ["Prestaciones del SGRL (AT-2026-014)", "ARL con domicilio en Bogotá D.C.",
            "Lugar de prestación del servicio (Soacha) o domicilio de la ARL (Bogotá D.C.)",
            "A elección del demandante: Juzgado Laboral del Circuito de Soacha o de Bogotá D.C."],
           ["Diferencias de aportes o prestaciones", "EXTINTORES S C M (Soacha)",
            "Lugar de prestación del servicio o domicilio del demandado (ambos en Soacha)",
            "Juzgado Laboral del Circuito de Soacha"]],
          [4.2, 3.6, 4.2, 4.0], size=8.4,
          titulo="Tabla 24. Aplicación del factor territorial")
    rich(doc, ["La reclamación cumple una función adicional relevante: el artículo 6 del C.P.T.S.S. "
               "establece que mientras esté pendiente el agotamiento de la reclamación administrativa "
               "se suspende el término de prescripción de la respectiva acción. Por esa razón el "
               "escrito debe radicarse con constancia (sello, correo certificado o radicado electrónico) y conservarse junto con la respuesta."])

    h(doc, 3, "Competencia por razón de la cuantía (artículo 12 del C.P.T.S.S.)")
    rich(doc, ["El artículo 12 del C.P.T.S.S., modificado por el artículo 46 de la Ley 1395 de 2010, "
               "distribuye la competencia así: los jueces laborales del circuito conocen en única "
               "instancia de los negocios cuya cuantía exceda el equivalente a veinte (20) salarios "
               "mínimos legales mensuales vigentes y en primera instancia de todos los demás; y los "
               "jueces municipales de pequeñas causas y competencia múltiple, donde existan, conocen "
               "en única instancia de los negocios cuya cuantía no exceda ese tope. Para 2026, el "
               "umbral es de %s." % money(TOPE_CUANTIA)])
    tabla(doc, ["Caso", "Pretensión principal", "Cuantía estimada", "Competencia"],
          [["P-03 · Ineficacia del traslado",
            "Diferencia de mesadas: %s mensuales × 12 mesadas × 20 años" % money(345095),
            money(345095 * 12 * 20),
            "Juez Laboral del Circuito (supera %s)" % money(TOPE_CUANTIA)],
           ["AT-2026-014 · Reparación plena",
            "Reparación plena: daño emergente, lucro cesante y perjuicio moral, más la indemnización "
            "por IPP: %s" % money(27284000),
            money(27284000),
            "Juez Laboral del Circuito (supera %s)" % money(TOPE_CUANTIA)],
           ["Reliquidación de aportes", "Diferencias de aportes de los últimos 36 meses: %s" %
            money(1850000), money(1850000),
            "Juez municipal de pequeñas causas, si existe en el lugar"]],
          [3.8, 5.6, 3.2, 3.4], size=8.4, align_centro=[2],
          titulo="Tabla 25. Aplicación del factor cuantía a los casos de la empresa")

    h(doc, 3, "Competencia funcional e impugnación")
    rich(doc, ["Fijada la competencia del juzgado laboral del circuito en primera instancia, la "
               "impugnación se rige por el factor funcional: la apelación corresponde al Tribunal "
               "Tribunal Superior del Distrito Judicial de Cundinamarca, Sala Laboral, y el recurso "
               "extraordinario de "
               "casación a la Sala de Casación Laboral de la Corte Suprema de Justicia, cuando la "
               "cuantía y la naturaleza del asunto lo permitan. El proceso se tramita por el "
               "procedimiento ordinario laboral de la Ley 1149 de 2007, con audiencias orales y "
               "públicas."])
    tabla(doc, ["Etapa", "Actuación", "Plazo o término"],
          [["1", "Radicación de la reclamación o del derecho de petición", "Antes de demandar; "
            "suspende la prescripción"],
           ["2", "Respuesta de la entidad o configuración del silencio", "15 días hábiles (Ley 1755 "
            "de 2015); un mes para la reclamación del art. 6 del C.P.T.S.S."],
           ["3", "Conciliación extrajudicial, si el asunto es conciliable", "Requisito de "
            "procedibilidad (Ley 640 de 2001)"],
           ["4", "Presentación de la demanda con anexos", "Dentro del término de prescripción: 3 años "
            "(art. 151, C.P.T.S.S.)"],
           ["5", "Audiencia del artículo 77 del C.P.T.S.S.", "Dentro de los 3 meses siguientes a la "
            "notificación de la demanda"],
           ["6", "Audiencia de trámite y juzgamiento", "Dentro de los 3 meses siguientes"],
           ["7", "Sentencia, apelación o casación", "Según el factor funcional"]],
          [1.3, 8.6, 6.1], size=8.4, align_centro=[0],
          titulo="Tabla 26. Ruta procesal del reclamo")

    h(doc, 3, "Requisitos formales del escrito de demanda")
    tabla(doc, ["Requisito", "Contenido exigible", "Fuente"],
          [["Designación del juez", "Juez Laboral del Circuito de Soacha (reparto)",
            "Arts. 11 y 12, C.P.T.S.S."],
           ["Identificación de las partes", "Nombre, domicilio y datos de notificación del "
            "demandante y del demandado", "Art. 25, C.P.T.S.S."],
           ["Hechos", "Narración cronológica, clara y numerada, con indicación de las pruebas",
            "Art. 25, num. 3, C.P.T.S.S."],
           ["Pretensiones", "Declarativas y de condena, determinadas y coherentes con los hechos",
            "Art. 25, num. 4, C.P.T.S.S."],
           ["Fundamentos de derecho", "Normas sustanciales y jurisprudencia aplicable",
            "Art. 25, num. 5, C.P.T.S.S."],
           ["Cuantía", "Estimación razonada de las pretensiones", "Art. 25, num. 6, C.P.T.S.S."],
           ["Pruebas", "Solicitud y práctica de las pruebas documentales, testimoniales y periciales",
            "Art. 25, num. 7, C.P.T.S.S."],
           ["Anexos", "Poder, copias, prueba del agotamiento y de la existencia y representación legal",
            "Art. 26, C.P.T.S.S."]],
          [3.8, 7.6, 4.6], size=8.4,
          titulo="Tabla 27. Lista de verificación del escrito de demanda")

    caja(doc, "Cierre del eje procesal",
         ["El insumo procesal es el punto en que el diagnóstico se convierte en actuación jurídica, "
          "y ahí un formato sirve de poco. Un escrito bien construido (con hechos verificables, "
          "fundamentos precisos, pretensiones congruentes y competencia correctamente fijada) multiplica las "
          "posibilidades de éxito y reduce el riesgo de que el proceso se pierda por razones ajenas "
          "al derecho sustancial."],
         tipo='info')
    salto_pagina(doc)
