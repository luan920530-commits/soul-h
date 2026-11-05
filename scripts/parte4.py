# -*- coding: utf-8 -*-
"""EJE IV: Insumo procesal y determinación de competencia judicial."""
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx_engine import (p, rich, h, bullet, cita, tabla, caja, salto_pagina,
                         aplicar_numeracion, AZUL, VINO, GRIS, GRIS_CLARO, NEGRO,
                         HEX_AZUL, HEX_CELESTE, sombrear, bordes_celda, set_col_widths,
                         margenes_celda, linea_horizontal, FUENTE_SANS)
from datos import SMLMV, money, IBC_TOTAL, TOPE_CUANTIA, ENTIDADES, EMPRESA
from datos import REPARACION_PLENA, MESADA_RAIS_SIN_GARANTIA, DIF_MENSUAL

AUTOR = "Julia Inés Buitrago García"
NIT = "39626936-8"
CC = "39626936"
DIR = "carrera 4 n.° 60-17, barrio La Despensa, comuna 3"
TEL = EMPRESA["telefono"]
MAIL = EMPRESA["correo"]
CIUDAD_FECHA = "Soacha (Cundinamarca), noviembre de 2026"
EPS = ENTIDADES["eps"]
ARL = ENTIDADES["arl"]
AFP_RPM = ENTIDADES["afp_rpm"]
AFP_RAIS = ENTIDADES["afp_rais"]
CAJA = ENTIDADES["caja"]


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
    r = par.add_run("DOCUMENTO PARA RADICACIÓN")
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
    p(doc, "C.C. %s  ·  Representante legal de EXTINTORES S C M  ·  NIT %s" % (CC, NIT),
      size=9, color=GRIS, space_after=2)
    p(doc, "Correo electrónico de notificaciones: %s  ·  Teléfono: %s" % (MAIL, TEL),
      size=9, color=GRIS, space_after=2)


# ================================================================ EJE IV
def eje_iv(doc):
    salto_pagina(doc)
    h(doc, 1, "EJE IV · RECLAMACIÓN ADMINISTRATIVA Y COMPETENCIA JUDICIAL")

    # ---------- 4.1
    h(doc, 2, "Calificación del mecanismo jurídico aplicable")
    rich(doc, ["La guía del proyecto acierta al separar dos situaciones que se confunden todo el "
               "tiempo: lo que se hace contra una entidad pública y lo que se hace contra una empresa "
               "privada. No es una diferencia de forma. La reclamación administrativa previa del "
               "artículo 6 del C.P.T.S.S. es requisito de procedibilidad frente a entidades de "
               "derecho público, y su omisión no se sanea. Aplicarla a un particular es trasladarle "
               "al empleador una carga que el legislador reservó para la Administración, y encima con "
               "el riesgo de que el escrito no interrumpa nada."])
    rich(doc, ["Basta leer la norma. El artículo 6 habla de «las acciones contenciosas contra la "
               "Nación, las entidades territoriales y cualquiera otra entidad de la administración "
               "pública», y la Corte, al examinarla en la Sentencia C-792 de 2006, la entendió como "
               "un «aseguramiento de la oportunidad para la autotutela administrativa»: la "
               "Administración revisa lo que hizo antes de que un juez lo revise. Ese presupuesto, "
               "frente a un empleador privado, simplemente no existe."])
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

    rich(doc, ["Aplicado a EXTINTORES S C M, esto deja tres caminos distintos. Contra el empleador, "
               "derecho de petición y solicitud de agotamiento de la vía interna. Contra las "
               "decisiones de la EPS, la ARL o la AFP, derecho de petición y, agotada la respuesta, "
               "demanda ante la jurisdicción laboral. Y la conciliación extrajudicial solo cabe donde "
               "hay derecho incierto y discutible: las semanas efectivamente cotizadas o una "
               "prestación ya reconocida no se concilian, porque no hay nada que negociar."])
    caja(doc, "El error que hunde expedientes",
         ["Radicar una «reclamación administrativa previa» contra una empresa privada no suspende la "
          "prescripción, y el juez puede declarar la demanda mal formulada. Al revés también pasa: "
          "demandar a una entidad pública sin reclamación previa termina en rechazo por falta de "
          "requisito de procedibilidad. Por eso lo primero que hay que decidir en cualquier caso no "
          "es el argumento, es quién es el demandado."],
         tipo='alerta')

    # ---------- 4.2 documentos
    h(doc, 2, "Insumo procesal 1: derecho de petición ante la ARL")
    rich(doc, ["El primer escrito va a la ARL y lleva dos pretensiones conexas: que se revise la "
               "clasificación del riesgo y que se reconozcan las prestaciones del accidente "
               "AT-2026-014. El modelo cumple lo que pide la Ley 1755 de 2015 (autoridad "
               "destinataria, identificación del peticionario, objeto, razones y petición concreta) y "
               "le suma dos cosas que casi nadie incluye: la solicitud de copias y la advertencia "
               "sobre el silencio administrativo. Lo segundo sí importa: el silencio es lo que "
               "permite contar el término desde una fecha cierta."])
    salto_pagina(doc)
    encabezado_escrito(doc, "%s  ·  OFICINA DE ATENCIÓN AL USUARIO" % ARL.upper(), CIUDAD_FECHA)
    rotulo(doc, "Referencia")
    parrafo_doc(doc, "Asunto: derecho de petición de interés particular. Solicitud de revisión de la "
                     "clasificación del riesgo, reconocimiento de prestaciones derivadas del "
                     "accidente de trabajo del 12 de agosto de 2026 y expedición de copias.",
                     bold=True, size=10.5)
    rotulo(doc, "Identificación del peticionario")
    parrafo_doc(doc, "%s, mayor de edad, identificada con cédula de ciudadanía número %s expedida en "
                     "Soacha, en calidad de representante legal y titular de la empresa EXTINTORES S C "
                     "M, identificada con cédula/NIT %s, con domicilio en %s, %s (Cundinamarca), "
                     "correo electrónico %s y teléfono %s, actuando en nombre propio y en ejercicio "
                     "del derecho fundamental de petición consagrado en el artículo 23 de la "
                     "Constitución Política y regulado por la Ley 1755 de 2015, presento la siguiente "
                     "solicitud."
                     % (AUTOR, CC, NIT, DIR, EMPRESA["municipio"], MAIL, TEL))

    rotulo(doc, "Hechos")
    for i, t in enumerate([
        "La empresa EXTINTORES S C M se encuentra afiliada a esa administradora con el código de "
        "actividad 1475901, clase de riesgo I, y tiene a su cargo tres (3) trabajadores dependientes, "
        "todos afiliados al Sistema General de Riesgos Laborales. La fecha y el número de afiliación "
        "constan en el certificado de afiliación que se anexa.",
        "La actividad económica principal registrada en el Registro Único Tributario corresponde al "
        "código CIIU 4759 (comercio al por menor de otros artículos domésticos en establecimientos "
        "especializados), que la tabla del Decreto 768 de 2022 clasifica en clase de riesgo I, "
        "subactividad 4759-01, con tarifa inicial del 0,522 %% del ingreso base de cotización. Ese "
        "código coincide con el reportado a la Secretaría de Salud de Soacha en el acta "
        "028-050164-26 del 09/09/2026.",
        "En el mismo establecimiento se ejecutan, además, tareas de recarga, mantenimiento y prueba "
        "hidrostática de extintores, que no corresponden a la subactividad 4759-01 y que fueron "
        "advertidas por la autoridad sanitaria en la visita del 09/09/2026. La empresa solicita que "
        "esa administradora evalúe la clasificación a la luz del proceso real y no solo del código "
        "declarado.",
        "El 12 de agosto de 2026 el trabajador Luis Antonio Aguilar Rodríguez, C.C. 1073694256, "
        "sufrió un accidente de trabajo durante la prueba hidrostática de un cilindro extintor, que "
        "le produjo fractura conminuta de los dedos índice y medio de la mano derecha y lesión del "
        "tendón flexor.",
        "El accidente fue reportado a esa administradora el 17 de agosto de 2026, cinco días después "
        "de su ocurrencia, y fue atendido por la red prestadora de la EPS %s." % EPS,
        "A la fecha de esta solicitud esa administradora ha reconocido el subsidio por incapacidad "
        "temporal, pero no ha notificado decisión de fondo sobre la calificación de la pérdida de "
        "capacidad laboral ni sobre la indemnización por incapacidad permanente parcial.",
    ], 1):
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
    for i, t in enumerate([
        "Que se revise la clasificación de la actividad y, si corresponde, se ajuste la clase de "
        "riesgo de EXTINTORES S C M con efectos a partir del periodo en que se produjo la "
        "discrepancia, previa evaluación del proceso de recarga y prueba hidrostática.",
        "Que se expida copia de la decisión de clasificación del riesgo y de los soportes técnicos en "
        "que se fundamentó.",
        "Que se confirme por escrito el reconocimiento y pago del subsidio por incapacidad temporal "
        "derivado del accidente del 12 de agosto de 2026, con la liquidación aplicada.",
        "Que se practique la calificación del origen y del porcentaje de pérdida de la capacidad "
        "laboral del trabajador, con sujeción al manual único adoptado por el Decreto 1507 de 2014, "
        "y se notifique el resultado.",
        "Que, previa calificación, se reconozca y pague la indemnización por incapacidad permanente "
        "parcial a que haya lugar.",
        "Que se expida copia íntegra del expediente administrativo del evento.",
        "Que se informe el nombre y el cargo del funcionario responsable de resolver y se notifique "
        "la respuesta al correo electrónico indicado.",
    ], 1):
        numerado_doc(doc, i, t)

    rotulo(doc, "Pruebas y anexos")
    for t in [
        "Copia del documento de identidad de la peticionaria y certificado de existencia y "
        "representación, si aplica.",
        "Registro Único Tributario y certificado de afiliación a esa administradora.",
        "Copias de las planillas de liquidación de aportes de los últimos doce periodos.",
        "Copia del reporte del accidente de trabajo y de los certificados de incapacidad.",
        "Copia del acta sanitaria 028-050164-26 del 09/09/2026 y de la historia clínica ocupacional, "
        "cuando exista.",
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
    rich(doc, ["El segundo va a la AFP y su objetivo es probatorio, no persuasivo. Se piden el "
               "expediente de afiliación y el acta de doble asesoría del caso P-03, o la constancia "
               "de que no existen. Cualquiera de las dos respuestas sirve: una prueba el vicio, la "
               "otra lo confirma por ausencia. Y con ese papel en la mano se decide si se demanda la "
               "ineficacia."])
    p(doc, "", space_after=6)
    encabezado_escrito(doc, "%s  ·  OFICINA DE SERVICIO AL CLIENTE" % AFP_RAIS.upper(), CIUDAD_FECHA)
    rotulo(doc, "Referencia")
    parrafo_doc(doc, "Asunto: derecho de petición. Solicitud de copia del expediente de afiliación y "
                     "del acta de doble asesoría de Luis Antonio Aguilar Rodríguez, y "
                     "reclamación por déficit de información en el traslado de régimen pensional.",
                     bold=True, size=10.5)
    rotulo(doc, "Identificación")
    parrafo_doc(doc, "%s, identificada con cédula de ciudadanía número %s, en calidad de "
                     "representante legal de EXTINTORES S C M (NIT %s) y empleadora de Luis Antonio "
                     "Aguilar Rodríguez, identificado con cédula de ciudadanía 1073694256, presento "
                     "derecho de petición y formulo reclamación en los siguientes términos."
                     % (AUTOR, CC, NIT))
    rotulo(doc, "Hechos")
    for i, t in enumerate([
        "Luis Antonio Aguilar Rodríguez está afiliado a esa administradora en el Régimen de Ahorro "
        "Individual con Solidaridad y cotiza como dependiente a través de EXTINTORES S C M. La fecha "
        "del traslado consta en el expediente que se solicita.",
        "A la fecha cuenta con 1.150 semanas cotizadas y 56 años de edad, y aspira a acceder a la "
        "pensión de vejez en el año 2027.",
        "El afiliado manifiesta que la decisión de traslado se produjo tras una visita de un asesor "
        "comercial, quien le indicó que obtendría una mesada superior, sin que se le entregara "
        "simulación, proyección o documento comparativo de los regímenes.",
        "El afiliado no conserva copia del formulario de afiliación ni recuerda haber suscrito acta "
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
    rich(doc, ["El tercero se dirige al propio empleador, que es lo que la guía exige cuando el "
               "demandado es privado. Cumple dos funciones al mismo tiempo: da la oportunidad de "
               "corregir el error sin llegar a juicio y deja documentada la reclamación para "
               "interrumpir la prescripción y probar el agotamiento de la vía interna."])
    p(doc, "", space_after=6)
    encabezado_escrito(doc, "EXTINTORES S C M  ·  GERENCIA GENERAL", CIUDAD_FECHA)
    rotulo(doc, "Referencia")
    parrafo_doc(doc, "Asunto: solicitud de agotamiento de la vía interna y reclamación por "
                     "diferencias en el reconocimiento de prestaciones económicas del Sistema de "
                     "Seguridad Social Integral.", bold=True, size=10.5)
    rotulo(doc, "Hechos")
    for i, t in enumerate([
        "La suscrita es titular y representante legal de la empresa EXTINTORES S C M, empleadora de "
        "tres (3) trabajadores dependientes.",
        "En la auditoría interna practicada en octubre de 2026 se identificaron diferencias en la "
        "liquidación de una licencia de paternidad y en el reporte de novedades de uno de los "
        "trabajadores.",
        "Las diferencias ascienden a %s y afectan el ingreso de Felipe Martínez (C.C. 8686834), por "
        "la licencia de paternidad, así como el reporte de las semanas de los tres trabajadores de "
        "la planta." % money(412000),
        "La empresa reconoce la procedencia de la corrección y requiere formalizar el trámite para "
        "que produzca efectos frente a la EPS %s y a las administradoras de pensiones." % EPS,
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
    rich(doc, ["Competencia se decide aplicando cuatro factores en orden: materia, territorio, "
               "cuantía y función. Fallar en cualquiera de los cuatro no anula el derecho, pero manda "
               "el expediente a otro despacho y el caso se enfría seis meses. Con tres trabajadores y "
               "una empresa pequeña, medio año es la diferencia entre cobrar y no cobrar."])

    h(doc, 3, "Competencia por razón de la materia")
    rich(doc, ["El artículo 2 del C.P.T.S.S. le da a la jurisdicción ordinaria laboral «las "
               "controversias referentes al sistema de seguridad social integral que se susciten "
               "entre los afiliados, beneficiarios o usuarios, los empleadores y las entidades "
               "administradoras o prestadoras, cualquiera que sea la naturaleza de la relación "
               "jurídica y de los actos jurídicos que se controviertan»."])

    rich(doc, ["Ese «cualquiera que sea» resuelve casi todo el Eje IV. La ineficacia del traslado, el "
               "reconocimiento de prestaciones de la ARL y la reliquidación de aportes van al juez "
               "laboral. No al administrativo, aunque haya una entidad pública en el fondo del "
               "asunto; no al civil, aunque se discutan mesadas."])

    h(doc, 3, "Competencia territorial (artículo 11 del C.P.T.S.S.)")
    rich(doc, ["El artículo 11 del C.P.T.S.S., modificado por el artículo 8 de la Ley 712 de 2001, "
               "dispone que en los procesos contra las entidades que conforman el sistema de "
               "seguridad social integral ",
               ("«será competente el juez laboral del circuito del lugar del domicilio de la entidad "
                "de seguridad social demandada o el del lugar donde se haya surtido la reclamación "
                "del respectivo derecho, a elección del demandante»", {"i": True}), ". La norma "
               "consagra, por tanto, un fuero concurrente: dos opciones y la elección en manos del "
               "demandante, que en la práctica es una garantía de acceso a la justicia. Contra "
               "particulares rige la regla general del artículo 5 (lugar de prestación del servicio o "
               "domicilio del demandado); contra la Nación, el artículo 7."])
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
    rich(doc, ["Y hay un efecto que casi nadie calcula: mientras está pendiente el agotamiento de la "
               "reclamación administrativa, el artículo 6 del C.P.T.S.S. suspende el término de "
               "prescripción. Eso obliga a radicar con constancia (sello, correo certificado o "
               "radicado electrónico) y a guardar el comprobante con la respuesta. Sin constancia, la "
               "suspensión no se puede probar y el efecto se pierde."])

    h(doc, 3, "Competencia por razón de la cuantía (artículo 12 del C.P.T.S.S.)")
    rich(doc, ["El artículo 12 del C.P.T.S.S., modificado por el artículo 46 de la Ley 1395 de 2010, "
               "traza la raya en veinte salarios mínimos. Por encima, juez laboral del circuito en "
               "única instancia; por debajo, primera instancia del circuito y, donde existan, jueces "
               "municipales de pequeñas causas en única instancia. Para 2026 la raya está en %s."
               % money(TOPE_CUANTIA)])
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
    rich(doc, ["Con el juzgado del circuito en primera instancia, la apelación va al Tribunal "
               "Superior del Distrito Judicial de Cundinamarca, Sala Laboral, y la casación a la Sala "
               "de Casación Laboral de la Corte Suprema, cuando la cuantía y la naturaleza del asunto "
               "lo permitan. El trámite es el ordinario laboral de la Ley 1149 de 2007, con "
               "audiencias orales y públicas."])
    rich(doc, ["La oralidad tiene una consecuencia práctica que conviene anticipar: el caso se gana o "
               "se pierde en la audiencia, y en la audiencia solo existe lo que está en el "
               "expediente. De ahí que el Eje IV termine donde empieza el Eje I: en los documentos."])
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

    caja(doc, "Lo que deja el Eje IV",
         ["Aquí el diagnóstico se vuelve actuación, y en ese paso un formato sirve de poco. Un escrito "
          "con hechos verificables, fundamentos precisos, pretensiones congruentes y competencia bien "
          "fijada no garantiza ganar, pero evita perder por razones que no tienen nada que ver con el "
          "derecho sustancial. En una empresa de tres personas, perder por eso es imperdonable."],
         tipo='info')
    salto_pagina(doc)
