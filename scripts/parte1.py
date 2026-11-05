# -*- coding: utf-8 -*-
"""Parte I: portada, tabla de contenido, resumen ejecutivo, metodología, introducción y EJE I."""
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx_engine import (p, rich, h, h_sin_num, bullet, cita, tabla, caja, salto_pagina,
                         linea_horizontal, imagen, add_tabla_contenido, sombrear, bordes_celda,
                         set_col_widths, AZUL, VINO, GRIS, GRIS_CLARO, HEX_AZUL, HEX_VINO,
                         FUENTE_SANS, BLANCO)
from datos import (SMLMV, AUX_TRANSPORTE, TOPE_AUX, TOPE_EXO, money, TARIFAS, NOMINA,
                   IBC_TOTAL, N_AUX, CONTRATISTA)

AUTOR = "Julia Inés Buitrago García"
RAZON = "EXTINTORES SCM"
NIT = "39626936-8"


# ---------------------------------------------------------------- Portada
def portada(doc):
    p(doc, "", space_after=34)
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    c = t.rows[0].cells[0]
    sombrear(c, HEX_AZUL)
    bordes_celda(c, color=HEX_AZUL)
    c.text = ""
    par = c.paragraphs[0]
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    par.paragraph_format.space_before = Pt(18)
    r = par.add_run("UNIVERSIDAD INCCA DE COLOMBIA")
    r.font.size = Pt(13); r.bold = True; r.font.color.rgb = BLANCO
    r.font.all_caps = True; r.font.name = FUENTE_SANS
    for txt, sz, cap in (("Facultad de Derecho", 10.5, False),
                         ("Asignatura: Derecho de la Seguridad Social", 10.5, False)):
        p2 = c.add_paragraph(); p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p2.add_run(txt); r.font.size = Pt(sz); r.font.color.rgb = BLANCO
        r.font.name = FUENTE_SANS
    p3 = c.add_paragraph(); p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p3.paragraph_format.space_after = Pt(16)
    r = p3.add_run("Proyecto integrador · 40 % de la nota del corte")
    r.font.size = Pt(9.5); r.font.color.rgb = BLANCO; r.font.name = FUENTE_SANS; r.italic = True
    set_col_widths(t, [16.0])

    p(doc, "", space_after=24)
    p(doc, "AUDITORÍA Y DIAGNÓSTICO DE SEGURIDAD SOCIAL ORGANIZACIONAL", size=21,
      bold=True, color=AZUL, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8).paragraph_format.line_spacing = 1.05
    p(doc, "Informe técnico-jurídico de consultoría", size=13, italic=True, color=GRIS,
      align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    p(doc, "Salud · Riesgos Laborales · Pensiones · Reclamación y competencia judicial",
      size=10.5, color=GRIS_CLARO, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=20)
    linea_horizontal(doc, color=HEX_AZUL, sz=12, space_after=14)

    p(doc, "ORGANIZACIÓN AUDITADA", size=9, bold=True, color=GRIS_CLARO,
      align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    p(doc, RAZON, size=19, bold=True, color=VINO, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    p(doc, "NIT 39626936-8 · Empresa privada", size=11, color=GRIS,
      align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    p(doc, "Comercio, mantenimiento y recarga de extintores y sistemas de protección contra incendio",
      size=9.5, color=GRIS_CLARO, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    p(doc, "Domicilio: Bogotá D.C., Colombia", size=9.5, color=GRIS_CLARO,
      align=WD_ALIGN_PARAGRAPH.CENTER, space_after=20)

    tabla(doc, ["Campo", "Información"],
          [["Elaborado por", AUTOR],
           ["Modalidad", "Trabajo académico colaborativo — proyecto integrador"],
           ["Periodo académico", "2026-2"],
           ["Fecha de presentación", "Noviembre de 2026"],
           ["Cifras de referencia", "SMMLV %s y auxilio de transporte %s (Decretos 1469 y 1470 de 2025)"
            % (money(SMLMV), money(AUX_TRANSPORTE))],
           ["Norma de pensiones aplicada", "Ley 100 de 1993, en tránsito a la Ley 2381 de 2024 "
            "(vigencia: 1.º de abril de 2027, Sentencia C-264 de 2026)"]],
          [4.4, 11.6], size=9)

    p(doc, "", space_after=8)
    p(doc, "Documento de trabajo académico. No constituye concepto jurídico ni asesoría profesional "
           "para un caso concreto. Las cifras de nómina y los casos identificados como ilustrativos "
           "deben sustituirse por los soportes reales de la organización antes de presentar el "
           "informe definitivo.",
      size=7.8, italic=True, color=GRIS_CLARO, align=WD_ALIGN_PARAGRAPH.CENTER)
    salto_pagina(doc)


# ---------------------------------------------------------------- Tabla de contenido
def contenido(doc):
    h_sin_num(doc, 1, "TABLA DE CONTENIDO")
    add_tabla_contenido(doc, "1-3")
    p(doc, "", space_after=4)
    caja(doc, "Cómo usar esta plantilla",
         ["El documento tiene dos tipos de contenido. El primero es el ",
          [("desarrollo jurídico ya redactado", {"b": True}),
           ": marco normativo, análisis de casos, liquidaciones, escritos procesales y "
           "conclusiones, con la norma y la jurisprudencia citadas en cada punto. El segundo son "
           "los "],
          [("campos por diligenciar", {"b": True}),
           ": los anexos del capítulo final y los espacios marcados con líneas de puntos, que debe "
           "llenar con los soportes reales de EXTINTORES SCM (planillas PILA, comprobantes de "
           "nómina, certificados de afiliación, historia laboral y documentos del SG-SST)."],
          "Al abrir el archivo, actualice la tabla de contenido: clic derecho sobre el texto azul → "
          "«Actualizar campo» → «Actualizar toda la tabla». Los recuadros de color son ayudas de "
          "estudio: los azules orientan la lectura, los rojos señalan riesgos jurídicos y los ámbar "
          "advierten puntos de verificación."],
         tipo='info')
    salto_pagina(doc)


# ---------------------------------------------------------------- Resumen ejecutivo
def resumen_ejecutivo(doc):
    h_sin_num(doc, 1, "RESUMEN EJECUTIVO")
    rich(doc, [
        "El presente informe expone los hallazgos de la auditoría jurídica en seguridad social "
        "practicada a ", (RAZON, {"b": True}), ", empresa privada identificada con NIT ",
        (NIT, {"b": True}),
        ", cuya actividad principal es el comercio, la recarga y el mantenimiento de extintores, "
        "equipos y sistemas de protección contra incendio, con domicilio en Bogotá D.C. La auditoría "
        "se estructura en los cuatro ejes exigidos por la guía del proyecto integrador y se apoya en "
        "la normativa vigente —Ley 100 de 1993, Ley 1562 de 2012, Ley Estatutaria 1751 de 2015, "
        "Código Sustantivo del Trabajo y Código Procesal del Trabajo y de la Seguridad Social—, en la "
        "jurisprudencia de la Corte Suprema de Justicia y de la Corte Constitucional, y en las normas "
        "de 2026 que inciden de manera directa en la organización: la Ley 2381 de 2024, cuyo "
        "calendario de entrada en vigencia fue fijado por la Sentencia C-264 de 2026, y los Decretos "
        "1469 y 1470 de 2025 sobre salario mínimo y auxilio de transporte.", ])

    p(doc, "Principales hallazgos", size=11.5, bold=True, color=AZUL, space_before=8, space_after=4)
    for titulo, texto in [
        ("Eje I · Afiliación y cotizaciones (riesgo medio). ",
         "La empresa cumple la afiliación obligatoria al Régimen Contributivo, pero se detectaron "
         "novedades reportadas por fuera del plazo, un beneficiario mayor de 25 años sin soporte de "
         "dependencia económica y una diferencia de $ 412.000 en el pago de una licencia de "
         "maternidad. La condición de persona natural con diez trabajadores permite aplicar la "
         "exoneración del artículo 114-1 del Estatuto Tributario, siempre que se verifique su "
         "procedencia y se reporte correctamente en la planilla."),
        ("Eje II · Riesgos laborales (riesgo alto). ",
         "La clasificación correcta corresponde a la clase de riesgo I, con tarifa inicial de "
         "0,522 % del IBC (Decreto 768 de 2022 y artículo 13 del Decreto 1772 de 1994), y no a la "
         "clase II que la organización venía aplicando. El caso AT-2026-014 revela el "
         "incumplimiento de seis de siete deberes de seguridad verificables, lo que configura culpa "
         "patronal del artículo 216 del C.S.T., con una exposición estimada de $ 27,3 millones de "
         "reparación plena, de los cuales $ 18,2 millones quedarían a cargo del empleador después de "
         "descontar las prestaciones de la ARL."),
        ("Eje III · Pensiones (riesgo alto). ",
         "El 70 % de la planta está en el Régimen de Prima Media y el 30 % en el Régimen de Ahorro "
         "Individual. El caso P-03 presenta un traslado de régimen sin soporte documental de doble "
         "asesoría: la demanda de ineficacia es viable y los efectos restitutorios comprenden el "
         "traslado íntegro del ahorro y una diferencia de mesada de $ 345.095 mensuales frente al "
         "escenario de ahorro individual, equivalente a $ 82,8 millones en veinte años de disfrute."),
        ("Eje IV · Insumo procesal (riesgo bajo). ",
         "Se redactan tres documentos modelo —derecho de petición ante la ARL, derecho de petición "
         "ante la AFP y solicitud de agotamiento de la vía interna ante el empleador— y se fija la "
         "competencia del juez laboral del circuito de Bogotá por los factores de materia, territorio "
         "(artículo 11 del C.P.T.S.S.) y cuantía (artículo 12)."),
    ]:
        bullet(doc, [(titulo, {"b": True}), texto])

    caja(doc, "Conclusión general",
         ["El cumplimiento formal es aceptable, pero existen tres brechas materiales: (i) la "
          "clasificación del riesgo ante la ARL no corresponde a la actividad económica registrada; "
          "(ii) el Sistema de Gestión de Seguridad y Salud en el Trabajo no produce la evidencia "
          "documental que exoneraría al empleador en un accidente de trabajo; y (iii) el archivo de "
          "asesoría pensional es insuficiente frente al estándar de doble asesoría de la Ley 1748 de "
          "2014 y la Sentencia SU-107 de 2024. El plan de mejoramiento del capítulo 5 cuantifica "
          "cada acción con responsables, plazos e indicadores verificables."],
         tipo='alerta')
    salto_pagina(doc)


# ---------------------------------------------------------------- Metodología
def metodologia(doc):
    h_sin_num(doc, 1, "ASPECTOS METODOLÓGICOS")
    h_sin_num(doc, 2, "Tipo y método de investigación")
    rich(doc, ["La investigación es de tipo ",
               ("jurídico-dogmático con componentes descriptivo y propositivo", {"i": True}),
               ", complementada con un estudio de caso organizacional. El enfoque es cualitativo y la "
               "técnica principal es la revisión documental: se examinan fuentes formales del derecho "
               "(Constitución, leyes, decretos y jurisprudencia) y fuentes materiales de la empresa "
               "(contratos, comprobantes de nómina, planillas integradas de liquidación de aportes, "
               "certificados de afiliación, historia laboral y documentos del SG-SST). El método "
               "combina la subsunción normativa —contrastar el hecho con la hipótesis legal— con el "
               "razonamiento jurisprudencial, en especial la ",
               ("ratio decidendi", {"i": True}), " de las providencias citadas."])
    rich(doc, ["El informe no agota el estudio de todas las obligaciones del sistema: delimita el "
               "objeto a los cuatro ejes fijados por la guía académica, de modo que los hallazgos son "
               "válidos para esos contornos y no para la totalidad del cumplimiento normativo de la "
               "empresa."])

    h_sin_num(doc, 2, "Técnicas e instrumentos")
    tabla(doc, ["Instrumento", "Contenido", "Eje"],
          [["Matriz de caracterización organizacional",
            "Razón social, NIT, naturaleza jurídica, actividad económica (CIIU), planta de personal y "
            "tipo de vinculación", "I"],
           ["Ficha de liquidación del IBC",
            "Reconstrucción del IBC de trabajadores dependientes e independientes, con topes y "
            "exclusiones legales", "I"],
           ["Lista de verificación de afiliación y cobertura",
            "Afiliación a EPS, ARL y AFP, inclusión de beneficiarios y trámite de prestaciones "
            "económicas", "I y II"],
           ["Matriz de clasificación del riesgo",
            "Contraste del código CIIU con la tabla del Decreto 768 de 2022 y la tarifa del Decreto "
            "1772 de 1994", "II"],
           ["Árbol de decisión ATEP",
            "Nexo causal, prestaciones de la ARL y culpa patronal del artículo 216 del C.S.T.", "II"],
           ["Mapa pensional y auditoría individual",
            "Distribución por régimen y verificación de requisitos, deber de información e "
            "ineficacia del traslado", "III"],
           ["Modelos de escritos procesales",
            "Derecho de petición, agotamiento de vía interna y fijación de competencia", "IV"]],
          [4.3, 9.4, 2.3], size=8.8, align_centro=[2],
          titulo="Tabla 1. Instrumentos de auditoría aplicados")

    h_sin_num(doc, 2, "Fuentes y criterios de valoración")
    rich(doc, ["Se emplean fuentes primarias (normas y jurisprudencia) y secundarias (doctrina y "
               "documentos de la empresa). Los hallazgos se valoran con una escala de tres niveles, "
               "explicitada en cada eje: ", ("riesgo bajo", {"b": True}),
               " (la obligación se cumple y está documentada), ", ("riesgo medio", {"b": True}),
               " (se cumple, pero sin evidencia suficiente o con errores de liquidación) y ",
               ("riesgo alto", {"b": True}),
               " (existe incumplimiento o exposición a responsabilidad patronal). El criterio permite "
               "que las conclusiones sean verificables y no dependan de apreciaciones subjetivas."])
    rich(doc, ["Advertencia metodológica: las cifras de nómina y los casos señalados como "
               "ilustrativos se construyeron con el salario mínimo de 2026 y las tarifas vigentes, "
               "para que los cálculos sean reproducibles. En la aplicación real deben sustituirse por "
               "los comprobantes y planillas de EXTINTORES SCM."])
    salto_pagina(doc)


# ---------------------------------------------------------------- Introducción
def introduccion(doc):
    h_sin_num(doc, 1, "INTRODUCCIÓN Y PLANTEAMIENTO DEL PROBLEMA")
    rich(doc, ["El Sistema de Seguridad Social Integral colombiano, diseñado por la Ley 100 de 1993 y "
               "desarrollado por un cuerpo normativo profuso, impone a los empleadores obligaciones "
               "que no se agotan en el pago de aportes: exige afiliar, reportar novedades con "
               "veracidad, gestionar prestaciones económicas, administrar el riesgo laboral y "
               "garantizar información comprensible en materia pensional. El incumplimiento de esos "
               "deberes produce consecuencias que van desde la sanción moratoria y las acciones de "
               "cobro de las administradoras (artículos 23 y 24 de la Ley 100 de 1993) hasta la "
               "responsabilidad patrimonial plena del artículo 216 del Código Sustantivo del Trabajo, "
               "pasando por la ineficacia de actos jurídicos."])
    rich(doc, ["En ese contexto, la microempresa es el escenario donde las brechas se manifiestan con "
               "mayor crudeza: la gestión de la seguridad social suele concentrarse en una sola "
               "persona, se confunde la afiliación con el pago y se descuida la prueba. ",
               RAZON, " no es ajena a ese patrón. La actividad que desarrolla —manipulación de "
               "cilindros presurizados, recarga con agentes extintores, pruebas hidrostáticas y "
               "transporte de equipo— introduce riesgos específicos que exigen una gestión del Sistema "
               "General de Riesgos Laborales técnicamente informada, mientras que su condición de "
               "persona natural empleadora plantea problemas particulares en materia de exoneración "
               "de aportes y de responsabilidad patrimonial."])
    rich(doc, ["El problema jurídico que orienta la auditoría puede formularse así: ",
               ("¿en qué medida el cumplimiento de las obligaciones de seguridad social de "
                "EXTINTORES SCM satisface el estándar normativo vigente y qué consecuencias "
                "jurídicas se derivan de las brechas identificadas en materia de salud, riesgos "
                "laborales, pensiones y acceso a la administración de justicia?", {"i": True})])
    rich(doc, ["Para responderlo, el informe se divide en cinco capítulos: el primero caracteriza la "
               "organización y analiza la afiliación, el IBC, la cobertura y las prestaciones "
               "económicas; el segundo aborda el Sistema General de Riesgos Laborales con el estudio "
               "de un caso de accidente de trabajo y el análisis de la culpa patronal; el tercero "
               "examina el diagnóstico pensional y la ineficacia del traslado; el cuarto contiene los "
               "insumos procesales y la fijación de competencia; y el quinto sintetiza los hallazgos, "
               "las recomendaciones y el plan de mejoramiento, con las conclusiones y la preparación "
               "de la sustentación oral."])
    salto_pagina(doc)


# ================================================================ EJE I
def eje_i(doc):
    salto_pagina(doc)
    h(doc, 1, "EJE I · CARACTERIZACIÓN Y RÉGIMEN DE AFILIACIÓN (SALUD Y COTIZACIONES)")

    # ---------- 1.1
    h(doc, 2, "Ficha técnica de la organización")
    rich(doc, ["La caracterización es el punto de partida de toda auditoría: sin identificar con "
               "precisión la naturaleza jurídica del empleador, su actividad económica y su planta de "
               "personal no es posible determinar cuáles obligaciones le son exigibles ni en qué "
               "cuantía. La siguiente ficha se construyó con los datos de la organización y el "
               "contraste con el Registro Único Tributario."])
    tabla(doc, ["Variable", "Dato", "Relevancia jurídica"],
          [["Razón social", RAZON,
            "Identifica al empleador como sujeto de obligaciones (art. 22, Ley 100 de 1993)"],
           ["NIT", NIT,
            "Se usa en la planilla PILA y en la identificación de los escritos procesales"],
           ["Naturaleza jurídica", "Empresa privada; persona natural comerciante (empresa unipersonal)",
            "Determina el régimen aplicable: no hay reclamación administrativa previa del art. 6 del "
            "C.P.T.S.S., sino solicitud directa al empleador y, en su caso, conciliación extrajudicial"],
           ["Actividad económica", "Comercio al por menor de artículos de ferretería, pinturas y "
            "productos de vidrio (CIIU 4752), con actividades conexas de recarga y mantenimiento de "
            "extintores",
            "El código CIIU define la clase de riesgo ante la ARL (Decreto 768 de 2022) y el "
            "tratamiento parafiscal"],
           ["Domicilio", "Bogotá D.C.", "Fija la competencia territorial del juez laboral (art. 11 del C.P.T.S.S.)"],
           ["Representante legal", AUTOR, "Única titular y responsable de las obligaciones patronales"],
           ["Planta de personal", "%d trabajadores dependientes con contrato de trabajo" % len(NOMINA),
            "La condición de persona natural con dos o más trabajadores activa el régimen general de "
            "aportes (art. 114-1 del E.T.)"],
           ["Contratistas", "1 contrato de prestación de servicios (contador público)",
            "Exige verificar la afiliación del contratista como independiente (art. 89, Ley 100 de 1993)"],
           ["Tamaño empresarial", "Microempresa",
            "No altera las obligaciones de afiliación, que son imperativas; sí incide en los "
            "estándares mínimos del SG-SST (Resolución 0312 de 2019)"]],
          [3.3, 6.5, 6.2], size=8.6,
          titulo="Tabla 2. Ficha técnica de EXTINTORES SCM",
          fuente="Fuente: elaboración propia con base en la información de la organización y el RUT.")

    caja(doc, "Advertencia sobre la naturaleza jurídica",
         ["La actividad se ejerce bajo la forma de empresa unipersonal de persona natural. Ello tiene "
          "dos consecuencias que suelen pasarse por alto: (i) el patrimonio personal del titular "
          "responde por las obligaciones sociales, sin la separación patrimonial que ofrece una "
          "sociedad comercial; y (ii) la exoneración de aportes del artículo 114-1 del Estatuto "
          "Tributario exige, para las personas naturales, emplear dos o más trabajadores, de modo que "
          "el beneficio debe verificarse antes de aplicarlo en la planilla."],
         tipo='alerta')

    # ---------- 1.2
    h(doc, 2, "Marco normativo de la afiliación y la cotización en salud")
    rich(doc, ["El Sistema General de Seguridad Social en Salud (SGSSS) se organiza en dos regímenes: "
               "el contributivo, financiado con cotizaciones de los afiliados y sus empleadores, y el "
               "subsidiado, financiado con recursos fiscales y del Fondo de Solidaridad y Garantía. La "
               "afiliación al primero es obligatoria para los trabajadores dependientes e "
               "independientes con capacidad de pago, y el empleador debe inscribir a sus trabajadores "
               "en la EPS que estos elijan, sin que la afiliación colectiva pueda coartar esa libertad "
               "(numeral 1 del artículo 161 de la Ley 100 de 1993)."])
    rich(doc, ["La Ley Estatutaria 1751 de 2015 elevó la salud a derecho fundamental autónomo e "
               "irrenunciable y fijó sus elementos esenciales —disponibilidad, aceptabilidad, "
               "accesibilidad, calidad e idoneidad profesional—, además de los principios de "
               "progresividad, integralidad y equidad. Esa ley prohíbe a las entidades administradoras "
               "y a los empleadores negar, retardar o condicionar la afiliación o el acceso a los "
               "servicios, lo que refuerza el deber patronal de reportar novedades de manera oportuna "
               "y veraz."])

    h(doc, 3, "Obligaciones del empleador en materia de salud")
    for t in [
        [("Afiliar", {"b": True}), " a todos los trabajadores con vinculación laboral, verbal o "
         "escrita, temporal o permanente, y reportar el ingreso en la planilla integrada (art. 161, "
         "num. 1, Ley 100 de 1993)."],
        [("Descontar y pagar", {"b": True}), " la cotización: el empleador responde por la totalidad "
         "del aporte, aun cuando no hubiere efectuado el descuento al trabajador (art. 22, Ley 100 de "
         "1993)."],
        [("Reportar novedades", {"b": True}), " de ingreso, retiro, variación de salario, licencias e "
         "incapacidades dentro de los plazos de la planilla integrada."],
        [("Garantizar la cobertura familiar", {"b": True}), " incluyendo a los beneficiarios del "
         "núcleo familiar (art. 163, Ley 100 de 1993)."],
        [("Tramitar y recobrar", {"b": True}), " las prestaciones económicas: adelantar el pago de "
         "licencias e incapacidades cuando corresponda y recobrar ante la EPS."],
    ]:
        bullet(doc, t)

    h(doc, 3, "Monto y distribución de la cotización")
    rich(doc, ["La cotización al Régimen Contributivo equivale al ",
               ("12,5 % del ingreso o salario base de cotización", {"b": True}),
               " (art. 204 de la Ley 100 de 1993, modificado por la Ley 1122 de 2007). De ese "
               "porcentaje, dos terceras partes (8,5 %%) están a cargo del empleador y una tercera "
               "parte (4 %%) a cargo del trabajador. Para 2026, el aporte patronal sobre un salario "
               "mínimo asciende a %s y el del trabajador a %s." % (
                   money(SMLMV * 0.085), money(SMLMV * 0.04))])
    imagen(doc, "assets/f2_aportes.png", 15.4,
           "Figura 1. Distribución de la cotización mensual al SGSSS para un trabajador dependiente. "
           "Fuente: elaboración propia con base en el artículo 204 de la Ley 100 de 1993 y el "
           "artículo 114-1 del Estatuto Tributario.")

    h(doc, 3, "La exoneración del artículo 114-1 del Estatuto Tributario")
    rich(doc, ["El artículo 114-1 del Estatuto Tributario exonera del pago de aportes parafiscales al "
               "SENA y al ICBF y de la cotización al Régimen Contributivo de Salud a las sociedades y "
               "personas jurídicas contribuyentes del impuesto de renta, respecto de los trabajadores "
               "que devengan individualmente menos de diez (10) salarios mínimos. El inciso segundo "
               "extiende el beneficio a las personas naturales empleadoras, ",
               ("pero excluye expresamente a quienes empleen menos de dos trabajadores", {"b": True}),
               ". La exoneración nunca cubre los aportes a pensión ni los de riesgos laborales, y el "
               "aporte del 4 % a cargo del trabajador se mantiene intacto."])
    rich(doc, ["Para EXTINTORES SCM la conclusión es concreta: al contar con diez trabajadores "
               "dependientes, la empresa ", ("puede aplicar la exoneración", {"b": True}),
               " respecto de los trabajadores cuya remuneración salarial sea inferior a diez salarios "
               "mínimos (%s en 2026), siempre que sea contribuyente declarante del impuesto sobre la "
               "renta y que la novedad se reporte correctamente en la planilla. Si la planta se "
               "redujera a un solo trabajador, la exoneración desaparecería y habría que reliquidar "
               "los aportes con intereses moratorios." % money(TOPE_EXO)])
    caja(doc, "Punto de verificación en la planilla",
         ["La exoneración se declara en la planilla con el código correspondiente. Aplicarla sin "
          "derecho genera un faltante de aportes que la ADRES o la EPS pueden cobrar con intereses "
          "(art. 23, Ley 100 de 1993) y que, en un proceso judicial, se traduce en tiempo no cotizado "
          "para el trabajador. Verifique además que el beneficio se aplique solo sobre la base "
          "salarial y no sobre los pagos no constitutivos de salario."],
         tipo='nota')

    # ---------- 1.3
    h(doc, 2, "Liquidación del Ingreso Base de Cotización (IBC)")
    rich(doc, ["El IBC es la magnitud económica sobre la cual se calculan los aportes a salud, "
               "pensiones y riesgos laborales. No equivale al salario bruto ni al valor pagado al "
               "trabajador: es una construcción jurídica que resulta de depurar el salario de los "
               "pagos que la ley excluye y de aplicar los topes legales. Su liquidación incorrecta "
               "produce dos tipos de daño: al sistema (faltantes de aportes) y al trabajador "
               "(disminución de su mesada pensional o de sus prestaciones económicas)."])

    h(doc, 3, "Regla general para trabajadores dependientes")
    rich(doc, ["El artículo 127 del Código Sustantivo del Trabajo define el salario como todo lo que "
               "recibe el trabajador en dinero o en especie como contraprestación directa del "
               "servicio, sin que importe la denominación del pago. Integran el IBC, por tanto, el "
               "salario básico, las horas extras y recargos, las comisiones, los porcentajes sobre "
               "ventas, los viáticos permanentes y las primas habituales. El artículo 128 excluye los "
               "pagos no salariales: el auxilio de transporte, los viáticos accidentales, la dotación, "
               "las propinas y las primas extralegales. El auxilio de transporte no integra el IBC, "
               "pero sí se suma a la base de cesantías y prima de servicios, distinción que suele "
               "confundirse en la práctica."])
    rich(doc, ["Sobre la base así depurada operan dos límites: el ",
               ("piso de un salario mínimo legal mensual vigente", {"b": True}), " y el ",
               ("techo de veinticinco salarios mínimos", {"b": True}),
               ", conforme al artículo 18 de la Ley 100 de 1993 y al artículo 3 del Decreto 1772 de "
               "1994. En el salario integral, la cotización se calcula sobre el 70 % del salario "
               "pactado (art. 132 del C.S.T. y art. 18 de la Ley 100 de 1993)."])
    imagen(doc, "assets/f1_ibc.png", 15.4,
           "Figura 2. Ruta de liquidación del IBC para trabajadores dependientes e independientes. "
           "Fuente: elaboración propia con base en los artículos 127, 128 y 132 del C.S.T., 18 y 89 "
           "de la Ley 100 de 1993 y los Decretos 1772 de 1994 y 1703 de 2002.")

    h(doc, 3, "Regla especial para contratistas independientes")
    rich(doc, ["El contratista de prestación de servicios personales no cotiza sobre lo que recibe, "
               "sino sobre el ",
               ("40 % del valor mensualizado del contrato, sin incluir el IVA", {"b": True}),
               " (art. 89 de la Ley 100 de 1993 y art. 23 del Decreto 1703 de 2002). El 60 % "
               "restante se entiende destinado a cubrir los costos de la actividad. Cuando el 40 % "
               "resulta inferior a un salario mínimo, la base se ajusta al mínimo; el retardo en el "
               "pago de los honorarios no altera la base; y el contratista asume la totalidad de la "
               "cotización: 12,5 % de salud, 16 % de pensión y el porcentaje de riesgos laborales "
               "según la clase de riesgo de la actividad."])
    rich(doc, ["Para el contrato de prestación de servicios del contador público, con honorarios de "
               "%s mensuales, el 40 %% equivale a %s, cifra superior a un salario mínimo; ese es su "
               "IBC y su aporte mensual a salud y pensión asciende a %s." % (
                   money(CONTRATISTA['honorarios']), money(CONTRATISTA['honorarios'] * 0.4),
                   money(CONTRATISTA['honorarios'] * 0.4 * 0.285))])

    h(doc, 3, "Aplicación a la nómina de EXTINTORES SCM")
    filas = []
    for cod, cargo, tipo, ibc, reg, obs in NOMINA:
        filas.append([cod, cargo, tipo, money(ibc), money(ibc * 0.085), money(ibc * 0.04),
                      money(ibc * 0.12), money(ibc * TARIFAS['I']), reg])
    filas.append([("TOTAL", {"b": True}), "", "", (money(IBC_TOTAL), {"b": True}),
                  (money(IBC_TOTAL * 0.085), {"b": True}), (money(IBC_TOTAL * 0.04), {"b": True}),
                  (money(IBC_TOTAL * 0.12), {"b": True}),
                  (money(IBC_TOTAL * TARIFAS['I']), {"b": True}), ""])
    tabla(doc, ["Cód.", "Cargo", "Vinculación", "IBC", "Salud 8,5 %", "Salud 4 %",
                "Pensión 12 %", "ARL 0,522 %", "Régimen"],
          filas, [1.3, 3.4, 2.2, 1.7, 1.7, 1.4, 1.5, 1.4, 1.4], size=7.4, header_size=7.0,
          titulo="Tabla 3. Nómina, IBC y aportes mensuales calculados (2026)",
          align_centro=[3, 4, 5, 6, 7],
          fuente="Fuente: elaboración propia. Calculado sobre el IBC de cada trabajador con el SMMLV "
                 "de 2026 (%s). La columna de salud patronal refleja el escenario sin exoneración, "
                 "es decir, de máxima exposición." % money(SMLMV))
    rich(doc, ["El IBC total de la planta asciende a %s mensuales. Los aportes que la empresa debe "
               "trasladar por salud patronal, salud del trabajador, pensión y riesgos laborales suman "
               "%s al mes, sin incluir parafiscales. De la proyección se advierte que:"
               % (money(IBC_TOTAL), money(IBC_TOTAL * (0.085 + 0.04 + 0.12 + TARIFAS['I'])))])
    bullet(doc, ["%d de los %d trabajadores devengan un salario igual o inferior a dos salarios "
                 "mínimos (%s) y, por tanto, reciben auxilio de transporte por %s, que no integra el "
                 "IBC." % (N_AUX, len(NOMINA), money(TOPE_AUX), money(AUX_TRANSPORTE))])
    bullet(doc, ["Los trabajadores EXT-007 y EXT-009 devengan exactamente un salario mínimo; su base "
                 "de cotización no puede ser inferior a ese valor aunque su salario permanente sea "
                 "inferior por jornada parcial."])
    bullet(doc, ["El aporte a riesgos laborales se calculó con la tarifa de la clase I (0,522 %), que "
                 "es la que corresponde a la actividad según el Decreto 768 de 2022, y no con la "
                 "clase II que la empresa venía aplicando (ver Eje II)."])

    # ---------- 1.4
    h(doc, 2, "Sistema de salud: afiliación, cobertura familiar y prestaciones económicas")
    h(doc, 3, "Verificación de la afiliación al Régimen Contributivo")
    tabla(doc, ["Punto de control", "Estándar normativo", "Hallazgo"],
          [["Afiliación de todos los trabajadores", "Art. 161, num. 1, Ley 100 de 1993",
            "Cumple: 10 de 10 trabajadores afiliados a EPS"],
           ["Libertad de elección de EPS", "Art. 161, num. 1, Ley 100 de 1993",
            "Cumple: la afiliación se hizo por elección del trabajador"],
           ["Reporte de novedades en la planilla integrada", "Decreto 1406 de 1999",
            "Cumple parcialmente: novedades extemporáneas en 2 de 10 casos"],
           ["Pago dentro de los plazos", "Arts. 22 y 23, Ley 100 de 1993",
            "Cumple: no se registran moras en el periodo auditado"],
           ["Cobertura del grupo familiar", "Art. 163, Ley 100 de 1993",
            "Cumple parcialmente: 1 beneficiario mayor de 25 años sin acreditar condición"],
           ["Afiliación del contratista independiente", "Art. 89, Ley 100 de 1993",
            "Cumple: aporta sobre el 40 % de los honorarios"]],
          [4.3, 4.7, 7.0], size=8.6,
          titulo="Tabla 4. Matriz de verificación de la afiliación en salud")

    h(doc, 3, "Cobertura del grupo familiar")
    rich(doc, ["El artículo 163 de la Ley 100 de 1993, modificado por el artículo 218 de la Ley 1753 "
               "de 2015, define el núcleo familiar del cotizante así: (i) el cónyuge o, a falta de "
               "este, el compañero o compañera permanente; (ii) los hijos hasta que cumplan "
               "veinticinco años que dependan económicamente del afiliado; y (iii) los hijos con "
               "incapacidad permanente. La cobertura familiar es una manifestación del principio de "
               "solidaridad y su denegación injustificada vulnera el derecho fundamental a la salud "
               "(art. 49 C.P. y Ley Estatutaria 1751 de 2015)."])
    rich(doc, ["En la auditoría se identificó que un beneficiario mayor de veinticinco años permanece "
               "incluido en el núcleo familiar sin soporte de estudios ni de incapacidad. La "
               "corrección de esa novedad no es un asunto meramente administrativo: si la EPS detecta "
               "el error puede cobrar retroactivamente la Unidad de Pago por Capitación y negar "
               "servicios, y el trabajador queda expuesto a una desprotección sobreviniente. Debe "
               "depurarse el novedoso y conservarse los soportes de dependencia económica."])

    h(doc, 3, "Gestión de las prestaciones económicas")
    rich(doc, ["Las prestaciones económicas del sistema de salud son los pagos en dinero que "
               "sustituyen el ingreso del trabajador cuando no puede trabajar por maternidad, "
               "paternidad o enfermedad general. Su reconocimiento corresponde a la EPS, pero el "
               "empleador cumple un papel esencial: verifica el derecho, adelanta el pago y recobra."])
    tabla(doc, ["Prestación", "Duración", "Valor", "Quién paga", "Norma"],
          [["Licencia de maternidad", "18 semanas (126 días calendario)", "100 % del IBC",
            "EPS (la empresa adelanta y recobra)", "Art. 236 C.S.T., mod. Ley 2114 de 2021"],
           ["Licencia de paternidad", "2 semanas (14 días calendario)", "100 % del IBC",
            "EPS (la empresa adelanta y recobra)", "Art. 236 C.S.T., mod. Ley 2114 de 2021"],
           ["Incapacidad por enfermedad general", "Días 1 y 2", "66,67 % del IBC", "Empleador",
            "Art. 1, Decreto 2943 de 2013"],
           ["Incapacidad por enfermedad general", "Días 3 a 90", "66,67 % del IBC", "EPS",
            "Art. 1, Decreto 2943 de 2013"],
           ["Incapacidad por enfermedad general", "Días 91 a 180", "50 % del IBC", "EPS",
            "Art. 1, Decreto 2943 de 2013"],
           ["Incapacidad por enfermedad general", "Más de 180 días", "50 % del IBC",
            "Fondo de pensiones (con concepto favorable de rehabilitación)",
            "Art. 40, Decreto 1406 de 1999"]],
          [3.7, 3.1, 2.2, 3.6, 3.4], size=8.2,
          titulo="Tabla 5. Prestaciones económicas del sistema de salud (2026)")

    h(doc, 4, "Trámite práctico de una licencia de maternidad")
    for t in [
        "La trabajadora informa el estado de gestación y aporta el certificado médico con la fecha "
        "probable de parto. La empresa verifica que la cotización se haya surtido durante todo el "
        "periodo de gestación como cotizante activa y no como beneficiaria.",
        "Al producirse el parto se aporta el registro civil de nacimiento y el certificado de "
        "licencia; la empresa reporta la novedad en la planilla integrada y liquida la prestación "
        "sobre el IBC del periodo.",
        "La empresa paga la licencia en la nómina y presenta la solicitud de recobro ante la EPS "
        "dentro de los términos reglamentarios, conservando el soporte de radicación.",
        "Si la EPS niega el reconocimiento, la trabajadora conserva la vía de la reclamación ante la "
        "entidad y, agotada esta, la acción judicial. La negativa no puede traducirse en pérdida del "
        "pago: el empleador que adelantó la suma puede repetir contra la administradora.",
    ]:
        bullet(doc, t, nivel=1)
    caja(doc, "Hallazgo del Eje I",
         ["Se identificó una licencia de maternidad (EXT-004) reconocida por la EPS con un pago "
          "inferior al 100 % del IBC durante las primeras semanas, por haberse reportado la novedad "
          "de variación de salario con extemporaneidad. La diferencia asciende a $ 412.000. La "
          "corrección exige solicitud de reliquidación ante la EPS y ajuste de la novedad; si no se "
          "corrige, el empleador queda expuesto a una demanda por la diferencia y la trabajadora a "
          "una disminución injustificada de su ingreso.",
          [("Nivel de riesgo: medio. Responsable: talento humano. Plazo: 30 días.", {"b": True})]],
         tipo='alerta')
    salto_pagina(doc)
