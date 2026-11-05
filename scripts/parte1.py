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
                   EMPRESA, ACTA, TRABAJADORES, EXONERADO_114_1, APORTE_ARL,
                   APORTE_PENSION, APORTE_SALUD_TRABAJADOR, APORTE_CAJA,
                   IBC_TOTAL, N_AUX, CONTRATISTA)

AUTOR = "Julia Inés Buitrago García"
RAZON = "EXTINTORES S C M"
NIT = "39626936-8"
MUNICIPIO = "Soacha (Cundinamarca)"


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
    p(doc, "C.C./NIT 39626936-8 · Empresa privada · Soacha (Cundinamarca)", size=11, color=GRIS,
      align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    p(doc, "Comercio, mantenimiento y recarga de extintores y sistemas de protección contra incendio",
      size=9.5, color=GRIS_CLARO, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    p(doc, "Domicilio: Soacha (Cundinamarca), Colombia", size=9.5, color=GRIS_CLARO,
      align=WD_ALIGN_PARAGRAPH.CENTER, space_after=20)

    tabla(doc, ["Campo", "Información"],
          [["Elaborado por", AUTOR],
           ["Modalidad", "Trabajo académico individual (adaptación del diseño colaborativo "
             "del proyecto integrador)"],
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
          ("desarrollo jurídico ya redactado", {"b": True}),
          ": marco normativo, análisis de casos, liquidaciones, escritos procesales y "
          "conclusiones, con la norma y la jurisprudencia citadas en cada punto. El segundo son "
          "los ",
          ("campos por diligenciar", {"b": True}),
          ": los anexos del capítulo final y los espacios marcados con líneas de puntos. Ahí van "
          "los soportes reales de EXTINTORES S C M: planillas PILA, comprobantes de nómina, "
          "certificados de afiliación, historia laboral y documentos del SG-SST.",
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
        "Auditar la seguridad social de una empresa de tres personas tiene algo de trampa. Uno "
        "espera encontrar el incumplimiento grande, el que sale en los periódicos, y termina "
        "encontrando otra cosa: tres trabajadores afiliados, la planilla al día y, al mismo tiempo, "
        "un requerimiento del año 2024 que nadie respondió. ",
        (RAZON, {"b": True}),
        " es una empresa unipersonal de ",
        (AUTOR, {"b": True}),
        ", con cédula/NIT ",
        (NIT, {"b": True}),
        ", domiciliada en Soacha y dedicada al comercio, la recarga y el mantenimiento de extintores. "
        "La auditoría siguió los cuatro ejes de la guía del curso y se apoyó en lo que había: el "
        "Registro Único Tributario, las planillas, los comprobantes de nómina y un documento que "
        "cambió el trabajo, el acta 028-050164-26 del 9 de septiembre de 2026.", ])

    rich(doc, ["Esa acta es la que manda. No la pedí yo: la levantó la Secretaría de Salud de Soacha "
               "en una visita de inspección y quedó firmada por una ingeniera química y por el "
               "vendedor de la empresa. En la revisión documental escribieron, con esas palabras, "
               "que «no se evidencia documento de sistemas de seguridad y salud en el trabajo». "
               "Uno puede discutir una auditoría interna; contra eso no hay mucho que alegar."])

    rich(doc, ["La consecuencia está en plata. Reparar un accidente de trabajo por la vía tarifada "
               "del sistema cuesta $ 9.044.000. Repararlo por la vía plena del artículo 216 del "
               "C.S.T., con la liquidación que armo en el capítulo 2, cuesta $ 27.284.000. Los "
               "$ 18.240.000 de diferencia salen del bolsillo de la titular, y aquí conviene "
               "detenerse: en una empresa unipersonal el patrimonio de la empresa y el de la persona "
               "son la misma cosa. No hay sociedad que separe el riesgo."])

    rich(doc, ["Y sin embargo, de eso no se sigue que la empresa esté en crisis. Hay tres cosas que "
               "están bien y conviene decirlas: la afiliación al régimen contributivo está completa, "
               "la cotización a pensión no tiene mora y la clasificación de riesgo coincide con el "
               "CIIU registrado. El problema es más aburrido y más difícil de arreglar, porque no se "
               "resuelve pagando: es la costumbre de dejar para después lo que ya se sabe exigible."])

    p(doc, "Lo que encontré, eje por eje", size=11.5, bold=True, color=AZUL, space_before=8,
      space_after=4)
    for titulo, texto in [
        ("Eje I · Afiliación y cotizaciones (riesgo medio). ",
         "Tres trabajadores, todos afiliados. Aquí aparece algo que suele leerse al revés: el "
         "artículo 114-1 del Estatuto Tributario excluye de la exoneración a las personas naturales "
         "que emplean menos de dos trabajadores, y esta empresa tiene tres. Por eso sí aplica, y el "
         "ahorro es de $ 1.100.372 al mes entre salud patronal, SENA e ICBF. El margen es de uno: si "
         "la planta bajara a una sola persona, habría que reliquidar con intereses. El otro hallazgo "
         "es una licencia de paternidad liquidada sobre una base incompleta: $ 412.000 de diferencia "
         "que se corrigen en una tarde y que, sin corregir, se convierten en una demanda pequeña."),
        ("Eje II · Riesgos laborales (riesgo alto). ",
         "El CIIU 4759 ubica a la empresa en clase I, tarifa del 0,522 %, y la tabla del Decreto 768 "
         "de 2022 incluye expresamente el comercio de extintores. O sea que la clasificación está "
         "bien puesta. El problema empieza cuando uno entra al local y ve que allí también se "
         "recargan y se prueban cilindros a presión, que no es comercio sino servicio técnico sobre "
         "equipos a presión. Esa discusión no se ha planteado a la ARL, y mientras no se plantee, la "
         "empresa cotiza por un riesgo que no es el suyo. Sumado a la ausencia del SG-SST, el caso "
         "AT-2026-014 deja al descubierto seis de siete deberes incumplidos: ahí ya no se discute "
         "tarifa, se discute culpa patronal."),
        ("Eje III · Pensiones (riesgo alto). ",
         "Dos de los tres trabajadores están en prima media; el tercero, en ahorro individual. El "
         "caso P-03 no se sostiene en la conveniencia de un régimen sobre otro, sino en algo más "
         "simple y más grave: no hay acta de doble asesoría. Con la SU-107 de 2024 la carga de la "
         "prueba se traslada a la administradora, y si no la acredita, el traslado es ineficaz. La "
         "diferencia de mesada que calculo es de $ 345.095 mensuales, $ 82.822.800 en veinte años. "
         "Que esa cifra sea suficiente para compensar el daño, ya es otra pregunta."),
        ("Eje IV · Insumo procesal (riesgo bajo). ",
         "La calidad del demandado decide el camino: reclamación administrativa previa si es entidad "
         "pública, derecho de petición y agotamiento de vía interna si es empresa privada. La "
         "competencia se fija por materia, territorio y cuantía, y aquí el detalle importa: con "
         "domicilio en Soacha, el circuito competente es el de Soacha, no el de Bogotá. Van tres "
         "documentos modelo con los espacios en blanco para radicar."),
    ]:
        par = doc.add_paragraph()
        par.paragraph_format.space_after = Pt(7)
        par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r1 = par.add_run(titulo)
        r1.bold = True
        r1.font.size = Pt(11)
        r2 = par.add_run(texto)
        r2.font.size = Pt(11)

    caja(doc, "Conclusión general",
         ["El cumplimiento formal es aceptable, pero existen tres brechas materiales: (i) la clase "
          "de riesgo registrada no cubre la actividad de recarga que se ejecuta en el local, y esa "
          "discusión no se ha planteado a la ARL; "
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
    rich(doc, ["Esto es un estudio de caso con método jurídico-dogmático. Dicho sin adorno: se toma "
               "una empresa concreta, se mira qué normas le aplican y se compara lo que la norma "
               "exige con lo que la empresa efectivamente hizo. Hay una parte descriptiva, que es "
               "la mayor, y una parte propositiva, que aparece al final de cada eje y en el plan de "
               "mejoramiento."])
    rich(doc, ["La técnica principal fue revisión documental. Por el lado del derecho, la "
               "Constitución, las leyes, los decretos y la jurisprudencia. Por el lado de la "
               "empresa, lo que existía: el Registro Único Tributario, las planillas integradas, los "
               "comprobantes de nómina, los certificados de afiliación, el acta de inspección "
               "sanitaria y lo poco que había de SG-SST. El método es el de siempre en esta "
               "materia: se contrasta el hecho con la hipótesis legal y, cuando la norma no resuelve, "
               "se acude a la ",
               ("ratio decidendi", {"i": True}),
               " de las providencias citadas."])
    rich(doc, ["Una precisión sobre el alcance, porque es la objeción que un jurado haría primero. "
               "El informe no revisa todas las obligaciones del sistema: se limita a los cuatro ejes "
               "de la guía. Los hallazgos valen para esos contornos y no para afirmar que la empresa "
               "esté incumpliendo en todo lo demás. Hay materias que quedaron por fuera y que "
               "merecerían su propio trabajo: cesantías, prima de servicios, vacaciones y el "
               "cumplimiento pleno de la Resolución 0312 de 2019, por ejemplo."])

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
    rich(doc, ["Los hallazgos se clasifican en tres niveles y el criterio es de evidencia, no de "
               "gravedad. ",
               ("Riesgo bajo", {"b": True}),
               ": la obligación se cumple y está documentada. ",
               ("Riesgo medio", {"b": True}),
               ": se cumple, pero sin soporte suficiente o con errores de liquidación. ",
               ("Riesgo alto", {"b": True}),
               ": hay incumplimiento o exposición a responsabilidad patronal. La escala se explica "
               "en cada eje para que el lector pueda discutirla."])
    rich(doc, ["Conviene advertir una cosa sobre esa escala. Califica evidencia, no intenciones. Una "
               "empresa puede incumplir por descuido y otra por cálculo, y el riesgo jurídico es el "
               "mismo en los dos casos. La diferencia solo importa cuando se discute la culpa, porque "
               "el artículo 216 del C.S.T. no pide mala fe: pide diligencia."])
    rich(doc, ["Y ahí el expediente de EXTINTORES S C M tiene un dato incómodo. El acta "
               "028-050164-26 consigna que la visita anterior, del 18 de octubre de 2024, había "
               "calificado el establecimiento con 57,50 % de cumplimiento y concepto «favorable con "
               "requerimientos». Es decir: lo que en 2024 era un hallazgo, en 2026 seguía siendo un "
               "hallazgo. Eso no convierte a la titular en una empleadora temeraria, pero mueve la "
               "discusión del terreno del desconocimiento al de la persistencia, que es donde la "
               "culpa se prueba con más facilidad."])
    rich(doc, ["Queda una advertencia sobre las cifras, y prefiero decirla aquí y no en una nota al "
               "pie. Las bases de cotización marcadas con asterisco son valores de referencia: la "
               "empresa todavía no aportó los comprobantes de cada trabajador. Los parámetros de los "
               "casos AT-2026-014 y P-03 (edad, semanas, capital acumulado, días de incapacidad) "
               "también son de trabajo. Todo lo demás viene de documentos: el acta, el RUT, la "
               "planilla y los decretos citados. Donde hay un supuesto, el documento lo dice."])
    salto_pagina(doc)


# ---------------------------------------------------------------- Introducción
def introduccion(doc):
    h_sin_num(doc, 1, "INTRODUCCIÓN Y PLANTEAMIENTO DEL PROBLEMA")
    rich(doc, ["La Ley 100 de 1993 no se conformó con obligar a los empleadores a pagar aportes. Les "
               "pidió además afiliar, reportar novedades con veracidad, tramitar prestaciones, "
               "administrar el riesgo laboral y explicar la decisión pensional de cada trabajador. "
               "Son deberes distintos, con sanciones distintas, y el error más común consiste en "
               "creer que cumplir el primero cubre los otros cuatro."])
    rich(doc, ["Las consecuencias de incumplirlos van de menor a mayor. En el extremo bajo están la "
               "sanción moratoria y el cobro de las administradoras (artículos 23 y 24 de la Ley 100 "
               "de 1993). En el medio, la ineficacia de actos jurídicos, que es lo que ocurre con un "
               "traslado de régimen sin asesoría. En el extremo alto está la responsabilidad "
               "patrimonial plena del artículo 216 del Código Sustantivo del Trabajo, que no tiene "
               "tope y que en una empresa unipersonal se cobra contra el patrimonio personal del "
               "dueño."])
    rich(doc, ["En la microempresa esas brechas se ven con más claridad porque no hay dónde "
               "esconderlas. La gestión de la seguridad social suele quedar en una sola persona, que "
               "además es la dueña, la que vende y la que firma; se confunde afiliación con pago; y "
               "la prueba documental se construye después, cuando ya hay un problema. ",
               RAZON, " responde a ese patrón y le agrega dos particularidades. La primera es el "
               "riesgo: manipular cilindros presurizados, recargar con agentes extintores y hacer "
               "pruebas hidrostáticas no es una actividad de mostrador. La segunda es la forma "
               "jurídica: al ser persona natural empleadora, el patrimonio de la empresa y el de la "
               "titular son el mismo, y eso cambia el cálculo de cualquier indemnización."])
    rich(doc, ["El problema que orienta el trabajo se puede formular en una pregunta: ",
               ("¿en qué medida el cumplimiento de las obligaciones de seguridad social de "
                "EXTINTORES S C M satisface el estándar normativo vigente, y qué consecuencias "
                "jurídicas se derivan de las brechas identificadas en salud, riesgos laborales, "
                "pensiones y acceso a la administración de justicia?", {"i": True})])
    rich(doc, ["La pregunta se responde en cinco capítulos. El primero caracteriza la organización y "
               "revisa afiliación, IBC, cobertura y prestaciones. El segundo entra al Sistema General "
               "de Riesgos Laborales con un caso de accidente de trabajo y el análisis de la culpa "
               "patronal. El tercero hace el diagnóstico pensional y examina la ineficacia del "
               "traslado. El cuarto reúne los insumos procesales y fija la competencia. El quinto "
               "ordena los hallazgos, propone el plan de mejoramiento y prepara la sustentación "
               "oral."])
    rich(doc, ["Una última advertencia sobre el método, que vale más que cualquier nota al pie. En "
               "una empresa de tres trabajadores cada expediente es un tercio de la planta. Un error "
               "en una historia laboral son un trabajador y su pensión comprometida. Por eso el "
               "informe se detiene en detalles que en una empresa grande "
               "pasarían inadvertidos, y por eso las recomendaciones son concretas y con fecha."])
    salto_pagina(doc)


# ================================================================ EJE I
def eje_i(doc):
    salto_pagina(doc)
    h(doc, 1, "EJE I · CARACTERIZACIÓN Y RÉGIMEN DE AFILIACIÓN (SALUD Y COTIZACIONES)")

    # ---------- 1.1
    h(doc, 2, "Ficha técnica de la organización")
    rich(doc, ["Antes de auditar hay que saber a quién se audita. De la naturaleza jurídica del "
               "empleador, de su actividad y de su planta depende qué obligaciones le son exigibles y "
               "por cuánto. La ficha que sigue no la armé con lo que la empresa dice de sí misma: sale "
               "del Registro Único Tributario y del acta de inspección sanitaria, que son documentos "
               "de terceros y por eso valen más."])
    tabla(doc, ["Variable", "Dato", "Relevancia jurídica"],
          [["Razón social", EMPRESA["razon_social"],
            "Identifica al empleador como sujeto de obligaciones (art. 22, Ley 100 de 1993). El acta "
            "028-050164-26 registra el nombre comercial EXTINTORES S C M"],
           ["Documento (cédula/NIT)", EMPRESA["documento"],
            "Se usa en la planilla PILA y en la identificación de los escritos procesales. La "
            "numeración corresponde a la cédula del titular, sin dígito de verificación"],
           ["Naturaleza jurídica", "Empresa privada; persona natural comerciante (empresa unipersonal)",
            "Determina el régimen aplicable: no hay reclamación administrativa previa del art. 6 del "
            "C.P.T.S.S., sino solicitud directa al empleador y, en su caso, conciliación extrajudicial"],
           ["Actividad económica registrada", "CIIU 4759, %s" % EMPRESA["ciiu_desc"],
            "El código CIIU define la clase de riesgo ante la ARL (Decreto 768 de 2022): la "
            "subactividad 4759-01 está clasificada en clase I e incluye expresamente el comercio de "
            "extintores"],
           ["Actividad operativa adicional", "Recarga, mantenimiento y prueba hidrostática de "
            "extintores en el mismo establecimiento",
            "Es una actividad conexa que introduce riesgo de presión, proyección de partes y "
            "exposición química; puede justificar un código adicional o una revisión de la clase "
            "(arts. 4 y 5 del Decreto 768 de 2022)"],
           ["Domicilio", "Carrera 4 n.° 60-17, barrio La Despensa, comuna 3, Soacha (Cundinamarca)",
            "Fija la competencia territorial del juez laboral (art. 11 del C.P.T.S.S.): el circuito "
            "judicial de Soacha, no el de Bogotá"],
           ["Representante legal", AUTOR + ", C.C. " + EMPRESA["cc_titular"],
            "Única titular y responsable de las obligaciones patronales"],
           ["Inscripción mercantil", "Matrícula 25 754 39626936-8; código DANE 25 754",
            "Acredita la existencia y representación del establecimiento ante la autoridad sanitaria"],
           ["Planta de personal", "%d trabajadores dependientes con contrato de trabajo" % len(NOMINA),
            "La condición de persona natural con dos o más trabajadores activa la exoneración del "
            "art. 114-1 del E.T., que sí la ampara en este caso"],
           ["Contratistas", "1 contrato de prestación de servicios (contador público)",
            "Exige verificar la afiliación del contratista como independiente (art. 89, Ley 100 de 1993)"],
           ["Tamaño empresarial", "Microempresa",
            "No altera las obligaciones de afiliación, que son imperativas; sí incide en los "
            "estándares mínimos del SG-SST (Resolución 0312 de 2019)"]],
          [3.3, 6.5, 6.2], size=8.6,
          titulo="Tabla 2. Ficha técnica de EXTINTORES S C M",
          fuente="Fuente: elaboración propia con base en el acta de inspección sanitaria "
                 "028-050164-26 del 09/09/2026 (Secretaría de Salud de Soacha) y el Registro Único "
                 "Tributario.")

    caja(doc, "La forma jurídica importa más de lo que parece",
         ["Trabajar bajo la forma de empresa unipersonal de persona natural tiene dos efectos que "
          "rara vez se tienen presentes. El primero: el patrimonio del titular responde por las "
          "obligaciones de la empresa, porque no existe la separación que sí da una sociedad "
          "comercial. El segundo: la exoneración del artículo 114-1 del Estatuto Tributario exige, "
          "tratándose de personas naturales, emplear dos o más trabajadores. Con tres, la empresa "
          "cumple el umbral, pero el cómputo se hace sobre trabajadores reales y activos en cada "
          "periodo, no sobre cargos que figuran en un papel."],
         tipo='alerta')

    # ---------- 1.2
    h(doc, 2, "Marco normativo de la afiliación y la cotización en salud")
    rich(doc, ["El sistema de salud tiene dos regímenes y solo uno de ellos es el que aquí interesa: "
               "el contributivo, que se paga con cotizaciones de trabajadores y empleadores. El "
               "subsidiado se financia con recursos fiscales y con el Fondo de Solidaridad y "
               "Garantía. Para los trabajadores dependientes la afiliación al contributivo es "
               "obligatoria, y el empleador debe inscribirlos en la EPS que ellos elijan. El numeral "
               "1 del artículo 161 de la Ley 100 de 1993 protege justamente eso: que la afiliación "
               "colectiva no sirva para decidir por el trabajador."])
    rich(doc, ["La Ley Estatutaria 1751 de 2015 cambió el estatus de la salud, que pasó de servicio "
               "público a derecho fundamental autónomo e irrenunciable, y fijó sus elementos "
               "esenciales (disponibilidad, aceptabilidad, accesibilidad, calidad e idoneidad). De "
               "ahí salen los principios de progresividad, integralidad y equidad. Esa ley prohíbe a "
               "las administradoras y a los empleadores negar, retardar o condicionar la afiliación "
               "o el acceso a los "
               "servicios, lo que refuerza el deber patronal de reportar novedades de manera oportuna "
               "y veraz."])

    h(doc, 3, "Obligaciones del empleador en materia de salud")
    rich(doc, ["Son cinco y conviene tenerlas separadas, porque se cumplen en momentos distintos y "
               "se prueban con documentos distintos. Pagar la planilla no acredita haber reportado "
               "la novedad a tiempo, y ese es el error que después cuesta."])
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
    rich(doc, ["La cotización equivale al ",
               ("12,5 % del ingreso base de cotización", {"b": True}),
               ", conforme al artículo 204 de la Ley 100 de 1993 modificado por la Ley 1122 de 2007. "
               "De ese total, 8,5 %% lo pone el empleador y 4 %% el trabajador. Sobre un salario "
               "mínimo de 2026 eso significa %s del empleador y %s del trabajador, todos los meses, "
               "con o sin utilidad de la empresa." % (
                   money(SMLMV * 0.085), money(SMLMV * 0.04))])
    rich(doc, ["Hay una regla que suele olvidarse y que el artículo 22 de la Ley 100 de 1993 deja "
               "clara: el empleador responde por la totalidad del aporte aunque no haya descontado la "
               "parte del trabajador. Si el descuento no se hizo, el problema es del empleador, no del "
               "afiliado. Y si la planilla se paga tarde, el sistema cobra mora con independencia de "
               "quién tenga la culpa."])
    imagen(doc, "assets/f2_aportes.png", 15.4,
           "Figura 1. Distribución de la cotización mensual al SGSSS para un trabajador dependiente. "
           "Fuente: elaboración propia con base en el artículo 204 de la Ley 100 de 1993 y el "
           "artículo 114-1 del Estatuto Tributario.")

    h(doc, 3, "La exoneración del artículo 114-1 del Estatuto Tributario")
    rich(doc, ["El artículo 114-1 del Estatuto Tributario exime del pago de SENA, ICBF y salud "
               "patronal a las sociedades contribuyentes de renta, por los trabajadores que devenguen "
               "menos de diez salarios mínimos. El inciso segundo hace lo mismo con las personas "
               "naturales empleadoras, ",
               ("pero saca del beneficio a quienes empleen menos de dos trabajadores", {"b": True}),
               ". La exoneración no toca pensión ni riesgos laborales, y el 4 % del trabajador se "
               "descuenta igual."])
    rich(doc, ["La regla parece sencilla hasta que se aplica. Una persona natural que contrata a "
               "alguien para que le ayude en el mostrador paga el 8,5 % completo. Si contrata a un "
               "segundo, deja de pagarlo. No hay razón jurídica para el umbral, pero así está "
               "redactada la norma, y el umbral es lo que decide si una microempresa es viable o no."])
    rich(doc, ["En EXTINTORES S C M la empresa ",
               ("puede aplicar la exoneración", {"b": True}),
               " por los tres trabajadores, porque ninguno llega a los diez salarios mínimos (%s en "
               "2026), siempre que sea contribuyente declarante de renta y que la novedad se reporte "
               "bien en la planilla. El margen, insisto, es de uno: si la planta bajara a una sola "
               "persona, la exoneración desaparece y toca reliquidar con intereses todo el periodo en "
               "que se aplicó sin derecho." % money(TOPE_EXO)])
    rich(doc, ["El efecto económico se puede medir sobre el IBC total de %s. La exoneración ahorra "
               "el 8,5 %% de salud patronal (%s), el 2 %% de SENA (%s) y el 3 %% de ICBF (%s): en "
               "conjunto, %s mensuales frente al escenario sin exoneración. Lo que no se exonera "
               "nunca es la pensión del 12 %% (%s), el 4 %% de salud a cargo del trabajador (%s), los "
               "riesgos laborales (%s con la tarifa de clase I) ni la caja de compensación del 4 %% "
               "(%s)."
               % (money(IBC_TOTAL),
                  money(IBC_TOTAL * 0.085), money(IBC_TOTAL * 0.02), money(IBC_TOTAL * 0.03),
                  money(IBC_TOTAL * 0.135),
                  money(IBC_TOTAL * 0.12), money(IBC_TOTAL * 0.04),
                  money(APORTE_ARL), money(IBC_TOTAL * 0.04))])
    tabla(doc, ["Concepto", "Tarifa", "Sin exoneración", "Con exoneración (art. 114-1 E.T.)"],
          [["Salud a cargo del empleador", "8,5 %", money(IBC_TOTAL * 0.085), "$ 0"],
           ["SENA", "2 %", money(IBC_TOTAL * 0.02), "$ 0"],
           ["ICBF", "3 %", money(IBC_TOTAL * 0.03), "$ 0"],
           ["Pensión (empleador 12 %)", "12 %", money(IBC_TOTAL * 0.12),
            money(IBC_TOTAL * 0.12)],
           ["Riesgos laborales (clase I)", "0,522 %", money(APORTE_ARL), money(APORTE_ARL)],
           ["Caja de compensación", "4 %", money(IBC_TOTAL * 0.04), money(IBC_TOTAL * 0.04)],
           [("Total a cargo del empleador", {"b": True}), "",
            (money(IBC_TOTAL * (0.085 + 0.02 + 0.03 + 0.12 + TARIFAS["I"] + 0.04)), {"b": True}),
            (money(IBC_TOTAL * (0.12 + TARIFAS["I"] + 0.04)), {"b": True})],
           [("Ahorro mensual", {"b": True}), "13,5 %", "",
            (money(IBC_TOTAL * 0.135), {"b": True})]],
          [6.0, 1.8, 3.6, 4.2], size=8.5, align_centro=[1, 2, 3],
          titulo="Tabla 3. Efecto de la exoneración del artículo 114-1 del E.T. sobre el IBC de la "
                 "planta",
          fuente="Fuente: elaboración propia sobre un IBC total de %s. El aporte a pensión del "
                 "empleador y el 4 %% de salud del trabajador no se ven afectados por la "
                 "exoneración." % money(IBC_TOTAL))
    caja(doc, "Dónde se cae esto en la práctica",
         ["La exoneración se declara en la planilla con un código. Aplicarla sin derecho deja un "
          "faltante que la ADRES o la EPS cobran con intereses (artículo 23 de la Ley 100 de 1993), y "
          "en un proceso eso aparece como tiempo no cotizado del trabajador, no como un error "
          "contable. Dos cosas para revisar: que el beneficio se aplique solo sobre la base salarial, "
          "nunca sobre pagos no constitutivos de salario, y que el número de trabajadores se cuente "
          "bien en cada periodo."],
         tipo='nota')

    # ---------- 1.3
    h(doc, 2, "Liquidación del Ingreso Base de Cotización (IBC)")
    rich(doc, ["El IBC no es el salario. Es lo que queda del salario después de quitarle lo que la "
               "ley excluye y de aplicarle topes, y sobre esa cifra se calculan salud, pensión y "
               "riesgos laborales. Cuando se liquida mal, el daño va en dos direcciones: al sistema le "
               "falta plata y al trabajador le baja la mesada. Lo segundo casi nunca se nota hasta "
               "que la persona se pensiona, y entonces ya es tarde para corregir sin demanda."])

    h(doc, 3, "Regla general para trabajadores dependientes")
    rich(doc, ["El artículo 127 del Código Sustantivo del Trabajo define el salario por lo que es, no "
               "por cómo se le llame: todo lo que recibe el trabajador en dinero o en especie como "
               "contraprestación directa del servicio. Entran entonces el básico, las horas extras y "
               "recargos, las comisiones, los porcentajes sobre ventas, los viáticos permanentes y las "
               "primas habituales. El artículo 128 deja por fuera el auxilio de transporte, los "
               "viáticos accidentales, la dotación, las propinas y las primas extralegales."])
    rich(doc, ["Ahí hay una trampa frecuente. El auxilio de transporte no integra el IBC, pero sí se "
               "suma para cesantías y prima de servicios. Son dos bases distintas y en la "
               "microempresa se confunden todo el tiempo, casi siempre en perjuicio del trabajador."])
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
    rich(doc, ["El contratista de prestación de servicios no cotiza sobre lo que le pagan. Cotiza "
               "sobre el ",
               ("40 % del valor mensualizado del contrato, sin incluir el IVA", {"b": True}),
               " (artículo 89 de la Ley 100 de 1993 y artículo 23 del Decreto 1703 de 2002). El otro "
               "60 % se entiende destinado a cubrir los costos de la actividad, aunque en la realidad "
               "no siempre alcance. Si el 40 % queda por debajo de un salario mínimo, la base se sube "
               "al mínimo. Y el contratista paga todo: 12,5 % de salud, 16 % de pensión y el "
               "porcentaje de riesgos laborales que corresponda a la actividad."])
    rich(doc, ["Aplicado al contador, que factura %s al mes, el 40 %% son %s. Ese es su IBC, y sobre "
               "él paga %s mensuales entre salud y pensión. Aquí hay un punto que conviene dejar "
               "anotado: si ese contrato encubre una relación laboral, la discusión no es de base de "
               "cotización sino de primacía de la realidad, y entonces la empresa responde por "
               "prestaciones sociales y por aportes completos." % (
                   money(CONTRATISTA['honorarios']), money(CONTRATISTA['honorarios'] * 0.4),
                   money(CONTRATISTA['honorarios'] * 0.4 * 0.285))])

    h(doc, 3, "Aplicación a la nómina de EXTINTORES S C M")
    rich(doc, ["La planta son tres personas y el acta lo confirma: el formato registra ",
               ("un solo trabajador operativo", {"b": True}),
               ", que es el técnico de recarga. Los otros dos son la titular, que administra y "
               "representa, y el vendedor de mostrador. Con esa estructura, cada trabajador pesa un "
               "tercio de la nómina y no hay manera de diluir un error entre muchos."])
    rich(doc, ["Las bases marcadas con asterisco son valores de referencia: los comprobantes de cada "
               "trabajador todavía no están en el expediente. Lo demás sí está verificado."])
    filas = []
    for t in TRABAJADORES:
        marca = " *" if t["por_confirmar"] else ""
        filas.append([t["codigo"], t["nombre"], t["documento"], t["cargo"],
                      money(t["ibc"]) + marca, money(t["ibc"] * 0.04),
                      money(t["ibc"] * 0.12), money(t["ibc"] * TARIFAS['I']), t["regimen"]])
    filas.append([("TOTAL", {"b": True}), "", "", "", (money(IBC_TOTAL), {"b": True}),
                  (money(APORTE_SALUD_TRABAJADOR), {"b": True}),
                  (money(APORTE_PENSION), {"b": True}),
                  (money(APORTE_ARL), {"b": True}), ""])
    tabla(doc, ["Cód.", "Trabajador", "Documento", "Cargo", "IBC *", "Salud 4 %",
                "Pensión 12 %", "ARL 0,522 %", "Régimen"],
          filas, [1.2, 4.0, 2.2, 3.6, 1.7, 1.4, 1.5, 1.4, 1.4], size=7.2, header_size=6.8,
          titulo="Tabla 4. Planta de personal, IBC y aportes mensuales (2026)",
          align_centro=[4, 5, 6, 7],
          fuente="Fuente: elaboración propia con base en el acta 028-050164-26, el Registro Único "
                 "Tributario y los datos aportados por la titular. (*) Base de referencia sujeta a "
                 "confirmación con los comprobantes de nómina. La salud patronal del 8,5 %% no "
                 "aparece porque está exonerada (art. 114-1 del E.T.).")
    rich(doc, ["El IBC sumado de la planta es %s mensuales, y de ahí sale todo lo demás: %s de "
               "pensión, %s de riesgos laborales y %s de caja de compensación. Los %s de salud que "
               "aparecen en la tabla son del trabajador, se descuentan de su salario y no son costo "
               "de la empresa. Con la exoneración, el costo parafiscal real del empleador es %s al "
               "mes. Tres observaciones sobre esa cifra:"
               % (money(IBC_TOTAL), money(APORTE_PENSION), money(APORTE_ARL), money(APORTE_CAJA),
                  money(APORTE_SALUD_TRABAJADOR), money(APORTE_PENSION + APORTE_ARL + APORTE_CAJA))])
    bullet(doc, ["%d de los %d trabajadores devengan un salario igual o inferior a dos salarios "
                 "mínimos (%s) y, por tanto, reciben auxilio de transporte por %s, que no integra el "
                 "IBC. La titular, por su nivel salarial, no lo recibe."
                 % (N_AUX, len(NOMINA), money(TOPE_AUX), money(AUX_TRANSPORTE))])
    bullet(doc, ["El trabajador WT-03 devenga exactamente un salario mínimo (%s); su base de "
                 "cotización no puede ser inferior a ese valor aunque su salario permanente sea "
                 "inferior por jornada parcial." % money(SMLMV)])
    bullet(doc, ["El aporte a riesgos laborales se calculó con la tarifa de la clase I (0,522 %), "
                 "que es la que corresponde al código CIIU 4759 registrado en el RUT según el "
                 "Decreto 768 de 2022. La actividad de recarga que se ejecuta en el mismo "
                 "establecimiento es el punto que puede discutirse ante la ARL (ver Eje II)."])

    # ---------- 1.4
    h(doc, 2, "Sistema de salud: afiliación, cobertura familiar y prestaciones económicas")
    h(doc, 3, "Verificación de la afiliación al Régimen Contributivo")
    tabla(doc, ["Punto de control", "Estándar normativo", "Hallazgo"],
          [["Afiliación de todos los trabajadores", "Art. 161, num. 1, Ley 100 de 1993",
            "Cumple: 3 de 3 trabajadores afiliados a EPS"],
           ["Libertad de elección de EPS", "Art. 161, num. 1, Ley 100 de 1993",
            "Cumple: la afiliación se hizo por elección del trabajador"],
           ["Reporte de novedades en la planilla integrada", "Decreto 1406 de 1999",
            "Cumple parcialmente: novedades extemporáneas en 1 de 3 casos (WT-02)"],
           ["Pago dentro de los plazos", "Arts. 22 y 23, Ley 100 de 1993",
            "Cumple: no se registran moras en el periodo auditado"],
           ["Cobertura del grupo familiar", "Art. 163, Ley 100 de 1993",
            "Cumple: sin beneficiarios por fuera del núcleo legal; falta el soporte documental de cada uno"],
           ["Afiliación del contratista independiente", "Art. 89, Ley 100 de 1993",
            "Cumple: aporta sobre el 40 % de los honorarios"]],
          [4.3, 4.7, 7.0], size=8.6,
          titulo="Tabla 5. Matriz de verificación de la afiliación en salud")

    h(doc, 3, "Cobertura del grupo familiar")
    rich(doc, ["El artículo 163 de la Ley 100 de 1993, con la modificación del artículo 218 de la Ley "
               "1753 de 2015, define el núcleo familiar en tres grupos: el cónyuge o el compañero "
               "permanente; los hijos hasta los veinticinco años que dependan económicamente; y los "
               "hijos con incapacidad permanente. Negar esa cobertura sin razón vulnera el derecho "
               "fundamental a la salud (artículo 49 de la Constitución y Ley Estatutaria 1751 de "
               "2015)."])
    rich(doc, ["En la planta auditada no apareció ningún beneficiario por fuera del núcleo legal, y "
               "eso es una buena noticia modesta: la empresa no arrastra el problema clásico del hijo "
               "de veintiocho años que sigue en la EPS del papá. Pero la verificación hay que "
               "repetirla cada periodo, porque el paso a la mayoría de edad extendida no opera solo. "
               "La EPS puede cobrar retroactivamente la UPC cuando detecta la inclusión indebida, y "
               "entonces el empleador termina pagando un error que nadie revisó."])
    rich(doc, ["El control es de papelería, no de sistema: por cada beneficiario, el registro civil, "
               "el certificado de estudios o la prueba de la dependencia económica, y la novedad de "
               "retiro reportada en el mes en que corresponde. Con tres trabajadores eso cabe en una "
               "carpeta."])

    h(doc, 3, "Gestión de las prestaciones económicas")
    rich(doc, ["Las prestaciones económicas son los pagos que reemplazan el salario cuando el "
               "trabajador no puede trabajar: maternidad, paternidad o enfermedad general. La EPS "
               "reconoce, pero el empleador hace el trabajo sucio: verifica el derecho, paga primero "
               "y después recobra. Cuando esa cadena se rompe, el que queda sin plata es el "
               "trabajador."])
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
          titulo="Tabla 6. Prestaciones económicas del sistema de salud (2026)")

    h(doc, 4, "Trámite práctico de una licencia de paternidad")
    for t in [
        "El trabajador avisa y aporta el registro civil de nacimiento. La empresa revisa que haya "
        "cotizado como cotizante activo durante todo el periodo de gestación y no como beneficiario.",
        "Con el registro civil y el certificado de licencia, la empresa reporta la novedad en la "
        "planilla y liquida la prestación sobre el IBC del periodo.",
        "La empresa paga la licencia en la nómina y radica el recobro ante la EPS dentro del término. "
        "Guarda el soporte de la radicación, que es lo que después prueba que hizo la gestión.",
        "Si la EPS niega, el trabajador reclama y, agotada esa vía, demanda. La negativa no puede "
        "convertirse en pérdida del pago: quien adelantó la suma repite contra la administradora.",
    ]:
        bullet(doc, t, nivel=1)
    caja(doc, "Los dos hallazgos del Eje I",
         ["El primero tiene nombre y cédula: la licencia de paternidad de Felipe Martínez se liquidó "
          "sobre una base inferior al 100 % del IBC porque la variación de salario se reportó tarde. "
          "Son $ 412.000, que se corrigen pidiendo reliquidación a la EPS y ajustando la novedad. El "
          "segundo es más incómodo porque no se arregla con una solicitud: el acta 028-050164-26 "
          "dejó en blanco los campos de empleados afiliados a riesgos laborales y al sistema de "
          "seguridad social. La autoridad sanitaria preguntó y no obtuvo respuesta. Eso hay que "
          "remitirlo a la Secretaría de Salud dentro del plazo.",
          [("Nivel de riesgo: medio. Responsable: talento humano. Plazo: 30 días.", {"b": True})]],
         tipo='alerta')
    salto_pagina(doc)
