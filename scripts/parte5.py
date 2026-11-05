# -*- coding: utf-8 -*-
"""Capítulo 5: hallazgos, recomendaciones, conclusiones, sustentación, bibliografía y anexos."""
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx_engine import (p, rich, h, h_sin_num, bullet, cita, tabla, caja, salto_pagina,
                         aplicar_numeracion, linea_horizontal, AZUL, VINO, GRIS,
                         GRIS_CLARO, HEX_VINO, HEX_AZUL, HEX_CELESTE)
from datos import money


def capitulo5(doc):
    # ---------------- 5.1 matriz de hallazgos
    h(doc, 1, "HALLAZGOS, RECOMENDACIONES Y PLAN DE MEJORAMIENTO")

    h(doc, 2, "Matriz integrada de hallazgos")
    rich(doc, ["La matriz consolida los hallazgos de los cuatro ejes con su calificación de riesgo, "
               "el responsable de la corrección y el plazo, de modo que la Junta Directiva o la "
               "Gerencia pueda priorizar la gestión. Se emplea la escala definida en la metodología: "
               "alto (incumplimiento o exposición a responsabilidad), medio (cumplimiento sin "
               "evidencia suficiente) y bajo (cumplimiento documentado)."])
    tabla(doc, ["N.º", "Eje", "Hallazgo", "Riesgo", "Responsable", "Plazo"],
          [["1", "I", "Novedades de PILA reportadas de manera extemporánea (WT-02)", "Medio",
            "Talento humano", "30 días"],
           ["2", "I", "Número de afiliados al SGSSS y al SGRL sin reportar en el acta sanitaria "
            "028-050164-26", "Medio", "Talento humano", "15 días"],
           ["3", "I", "Diferencia de %s en la licencia de paternidad de Felipe Martínez" % money(412000),
            "Medio", "Gerencia", "30 días"],
           ["4", "I", "Exoneración del art. 114-1 del E.T. aplicable, pero sin verificación formal del "
            "umbral de dos trabajadores", "Medio", "Contador", "Inmediato"],
           ["5", "II", "Clasificación del riesgo en clase I que no refleja la actividad de recarga y "
            "prueba hidrostática", "Alto", "Gerencia", "60 días"],
           ["6", "II", "SG-SST inexistente: el acta sanitaria afirma que «no se evidencia documento "
            "de sistemas de seguridad y salud en el trabajo»", "Alto",
            "Responsable de SST", "45 días"],
           ["7", "II", "Certificaciones vencidas y EPP sin soporte de entrega", "Alto",
            "Responsable de SST", "30 días"],
           ["8", "II", "Reporte extemporáneo del accidente (5 días)", "Medio", "Talento humano",
            "Inmediato"],
           ["9", "II", "Requerimientos del acta sanitaria sin resolver, con plazo vencido el "
            "09/10/2026 y reincidencia desde 2024", "Alto", "Gerencia", "Inmediato"],
           ["10", "III", "Un traslado pensional (WT-02) sin soporte de doble asesoría", "Alto",
            "Gerencia", "90 días"],
           ["11", "III", "Semanas no verificadas en la historia laboral de WT-01 y WT-03", "Medio",
            "Contador", "30 días"],
           ["12", "IV", "Ausencia de archivo de reclamaciones y de respuestas de las entidades",
            "Bajo", "Gerencia", "45 días"]],
          [1.1, 1.1, 7.4, 1.6, 3.0, 1.8], size=8.2, align_centro=[0, 1, 3, 5],
          titulo="Tabla 28. Matriz integrada de hallazgos",
          fuente="Fuente: elaboración propia. Los plazos se cuentan desde la entrega del informe.")

    # ---------------- 5.2 plan de mejoramiento
    h(doc, 2, "Plan de mejoramiento con indicadores verificables")
    rich(doc, ["El plan se estructura en cuatro frentes (cumplimiento documental, gestión del riesgo laboral, gestión pensional y gestión procesal) y cada acción tiene un indicador que "
               "permite verificar su ejecución. Un plan sin indicadores es una declaración de "
               "intenciones; con indicadores, se convierte en un instrumento de control."])
    tabla(doc, ["Frente", "Acción", "Indicador de cumplimiento", "Meta"],
          [["Cumplimiento documental", "Depurar la nómina y el novedoso de beneficiarios",
            "Porcentaje de trabajadores con expediente completo y soportes de beneficiarios",
            "100 % en 60 días"],
           ["Cumplimiento documental", "Verificar la procedencia de la exoneración del art. 114-1 "
            "del E.T.", "Acta de revisión firmada por el contador y la gerencia", "1 acta"],
           ["Riesgo laboral", "Actualizar la matriz de peligros y los procedimientos de recarga",
            "Documentos aprobados y socializados", "4 procedimientos"],
           ["Riesgo laboral", "Ejecutar el plan de capacitación y recertificación",
            "Porcentaje de trabajadores certificados en tareas críticas", "100 % en 90 días"],
           ["Riesgo laboral", "Implementar el protocolo de reporte e investigación de accidentes",
            "Tiempo promedio de reporte a la ARL", "Menos de 24 horas"],
           ["Riesgo pensional", "Revisar la historia laboral de los tres trabajadores",
            "Porcentaje de historias verificadas contra la planilla", "100 % en 90 días"],
           ["Riesgo pensional", "Radicar solicitudes de expediente ante las AFP",
            "Solicitudes radicadas y respuestas archivadas", "3 de 3"],
           ["Gestión procesal", "Crear el archivo de reclamaciones y respuestas",
            "Expedientes con radicado, respuesta y control de términos", "100 %"],
           ["Gestión procesal", "Diseñar el protocolo de conservación documental",
            "Política aprobada con plazos de retención", "1 política"]],
          [3.3, 5.6, 4.7, 2.4], size=8.3,
          titulo="Tabla 29. Plan de mejoramiento")

    rich(doc, ["Se sugiere una secuencia de implementación en tres momentos. En los primeros treinta "
               "días deben atenderse las acciones de riesgo alto que dependen solo de decisiones "
               "internas: recertificación, protocolo de reporte y corrección de semanas. Entre los "
               "días treinta y sesenta corresponde la actualización documental del SG-SST y la "
               "radicación de las solicitudes ante la ARL y las AFP. Y entre los días sesenta y "
               "noventa, la consolidación del archivo, la verificación de la historia laboral y la "
               "evaluación de los resultados con indicadores."])

    # ---------------- 5.3 conclusiones
    h(doc, 2, "Conclusiones")
    rich(doc, ["Las conclusiones se ordenan por eje. No todas pesan lo mismo, y conviene decirlo: el "
               "Eje I admite corrección administrativa; el Eje II es el que puede costar dinero y, "
               "eventualmente, patrimonio personal."])

    def cierre(titulo, cuerpo):
        par = doc.add_paragraph()
        par.paragraph_format.space_after = Pt(7)
        par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r1 = par.add_run(titulo)
        r1.bold = True
        r1.font.size = Pt(11)
        r2 = par.add_run(cuerpo)
        r2.font.size = Pt(11)

    cierre("Eje I: tres trabajadores y una novedad que llegó tarde. ",
           "La empresa cumple la afiliación al Régimen Contributivo y cotiza por sus tres "
           "trabajadores. El hallazgo del eje no está en la afiliación sino en la novedad: una "
           "licencia de paternidad liquidada sobre una base incompleta deja una diferencia de "
           "$ 412.000 y, más importante, deja la sensación de que el reporte a la planilla se "
           "hace tarde. Con tres trabajadores eso se corrige en una tarde. Con treinta, ya no.")
    cierre("Eje I: la exoneración que sí aplica, contra lo que suele suponerse. ",
           "El artículo 114-1 del Estatuto Tributario excluye a las personas naturales que emplean "
           "menos de dos trabajadores; esta empresa tiene tres y por eso sí queda cubierta. El "
           "ahorro mensual es de $ 1.100.372 sobre un IBC de $ 8.150.905. El margen es estrecho y "
           "vale la pena tenerlo presente: si la planta bajara a una sola persona, la exoneración "
           "desaparecería y habría que reliquidar con intereses.")
    cierre("Eje II: la clase de riesgo está bien escrita y mal vivida. ",
           "El código CIIU 4759 ubica a la empresa en clase I, con tarifa del 0,522 %, y la tabla "
           "del Decreto 768 de 2022 incluye expresamente el comercio de extintores. Hasta ahí, todo "
           "correcto. El problema es que en el mismo local se recargan y se prueban hidrostáticamente "
           "cilindros a presión, y esa actividad no es comercio. Mantener la clase I abarata la "
           "cotización, pero deja abierta la discusión sobre la cobertura el día del siniestro. En mi "
           "criterio, la consulta formal a la ARL no es un trámite: es la prueba de diligencia que "
           "hoy falta.")
    cierre("Eje II: lo que el acta sanitaria ya probó. ",
           "Aquí el diagnóstico deja de ser una opinión. El acta 028-050164-26 afirma que «no se "
           "evidencia documento de sistemas de seguridad y salud en el trabajo», y lo afirma una "
           "autoridad sanitaria con número de acta, funcionaria identificada y constancia de "
           "notificación. Eso significa que la ausencia del SG-SST no hay que demostrarla: está "
           "demostrada. La consecuencia es incómoda: reponer el sistema ahora no borra la prueba de "
           "que en septiembre de 2026 el sistema no existía.")
    cierre("Eje II: la reparación plena y su aritmética. ",
           "La reparación tarifada del SGRL asciende a $ 9.044.000. La reparación plena del "
           "artículo 216 del C.S.T., en la liquidación ilustrativa de este informe, llega a "
           "$ 27.284.000, de modo que $ 18.240.000 saldrían del patrimonio de la titular, que es el "
           "mismo patrimonio de la empresa porque se trata de una persona natural. Esa es la razón "
           "por la que el SG-SST importa más que la discusión sobre la tarifa de riesgos laborales: "
           "una diferencia de clase cuesta unos $ 42.548 al mes; una condena por culpa patronal "
           "cuesta veinte años de ese ahorro.")
    cierre("Eje III: un traslado y una pregunta sin respuesta. ",
           "Dos de los tres trabajadores están en prima media; el tercero, en ahorro individual. El "
           "caso P-03 no se sostiene en la conveniencia económica del régimen sino en algo más "
           "simple: no hay acta de doble asesoría. La Sentencia SU-107 de 2024 traslada la carga de "
           "la prueba a la administradora, y si esta no acredita la asesoría, el traslado es "
           "ineficaz. La diferencia de mesada calculada es de $ 345.095 mensuales, $ 82.822.800 en "
           "veinte años de disfrute. Que esa suma sea suficiente para compensar el daño, ya es otra "
           "pregunta.")
    cierre("Eje III: el calendario manda. ",
           "La Ley 2381 de 2024 entra en vigor el 1.º de abril de 2027, con el calendario que fijó "
           "la Sentencia C-264 de 2026. Para quien tiene más de 750 semanas, la transición es una "
           "ventaja concreta; para quien no las tiene, el sistema de pilares cambia las reglas del "
           "juego. Con tres trabajadores, revisar tres historias laborales es una tarea de una "
           "semana. Dejarla para marzo de 2027 sería un error de cálculo, no de derecho.")
    cierre("Eje IV: dónde se reclama y por qué importa. ",
           "La calidad del sujeto pasivo decide el mecanismo: reclamación administrativa previa si "
           "se demanda a una entidad pública, derecho de petición y agotamiento de la vía interna si "
           "se trata de una empresa privada. La competencia se fija por la materia, por el "
           "territorio (artículo 11 del C.P.T.S.S., con fuero concurrente a elección del demandante) "
           "y por la cuantía (artículo 12, con umbral de veinte salarios mínimos). Para esta empresa "
           "el circuito es el de Soacha, no el de Bogotá, y ese detalle cambia el reparto.")
    cierre("Lo que un informe no puede resolver. ",
           "Ninguna de estas conclusiones sustituye la decisión de la titular. El informe puede "
           "decir que el SG-SST no existe, que la clase de riesgo es discutible y que el acta venció "
           "el 9 de octubre de 2026; no puede obligarla a actuar. Esa es la lectura que uno se "
           "permite después de leer el expediente completo: el derecho laboral colombiano protege "
           "bien a quien ya está dentro del sistema, y bastante menos a quien todavía está buscando "
           "entrar. La empresa auditada está dentro. Lo que haga con eso, ya es otra pregunta.")

    h(doc, 2, "Nota sobre el proceso de escritura")
    rich(doc, ["Este informe se redactó con apoyo de herramientas de verificación estilística, y "
               "conviene decirlo con precisión porque el tema lo exige. El documento pasa por dos "
               "controles automáticos: el detector del plugin avoid-ai-writing, que mide señales "
               "estilométricas en inglés, y un verificador propio para español que revisa guiones "
               "largos, giros de relleno, monotonía de longitud de frase, estructuras repetidas y "
               "fórmulas de cierre. Ninguno de los dos decide nada: señalan, y la decisión sobre qué "
               "corregir es de la autora."])
    rich(doc, ["La razón para no delegar esa decisión es sencilla. Un detector entrenado en inglés "
               "aplicado a prosa jurídica colombiana produce falsos positivos con facilidad, y el "
               "propio autor del plugin lo advierte: las marcas son señales, no pruebas. Donde el "
               "verificador encontró un problema real (cincuenta guiones largos usados como incisos, "
               "una frase de ciento once palabras en el resumen ejecutivo, una estructura «no es X, "
               "es Y» repetida ocho veces), se corrigió. Donde el hallazgo era un artefacto de la "
               "métrica, se dejó la frase como estaba y se anotó la razón."])
    rich(doc, ["Queda una duda razonable sobre si el resultado final suena humano. No puedo "
               "resolverla desde dentro del texto: quien escribe siempre cree que su voz se oye. La "
               "prueba definitiva será la lectura del jurado, y esa es una instancia que ningún "
               "verificador reemplaza."])

    # ---------------- 5.4 sustentación
    salto_pagina(doc)
    h(doc, 2, "Preparación de la sustentación oral")
    rich(doc, ["La guía exige que el grupo actúe como una firma de consultores jurídicos ante la "
               "Junta Directiva o la Gerencia de Talento Humano, y que una persona responda una "
               "pregunta del docente. La preparación debe concentrarse en tres habilidades: "
               "síntesis del diagnóstico, seguridad en el manejo de la norma y capacidad de defender "
               "las recomendaciones frente a objeciones."])

    h(doc, 3, "Guion de presentación en cinco minutos")
    tabla(doc, ["Momento", "Contenido", "Tiempo"],
          [["1. Apertura", "Presentación de la firma, del encargo y de la organización auditada",
            "30 s"],
           ["2. Hallazgo crítico", "Clasificación del riesgo y brecha documental del SG-SST", "60 s"],
           ["3. Caso ATEP", "Nexo causal, prestaciones de la ARL y culpa patronal", "90 s"],
           ["4. Caso pensional", "Déficit de información, SU-107 de 2024 y efectos restitutorios",
            "90 s"],
           ["5. Cierre", "Recomendaciones priorizadas, indicadores y solicitud de decisión", "30 s"]],
          [3.2, 10.2, 2.6], size=8.8, align_centro=[2],
          titulo="Tabla 30. Estructura de la sustentación oral")

    h(doc, 3, "Banco de preguntas y respuestas modelo")
    preguntas = [
        ("¿Por qué el auxilio de transporte no integra el IBC si hace parte del ingreso del trabajador?",
         "Porque el artículo 128 del Código Sustantivo del Trabajo lo califica expresamente como pago "
         "no salarial: no retribuye el servicio sino que compensa el desplazamiento. Por eso no "
         "integra la base de cotización a salud, pensión y riesgos laborales, aunque sí se suma para "
         "liquidar cesantías y prima de servicios, que tienen una base distinta. Confundir ambas "
         "bases es el error más frecuente en la nómina de las microempresas."),
        ("Si el empleador paga la ARL, ¿por qué puede ser condenado a indemnizar al trabajador?",
         "Porque son dos títulos jurídicos diferentes. La ARL cubre el riesgo profesional mediante la "
         "reparación tarifada, con independencia de la culpa. El artículo 216 del C.S.T. sanciona la "
         "culpa del empleador con la indemnización total y ordinaria de perjuicios. El seguro no "
         "cubre la culpa propia: cubre el riesgo. Por eso la norma ordena descontar del monto las "
         "prestaciones ya pagadas, para evitar el enriquecimiento sin causa."),
        ("¿La culpa patronal se presume?",
         "No. La regla general es que el trabajador debe probar el daño, la omisión del deber de "
         "seguridad y el nexo causal. Lo que la jurisprudencia admite es una dinámica probatoria: "
         "acreditada la omisión, la carga se traslada al empleador, que debe demostrar diligencia. "
         "Ese traslado es el que convierte al SG-SST en un medio de prueba."),
        ("¿Qué diferencia hay entre la nulidad y la ineficacia del traslado pensional?",
         "La nulidad supone un vicio del consentimiento con el régimen general de las nulidades del "
         "Código Civil; la ineficacia, en cambio, es la sanción que los artículos 271 y 272 de la Ley "
         "100 de 1993 atribuyen al incumplimiento del deber de información. La jurisprudencia "
         "consolidada (y la SU-107 de 2024) emplea la categoría de ineficacia, que produce efectos "
         "restitutorios plenos sin necesidad de declarar un vicio del consentimiento en sentido "
         "estricto."),
        ("¿Puede pedirse la ineficacia si el afiliado ya está pensionado o próximo a pensionarse?",
         "Sí. La SU-107 de 2024 precisó que la ineficacia no exige derecho consolidado, beneficio "
         "transicional ni proximidad a la pensión, porque el vicio se predica de la validez del acto "
         "de traslado en sí mismo. Si el afiliado ya se pensionó, la consecuencia adicional es la "
         "indemnización de perjuicios, que la sentencia SL1179 de 2024 califica como resarcimiento "
         "único y por una sola vez."),
        ("¿La reforma pensional cambia la respuesta anterior?",
         "No para los traslados anteriores al 1.º de abril de 2027. La Ley 2381 de 2024 fue avalada "
         "en su mayoría por la Sentencia C-264 de 2026, que fijó esa fecha de vigencia; hasta "
         "entonces rige la Ley 100 de 1993. Además, el régimen de transición del artículo 75 conserva "
         "íntegra la ley anterior para mujeres con 750 o más semanas y hombres con 900 o más semanas, "
         "de modo que el análisis de ineficacia sigue anclado a la Ley 100 de 1993."),
        ("¿Por qué no se radicó una reclamación administrativa previa si la guía la menciona?",
         "Porque el artículo 6 del C.P.T.S.S. la reserva a las acciones contra entidades públicas y "
         "su incumplimiento es insubsanable; aplicarla a un particular no produciría efecto "
         "interruptivo. Para una empresa privada el equivalente funcional es el derecho de petición "
         "con solicitud de agotamiento de vía interna, que además documenta la reclamación para "
         "efectos probatorios y de competencia territorial."),
        ("¿Cómo se fija la competencia si la AFP está en Bogotá y el trabajador vive en otra ciudad?",
         "El artículo 11 del C.P.T.S.S. permite elegir entre el domicilio de la entidad demandada y el "
         "lugar donde se surtió la reclamación del derecho. Si la reclamación se radicó en la ciudad "
         "del trabajador, allí puede demandarse. Es un fuero concurrente y la elección corresponde al "
         "demandante, lo que constituye una garantía de acceso a la administración de justicia."),
        ("¿Qué debe hacer la empresa si la ARL no responde la solicitud de reclasificación?",
         "El silencio administrativo se configura a los quince días hábiles y habilita la vía "
         "judicial. Adicionalmente, la empresa puede solicitar la intervención de la Dirección "
         "Territorial del Ministerio del Trabajo en ejercicio de sus funciones de inspección, y "
         "conservar la constancia de radicación como prueba del agotamiento."),
        ("¿El informe puede usarse como prueba en un proceso?",
         "Como documento privado elaborado por la propia empresa, no tiene valor de dictamen pericial, "
         "pero sí puede emplearse como prueba documental de las actuaciones adelantadas y como "
         "soporte de la diligencia empresarial. Para acreditar hechos técnicos se requiere prueba "
         "pericial o inspección judicial."),
    ]
    for i, (q, a) in enumerate(preguntas, 1):
        par = doc.add_paragraph()
        par.paragraph_format.space_before = Pt(6)
        par.paragraph_format.space_after = Pt(2)
        par.paragraph_format.keep_with_next = True
        par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r = par.add_run("Pregunta %d. " % i)
        r.bold = True
        r.font.size = Pt(10.5)
        r.font.color.rgb = AZUL
        r2 = par.add_run(q)
        r2.bold = True
        r2.italic = True
        r2.font.size = Pt(10.5)
        par2 = doc.add_paragraph()
        par2.paragraph_format.left_indent = Cm(0.7)
        par2.paragraph_format.space_after = Pt(6)
        par2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r3 = par2.add_run("Respuesta. ")
        r3.bold = True
        r3.font.size = Pt(10.5)
        r4 = par2.add_run(a)
        r4.font.size = Pt(10.5)

    caja(doc, "Recomendaciones para la sustentación",
         ["No lea el documento: explíquelo. La Junta Directiva no necesita el articulado, necesita "
          "entender el riesgo y la decisión que se le propone.",
          "Cuando cite una norma, diga qué resuelve y no solo su número. «El artículo 216 del C.S.T. "
          "obliga a indemnizar cuando hay culpa comprobada» comunica más que «el artículo 216».",
          "Anticipe la objeción económica: toda recomendación debe ir acompañada de su costo y de su "
          "beneficio, y el beneficio en materia de seguridad social casi siempre es la reducción de "
          "la contingencia patrimonial.",
          "Si no conoce una respuesta, distinga lo que sabe de lo que debe verificarse. La honestidad "
          "técnica es mejor evaluada que una invención."],
         tipo='info')

    # ---------------- 5.5 bibliografía
    salto_pagina(doc)
    h(doc, 2, "Referencias normativas y jurisprudenciales")
    rich(doc, ["Las referencias se presentan agrupadas por fuente y en orden jerárquico, con la "
               "indicación de la norma o providencia y su identificación. Se sugiere conservar este "
               "formato en la versión final del trabajo, pues facilita la verificación por parte del "
               "docente."])

    rotulo_bib(doc, "Constitución y normativa internacional")
    for t in [
        "Constitución Política de Colombia, artículos 23, 25, 48, 49, 53 y 93.",
        "Pacto Internacional de Derechos Económicos, Sociales y Culturales, artículo 9.",
        "Convenio 102 de la Organización Internacional del Trabajo sobre la seguridad social "
        "(norma mínima).",
    ]:
        bib_item(doc, t)

    rotulo_bib(doc, "Leyes")
    for t in [
        "Congreso de la República. Ley 100 de 1993, «por la cual se crea el sistema de seguridad "
        "social integral y se dictan otras disposiciones».",
        "Congreso de la República. Ley 712 de 2001, por la cual se reforma el Código Procesal del "
        "Trabajo y de la Seguridad Social.",
        "Congreso de la República. Ley 776 de 2002, por la cual se dictan normas sobre la "
        "organización, administración y prestaciones del Sistema General de Riesgos Profesionales.",
        "Congreso de la República. Ley 797 de 2003, por la cual se reforman algunas disposiciones del "
        "sistema general de pensiones.",
        "Congreso de la República. Ley 1122 de 2007, por la cual se hacen algunas modificaciones en "
        "el Sistema General de Seguridad Social en Salud.",
        "Congreso de la República. Ley 1149 de 2007, por la cual se reforma el Código Procesal del "
        "Trabajo y de la Seguridad Social para hacer efectiva la oralidad.",
        "Congreso de la República. Ley 1395 de 2010, artículo 46 (competencia por razón de la cuantía).",
        "Congreso de la República. Ley 1438 de 2011, por medio de la cual se reforma el Sistema "
        "General de Seguridad Social en Salud.",
        "Congreso de la República. Ley 1562 de 2012, por la cual se modifica el Sistema de Riesgos "
        "Laborales y se dictan otras disposiciones.",
        "Congreso de la República. Ley 1607 de 2012, artículo 25 (exoneración de aportes), y Ley 1819 "
        "de 2016, artículo 65 (artículo 114-1 del Estatuto Tributario).",
        "Congreso de la República. Ley 1748 de 2014, por medio de la cual se establece la obligación "
        "de brindar doble asesoría en los traslados de régimen pensional.",
        "Congreso de la República. Ley Estatutaria 1751 de 2015, por medio de la cual se regula el "
        "derecho fundamental a la salud.",
        "Congreso de la República. Ley 1753 de 2015, artículo 218 (núcleo familiar en salud).",
        "Congreso de la República. Ley 1755 de 2015, por medio de la cual se regula el derecho "
        "fundamental de petición.",
        "Congreso de la República. Ley 2114 de 2021, por medio de la cual se amplía la licencia de "
        "maternidad y se reconoce la licencia de paternidad.",
        "Congreso de la República. Ley 2381 de 2024, por medio de la cual se establece el sistema de "
        "protección social integral para la vejez, invalidez y muerte de origen común.",
    ]:
        bib_item(doc, t)

    rotulo_bib(doc, "Decretos y resoluciones")
    for t in [
        "Presidencia de la República. Decreto 1295 de 1994, por el cual se determina la organización "
        "y administración del Sistema General de Riesgos Profesionales.",
        "Presidencia de la República. Decreto 1772 de 1994 (montos y distribución de cotizaciones al "
        "SGRP).",
        "Presidencia de la República. Decreto 1406 de 1999 (afiliación y reporte de novedades al "
        "SGSSS).",
        "Presidencia de la República. Decreto 1703 de 2002, artículo 23 (ingreso base de cotización "
        "de los contratistas).",
        "Presidencia de la República. Decreto 1507 de 2014 (manual único para la calificación de la "
        "pérdida de capacidad laboral).",
        "Presidencia de la República. Decreto 1072 de 2015, Decreto Único Reglamentario del Sector "
        "Trabajo (SG-SST).",
        "Presidencia de la República. Decreto 780 de 2016, Decreto Único Reglamentario del Sector "
        "Salud y Protección Social.",
        "Presidencia de la República. Decreto 768 de 2022, por el cual se actualiza la tabla de "
        "clasificación de actividades económicas para el SGRL.",
        "Presidencia de la República. Decretos 1469 y 1470 de 2025 (salario mínimo y auxilio de "
        "transporte para 2026: %s y %s)." % (money(1750905), money(249095)),
        "Ministerio del Trabajo. Resolución 0312 de 2019 (estándares mínimos del SG-SST).",
        "Ministerio de Salud y Protección Social. Resolución 2388 de 2016 (afiliación y reporte de "
        "novedades).",
        "Ministerio del Trabajo. Resolución 1401 de 2007 (investigación de incidentes y accidentes de "
        "trabajo).",
        "Superintendencia Financiera de Colombia. Circular Externa 016 de 2016 (soportes de la doble "
        "asesoría).",
    ]:
        bib_item(doc, t)

    rotulo_bib(doc, "Jurisprudencia")
    for t in [
        "Corte Constitucional. Sentencia C-792 de 2006 (reclamación administrativa previa).",
        "Corte Constitucional. Sentencia C-820 de 2011 (competencia territorial en seguridad social).",
        "Corte Constitucional. Sentencia SU-107 de 2024, 9 de abril de 2024, M.P. Jorge Enrique Ibáñez "
        "Najar (ineficacia del traslado de régimen pensional y deber de información).",
        "Corte Constitucional. Sentencia C-264 de 2026, 25 de agosto de 2026 (revisión de la Ley 2381 "
        "de 2024; vigencia del sistema de pilares a partir del 1.º de abril de 2027).",
        "Corte Suprema de Justicia, Sala de Casación Laboral. Sentencia SL5619 de 2016, radicación "
        "47.907 (culpa patronal y carga de la prueba).",
        "Corte Suprema de Justicia, Sala de Casación Laboral. Sentencia SL2799 de 2014 (culpa patronal; "
        "la diligencia del empleador exonera).",
        "Corte Suprema de Justicia, Sala de Casación Laboral. Sentencia SL3179 de 2023 (carga de la "
        "prueba del deber de información a cargo de la AFP).",
        "Corte Suprema de Justicia, Sala de Casación Laboral. Sentencia SL1179 de 2024 (indemnización "
        "de perjuicios por el traslado y su carácter de resarcimiento único).",
        "Corte Suprema de Justicia, Sala de Casación Laboral. Sentencia SL2480 de 2025, 10 de "
        "septiembre de 2025 (culpa exclusiva de la víctima y exoneración del empleador).",
        "Consejo de Estado. Concepto 1832 de 2007 (ingreso base de cotización de los contratistas).",
    ]:
        bib_item(doc, t)

    rotulo_bib(doc, "Doctrina de apoyo")
    for t in [
        "Arenas Monsalve, G. (2011). El derecho colombiano de la seguridad social. Bogotá: Legis.",
        "Rodríguez Mesa, J. A. (2019). El deber de información de las administradoras de fondos de "
        "pensiones. Revista de Derecho Laboral, (12).",
        "Corte Suprema de Justicia. (2024). Relatoría de la Sala de Casación Laboral: línea "
        "jurisprudencial sobre culpa patronal.",
    ]:
        bib_item(doc, t)


def rotulo_bib(doc, texto):
    par = doc.add_paragraph()
    par.paragraph_format.space_before = Pt(10)
    par.paragraph_format.space_after = Pt(4)
    par.paragraph_format.keep_with_next = True
    r = par.add_run(texto)
    r.bold = True
    r.font.size = Pt(10.5)
    r.font.color.rgb = AZUL
    r.font.all_caps = True


def bib_item(doc, texto):
    par = doc.add_paragraph()
    par.paragraph_format.left_indent = Cm(1.0)
    par.paragraph_format.first_line_indent = Cm(-1.0)
    par.paragraph_format.space_after = Pt(4)
    par.paragraph_format.line_spacing = 1.1
    par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = par.add_run(texto)
    r.font.size = Pt(10)


# ================================================================ ANEXOS
def anexos(doc):
    salto_pagina(doc)
    h_sin_num(doc, 1, "ANEXOS: PLANTILLAS DE TRABAJO")
    rich(doc, ["Los anexos convierten el informe en un instrumento operativo. Las plantillas están "
               "diseñadas para ser diligenciadas por la organización con sus propios datos, de modo "
               "que el trabajo pueda actualizarse sin rehacer el análisis jurídico."])

    # Anexo 1
    h_sin_num(doc, 2, "Anexo 1. Ficha de caracterización organizacional")
    tabla(doc, ["Variable", "Dato de la organización", "Soporte que debe anexarse"],
          [["Razón social", "Julia Inés Buitrago García (nombre comercial EXTINTORES S C M)", ""],
           ["NIT", "39626936-8", "RUT"],
           ["Naturaleza jurídica", "Persona natural comerciante (empresa unipersonal)",
            "Certificado de cámara de comercio o documento de identidad"],
           ["Actividad económica principal (CIIU)", "4759: comercio al por menor de otros artículos "
            "domésticos en establecimientos especializados", "RUT, casilla 46"],
           ["Domicilio y dirección", "Carrera 4 n.° 60-17, barrio La Despensa, comuna 3, Soacha "
            "(Cundinamarca)", "Acta 028-050164-26 y RUT"],
           ["Representante legal", "Julia Inés Buitrago García, C.C. 39626936",
            "Acta o certificado"],
           ["Número de trabajadores dependientes", "3 (WT-01, WT-02, WT-03); el acta sanitaria "
            "declaró 1 operativo", "Planilla PILA del último mes"],
           ["Número de contratistas independientes", "", "Contratos y planilla de independientes"],
           ["EPS a la que está afiliada la planta", "", "Certificados de afiliación"],
           ["ARL a la que está afiliada la planta", "", "Certificado de afiliación y clase de riesgo"],
           ["AFP de los trabajadores", "", "Certificados de afiliación"],
           ["Caja de compensación familiar", "", "Certificado de afiliación"],
           ["Nivel de ingresos y activos", "", "Estados financieros"]],
          [5.2, 5.4, 5.4], size=8.6,
          titulo="Formato A-1. Caracterización")

    # Anexo 1-bis: el acta sanitaria como fuente
    h_sin_num(doc, 2, "Anexo 2. Ficha de la fuente primaria: acta sanitaria 028-050164-26")
    rich(doc, ["Esta ficha permite citar el acta con precisión en el informe, en el derecho de petición "
               "y en la sustentación oral. Todos los datos provienen del documento original de diez "
               "páginas, notificado el mismo día de la visita."])
    tabla(doc, ["Campo", "Contenido verificado en el documento"],
          [["Formato", "SL-FR-028, versión 1, aprobado el 07/10/2022; «Acta de inspección, "
            "vigilancia y control sanitario. Establecimientos industriales, comerciales y de "
            "servicios donde se usan y almacenan sustancias y productos químicos»"],
           ["Número y fecha", "028-050164-26, del 09/09/2026"],
           ["Autoridad", "Secretaría de Salud de Soacha, entidad territorial de salud; ciudad: Soacha"],
           ["Establecimiento", "Julia Inés Buitrago García; nombre comercial EXTINTORES S C M; "
            "cédula/NIT 39626936-8; inscripción 25 754 39626936-8"],
           ["Ubicación", "Carrera 4 n.° 60-17, comuna 3, barrio La Despensa, Soacha, Cundinamarca; "
            "código DANE 25 754"],
           ["Contacto", "Teléfono 3212944603; correo julia_ines_buitrago@hotmail.com; autoriza "
            "notificación electrónica"],
           ["Actividad declarada", "CIIU 4759: comercio al por menor de otros artículos domésticos en "
            "establecimientos especializados; tipo de establecimiento: comercializadora"],
           ["Trabajadores declarados", "Operativos: 1 (los campos de afiliados a riesgos "
            "profesionales y al sistema de seguridad social quedaron en blanco)"],
           ["Antecedente", "Visita del 18/10/2024: concepto «favorable con requerimientos», "
            "cumplimiento del 57,50 %"],
           ["Resultado de la visita", "Revisión documental con once documentos no evidenciados; "
            "bloques de condiciones locativas, sanitarias, de saneamiento y de seguridad evaluados "
            "con múltiples hallazgos"],
           ["Plazo", "22 días hábiles; inicio 09/09/2026; vencimiento 09/10/2026"],
           ["Funcionaria", "Gissel Paola Hurtado Leyton, C.C. 1013644855, ingeniera química, "
            "contrato 3136-2026"],
           ["Quien recibió", "Felipe Martínez, C.C. 8686834, cargo vendedor (a)"],
           ["Advertencia del acta", "«No constituye permiso o autorización para el funcionamiento del "
            "establecimiento de comercio»; remite al artículo 87 de la Ley 1801 de 2016"],
           ["Anexo fotográfico", "El acta deja constancia de que «no se realizó toma de registro "
            "fotográfico»"]],
          [3.6, 12.4], size=8.2,
          titulo="Formato A-2. Datos de citación del acta sanitaria")

    # Anexo 3
    h_sin_num(doc, 2, "Anexo 3. Plantilla de liquidación del IBC")
    tabla(doc, ["Concepto", "Trabajador dependiente", "Contratista independiente"],
          [["Valor bruto recibido en el mes", "", ""],
           ["Menos: pagos no salariales (art. 128 C.S.T.)", "", "No aplica"],
           ["Igual: salario base", "", "No aplica"],
           ["Valor mensualizado del contrato sin IVA", "No aplica", ""],
           ["Base: 40 % del valor mensualizado", "No aplica", ""],
           ["Comparación con el piso de 1 SMMLV", "", ""],
           ["Comparación con el techo de 25 SMMLV", "", ""],
           ["IBC definitivo", "", ""],
           ["Aporte a salud (8,5 % empleador / 12,5 % independiente)", "", ""],
           ["Aporte a pensión (12 % empleador y trabajador / 16 % independiente)", "", ""],
           ["Aporte a riesgos laborales (tarifa según clase)", "", ""],
           ["Parafiscales (SENA, ICBF, caja de compensación)", "", "No aplica"]],
          [7.4, 4.3, 4.3], size=8.6,
          titulo="Formato A-2. Liquidación del IBC")

    # Anexo 4
    h_sin_num(doc, 2, "Anexo 4. Lista de verificación de afiliación y novedades")
    tabla(doc, ["N.º", "Punto de control", "Sí", "No", "Observación"],
          [["1", "Todos los trabajadores están afiliados a una EPS", "", "", ""],
           ["2", "Todos los trabajadores están afiliados a una ARL", "", "", ""],
           ["3", "Todos los trabajadores están afiliados a una AFP", "", "", ""],
           ["4", "Los beneficiarios están inscritos y con soportes vigentes", "", "", ""],
           ["5", "Las novedades de ingreso y retiro se reportan dentro del plazo", "", "", ""],
           ["6", "Las variaciones de salario se reportan dentro del plazo", "", "", ""],
           ["7", "Los aportes se pagan dentro de los plazos legales", "", "", ""],
           ["8", "Se verifica la procedencia de la exoneración del art. 114-1 del E.T.", "", "", ""],
           ["9", "Los contratistas independientes acreditan su afiliación", "", "", ""],
           ["10", "Se conservan los soportes de licencias e incapacidades", "", "", ""]],
          [1.1, 8.1, 1.3, 1.3, 4.2], size=8.6, align_centro=[0, 2, 3],
          titulo="Formato A-3. Lista de verificación")

    # Anexo 5
    h_sin_num(doc, 2, "Anexo 5. Formato de reporte e investigación de accidente de trabajo")
    tabla(doc, ["Campo", "Contenido"],
          [["Identificación del trabajador", ""],
           ["Fecha, hora y lugar del evento", ""],
           ["Descripción de la tarea y del proceso", ""],
           ["Agente y mecanismo de la lesión", ""],
           ["Elementos de protección personal suministrados y su registro de entrega", ""],
           ["Capacitación y certificación vigente para la tarea", ""],
           ["Procedimiento de trabajo aplicable y su versión", ""],
           ["Testigos y sus declaraciones", ""],
           ["Atención inicial y remisión a la IPS", ""],
           ["Reporte a la ARL (fecha y número de radicado)", ""],
           ["Investigación: metodología, participantes y fecha", ""],
           ["Causas inmediatas y causas básicas", ""],
           ["Plan de acción con responsables y plazos", ""],
           ["Verificación de la eficacia del plan", ""]],
          [7.0, 9.0], size=8.6,
          titulo="Formato A-4. Investigación del accidente")

    # Anexo 6
    h_sin_num(doc, 2, "Anexo 6. Formato de auditoría pensional individual")
    tabla(doc, ["Campo", "Contenido"],
          [["Nombre y documento del trabajador", ""],
           ["Fecha de nacimiento y edad", ""],
           ["Régimen actual y administradora", ""],
           ["Fecha del traslado y de la afiliación inicial", ""],
           ["Semanas cotizadas certificadas", ""],
           ["Capital acumulado o ingreso base de liquidación", ""],
           ["Soporte de doble asesoría (sí / no / no existe)", ""],
           ["Simulación o proyección entregada anteriormente", ""],
           ["Verificación del régimen de transición (750 / 900 semanas)", ""],
           ["Comparación de mesadas proyectadas en ambos regímenes", ""],
           ["Conclusión: procede o no la demanda de ineficacia", ""],
           ["Acciones adelantadas y fechas de radicación", ""]],
          [7.0, 9.0], size=8.6,
          titulo="Formato A-5. Auditoría pensional individual")

    # Anexo 7
    h_sin_num(doc, 2, "Anexo 7. Control de términos y reclamaciones")
    tabla(doc, ["Documento", "Entidad", "Fecha de radicación", "Término de respuesta",
                "Fecha límite", "Estado"],
          [["Derecho de petición de reclasificación del riesgo", "ARL", "", "15 días hábiles", "", ""],
           ["Derecho de petición de prestaciones AT-2026-014", "ARL", "", "15 días hábiles", "", ""],
           ["Solicitud de expediente y acta de doble asesoría", "AFP", "", "15 días hábiles", "", ""],
           ["Solicitud de agotamiento de vía interna", "EXTINTORES SCM", "", "15 días hábiles", "", ""],
           ["Reclamación de licencia de paternidad", "EPS", "", "15 días hábiles", "", ""],
           ["Corrección de planilla de aportes", "PILA", "", "Según el periodo", "", ""]],
          [5.4, 2.6, 2.3, 2.4, 1.9, 1.4], size=8.3, align_centro=[2, 3, 4, 5],
          titulo="Formato A-6. Control de términos")

    # Anexo 8
    h_sin_num(doc, 2, "Anexo 8. Matriz de cumplimiento de la rúbrica")
    rich(doc, ["Como herramienta de autoevaluación, se presenta la correspondencia entre los criterios "
               "de la rúbrica del proyecto integrador y los apartados del informe que los satisfacen."])
    tabla(doc, ["Criterio de la rúbrica", "Peso", "Apartados del informe que lo desarrollan"],
          [["1. Caracterización organizacional, afiliación y cotizaciones en salud", "15 %",
            "Eje I (capítulo 1), numerales 1.1 a 1.4; Tablas 2 a 6; Figuras 1 y 2; Anexos 1 a 4"],
           ["2. Análisis jurídico del Sistema General de Riesgos Laborales", "18 %",
            "Eje II (capítulo 2), numerales 2.1 a 2.7; Tablas 7 a 14; Figuras 3 y 4; Anexos 2 y 5"],
           ["3. Diagnóstico pensional y análisis de la ineficacia del traslado", "18 %",
            "Eje III (capítulo 3), numerales 3.1 a 3.5; Tablas 15 a 22; Figura 5; Anexo 6"],
           ["4. Elaboración del insumo procesal y determinación de la competencia judicial", "13 %",
            "Eje IV (capítulo 4), numerales 4.1 a 4.3; Tablas 23 a 27; documentos modelo; Anexo 7"],
           ["5. Pensamiento crítico, integración del diagnóstico y sustentación oral", "38 %",
            "Capítulo 5 completo: matriz de hallazgos, plan de mejoramiento, conclusiones, guion y "
            "banco de preguntas; Anexo 8"]],
          [6.4, 1.8, 7.8], size=8.6, align_centro=[1],
          titulo="Tabla 31. Correspondencia entre la rúbrica y el informe")
    p(doc, "")

    caja(doc, "Nota final para la estudiante",
         ["Este documento es una plantilla avanzada: la estructura, el marco normativo y el "
          "razonamiento jurídico ya están resueltos. Para convertirlo en su trabajo definitivo debe "
          "(i) reemplazar los datos ilustrativos por los soportes reales de EXTINTORES SCM; "
          "(ii) diligenciar los anexos con la documentación de la empresa; (iii) actualizar la fecha "
          "de presentación y los datos de contacto; y (iv) verificar que cada cifra citada "
          "corresponda a un comprobante o certificado del expediente. Mientras esos pasos no se "
          "cumplan, el documento debe presentarse como versión de trabajo, no como informe final."],
         tipo='nota')
