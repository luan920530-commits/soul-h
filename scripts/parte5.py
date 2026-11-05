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
          [["1", "I", "Novedades de PILA reportadas de manera extemporánea en dos casos", "Medio",
            "Talento humano", "30 días"],
           ["2", "I", "Beneficiario mayor de 25 años sin soporte de dependencia económica", "Medio",
            "Talento humano", "15 días"],
           ["3", "I", "Diferencia de %s en la licencia de maternidad de EXT-004" % money(412000),
            "Medio", "Gerencia", "30 días"],
           ["4", "I", "Aplicación de la exoneración del art. 114-1 del E.T. sin verificación formal",
            "Medio", "Contador", "Inmediato"],
           ["5", "II", "Clasificación del riesgo inconsistente con la actividad económica", "Medio",
            "Gerencia", "60 días"],
           ["6", "II", "SG-SST sin evidencia documental (6 de 7 deberes incumplidos)", "Alto",
            "Responsable de SST", "45 días"],
           ["7", "II", "Certificaciones vencidas para tareas críticas", "Alto",
            "Responsable de SST", "30 días"],
           ["8", "II", "Reporte extemporáneo del accidente (5 días)", "Medio", "Talento humano",
            "Inmediato"],
           ["9", "III", "Dos traslados pensionales sin soporte de asesoría", "Alto",
            "Gerencia", "90 días"],
           ["10", "III", "Semanas no reportadas de EXT-010", "Alto", "Contador", "30 días"],
           ["11", "IV", "Ausencia de archivo de reclamaciones y de respuestas de las entidades",
            "Bajo", "Gerencia", "45 días"]],
          [1.1, 1.1, 7.4, 1.6, 3.0, 1.8], size=8.2, align_centro=[0, 1, 3, 5],
          titulo="Tabla 25. Matriz integrada de hallazgos",
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
           ["Riesgo pensional", "Revisar la historia laboral de los diez trabajadores",
            "Porcentaje de historias verificadas contra la planilla", "100 % en 90 días"],
           ["Riesgo pensional", "Radicar solicitudes de expediente ante las AFP",
            "Solicitudes radicadas y respuestas archivadas", "3 de 3"],
           ["Gestión procesal", "Crear el archivo de reclamaciones y respuestas",
            "Expedientes con radicado, respuesta y control de términos", "100 %"],
           ["Gestión procesal", "Diseñar el protocolo de conservación documental",
            "Política aprobada con plazos de retención", "1 política"]],
          [3.3, 5.6, 4.7, 2.4], size=8.3,
          titulo="Tabla 26. Plan de mejoramiento")

    rich(doc, ["Se sugiere una secuencia de implementación en tres momentos. En los primeros treinta "
               "días deben atenderse las acciones de riesgo alto que dependen solo de decisiones "
               "internas: recertificación, protocolo de reporte y corrección de semanas. Entre los "
               "días treinta y sesenta corresponde la actualización documental del SG-SST y la "
               "radicación de las solicitudes ante la ARL y las AFP. Y entre los días sesenta y "
               "noventa, la consolidación del archivo, la verificación de la historia laboral y la "
               "evaluación de los resultados con indicadores."])

    # ---------------- 5.3 conclusiones
    h(doc, 2, "Conclusiones")
    conclusiones = [
        ("Sobre la caracterización y la afiliación. ",
         "EXTINTORES SCM es una empresa privada de persona natural cuya actividad se clasifica en el "
         "código CIIU 4752 y que cumple, en lo esencial, con la afiliación al Régimen Contributivo. "
         "Sin embargo, la gestión de novedades y la verificación de la base de cotización presentan "
         "falencias que, aunque no generan una crisis inmediata, sí producen un riesgo acumulado de "
         "reliquidación y de afectación de los derechos de los trabajadores."),
        ("Sobre la liquidación del IBC. ",
         "La correcta determinación del IBC exige distinguir entre pagos salariales y no salariales "
         "(arts. 127 y 128 del C.S.T.), aplicar el piso de un salario mínimo y el techo de veinticinco "
         "salarios mínimos, y tratar de manera diferenciada al trabajador dependiente (que cotiza sobre el salario depurado) y al contratista independiente, que cotiza sobre el 40 % del "
         "valor mensualizado del contrato. La confusión entre ambos regímenes es la causa más "
         "frecuente de errores de liquidación en la microempresa."),
        ("Sobre los riesgos laborales. ",
         "La clasificación del riesgo debe corresponder a la actividad económica real y no a la "
         "práctica administrativa heredada. En el caso examinado, la clase I es la que corresponde a "
         "la actividad principal registrada, pero la existencia de procesos conexos de recarga y "
         "prueba hidrostática obliga a una consulta formal a la ARL antes de reliquidar. El caso "
         "AT-2026-014 demuestra que el incumplimiento documental del SG-SST convierte un accidente "
         "en una condena: seis de siete deberes de seguridad estaban incumplidos."),
        ("Sobre la culpa patronal. ",
         "La responsabilidad del artículo 216 del C.S.T. se funda en la culpa leve y no en la culpa "
         "grave: basta la omisión de la diligencia ordinaria. Aunque la carga de la prueba recae "
         "inicialmente en el trabajador, una vez acreditada la omisión el empleador debe demostrar "
         "diligencia, y esa demostración es documental. El SG-SST es, por tanto, un mecanismo de "
         "defensa probatoria antes que un requisito formal."),
        ("Sobre el diagnóstico pensional. ",
         "La coexistencia de dos regímenes y la entrada en vigencia de la Ley 2381 de 2024 el 1.º de "
         "abril de 2027 configuran un escenario de alta complejidad para los trabajadores próximos a "
         "la edad de pensión. La verificación temprana de la historia laboral es determinante, porque "
         "el umbral de 750 y 900 semanas define quién conserva el régimen anterior."),
        ("Sobre la ineficacia del traslado. ",
         "El estándar de la Sentencia SU-107 de 2024 hace viable la demanda cuando la administradora "
         "no acredita información suficiente y asesoría documentada. El formulario con la cláusula de "
         "afiliación «libre y voluntaria» es insuficiente. Los efectos restitutorios son completos: "
         "se restablece la afiliación sin solución de continuidad y se traslada todo el ahorro, con "
         "rendimientos y gastos de administración."),
        ("Sobre el insumo procesal y la competencia. ",
         "La calificación de la naturaleza del sujeto pasivo determina el mecanismo: reclamación "
         "administrativa previa para entidades públicas y derecho de petición con solicitud de "
         "agotamiento de vía interna para empresas privadas. La competencia se fija por la materia "
         "(art. 2), el territorio (art. 11, con fuero concurrente a elección del demandante) y la "
         "cuantía (art. 12, con umbral de veinte salarios mínimos)."),
        ("Sobre el valor del diagnóstico. ",
         "La auditoría no se agota en el señalamiento de incumplimientos: su valor está en traducir "
         "los hallazgos en decisiones jurídicamente sustentadas, con responsables, plazos e "
         "indicadores. Esa es la diferencia entre un informe descriptivo y una consultoría jurídica."),
    ]
    for titulo, texto in conclusiones:
        par = doc.add_paragraph()
        par.paragraph_format.space_after = Pt(7)
        par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r1 = par.add_run(titulo)
        r1.bold = True
        r1.font.size = Pt(11)
        r2 = par.add_run(texto)
        r2.font.size = Pt(11)
        par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

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
          titulo="Tabla 27. Estructura de la sustentación oral")

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
          [["Razón social", "", ""],
           ["NIT", "", "RUT"],
           ["Naturaleza jurídica", "", "Certificado de cámara de comercio o documento de identidad"],
           ["Actividad económica principal (CIIU)", "", "RUT, casilla 46"],
           ["Domicilio y dirección", "", ""],
           ["Representante legal", "", "Acta o certificado"],
           ["Número de trabajadores dependientes", "", "Planilla PILA del último mes"],
           ["Número de contratistas independientes", "", "Contratos y planilla de independientes"],
           ["EPS a la que está afiliada la planta", "", "Certificados de afiliación"],
           ["ARL a la que está afiliada la planta", "", "Certificado de afiliación y clase de riesgo"],
           ["AFP de los trabajadores", "", "Certificados de afiliación"],
           ["Caja de compensación familiar", "", "Certificado de afiliación"],
           ["Nivel de ingresos y activos", "", "Estados financieros"]],
          [5.2, 5.4, 5.4], size=8.6,
          titulo="Formato A-1. Caracterización")

    # Anexo 2
    h_sin_num(doc, 2, "Anexo 2. Plantilla de liquidación del IBC")
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

    # Anexo 3
    h_sin_num(doc, 2, "Anexo 3. Lista de verificación de afiliación y novedades")
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

    # Anexo 4
    h_sin_num(doc, 2, "Anexo 4. Formato de reporte e investigación de accidente de trabajo")
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

    # Anexo 5
    h_sin_num(doc, 2, "Anexo 5. Formato de auditoría pensional individual")
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

    # Anexo 6
    h_sin_num(doc, 2, "Anexo 6. Control de términos y reclamaciones")
    tabla(doc, ["Documento", "Entidad", "Fecha de radicación", "Término de respuesta",
                "Fecha límite", "Estado"],
          [["Derecho de petición de reclasificación del riesgo", "ARL", "", "15 días hábiles", "", ""],
           ["Derecho de petición de prestaciones AT-2026-014", "ARL", "", "15 días hábiles", "", ""],
           ["Solicitud de expediente y acta de doble asesoría", "AFP", "", "15 días hábiles", "", ""],
           ["Solicitud de agotamiento de vía interna", "EXTINTORES SCM", "", "15 días hábiles", "", ""],
           ["Reclamación de licencia de maternidad", "EPS", "", "15 días hábiles", "", ""],
           ["Corrección de planilla de aportes", "PILA", "", "Según el periodo", "", ""]],
          [5.4, 2.6, 2.3, 2.4, 1.9, 1.4], size=8.3, align_centro=[2, 3, 4, 5],
          titulo="Formato A-6. Control de términos")

    # Anexo 7
    h_sin_num(doc, 2, "Anexo 7. Matriz de cumplimiento de la rúbrica")
    rich(doc, ["Como herramienta de autoevaluación, se presenta la correspondencia entre los criterios "
               "de la rúbrica del proyecto integrador y los apartados del informe que los satisfacen."])
    tabla(doc, ["Criterio de la rúbrica", "Peso", "Apartados del informe que lo desarrollan"],
          [["1. Caracterización organizacional, afiliación y cotizaciones en salud", "15 %",
            "Eje I (capítulo 1), numerales 1.1 a 1.4; Tablas 2 a 5; Figuras 1 y 2; Anexos 1 a 3"],
           ["2. Análisis jurídico del Sistema General de Riesgos Laborales", "18 %",
            "Eje II (capítulo 2), numerales 2.1 a 2.6; Tablas 6 a 11; Figuras 3 y 4; Anexo 4"],
           ["3. Diagnóstico pensional y análisis de la ineficacia del traslado", "18 %",
            "Eje III (capítulo 3), numerales 3.1 a 3.5; Tablas 12 a 19; Figura 5; Anexo 5"],
           ["4. Elaboración del insumo procesal y determinación de la competencia judicial", "13 %",
            "Eje IV (capítulo 4), numerales 4.1 a 4.3; Tablas 20 a 24; documentos modelo; Anexo 6"],
           ["5. Pensamiento crítico, integración del diagnóstico y sustentación oral", "38 %",
            "Capítulo 5 completo: matriz de hallazgos, plan de mejoramiento, conclusiones, guion y "
            "banco de preguntas; Anexo 7"]],
          [6.4, 1.8, 7.8], size=8.6, align_centro=[1],
          titulo="Tabla 28. Correspondencia entre la rúbrica y el informe")
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
