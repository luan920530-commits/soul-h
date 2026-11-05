# -*- coding: utf-8 -*-
"""EJE II: Sistema General de Riesgos Laborales."""
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx_engine import (p, rich, h, bullet, cita, tabla, caja, salto_pagina,
                         imagen, AZUL, VINO, GRIS, GRIS_CLARO, HEX_VINO)
from datos import SMLMV, money, num, TARIFAS, IBC_TOTAL, NOMINA

IBC_CASO = 1900000


def eje_ii(doc):
    salto_pagina(doc)
    h(doc, 1, "EJE II · SISTEMA GENERAL DE RIESGOS LABORALES (SGRL)")

    # ---------- 2.1 marco
    h(doc, 2, "Marco normativo del Sistema General de Riesgos Laborales")
    rich(doc, ["El Sistema General de Riesgos Laborales (SGRL) fue creado por la Ley 100 de 1993 "
               "(Libro Tercero), organizado por el Decreto 1295 de 1994 y reformado de manera "
               "sustancial por la Ley 1562 de 2012, que cambió la denominación de «riesgos "
               "profesionales» por la de «riesgos laborales», amplió la definición de accidente de "
               "trabajo, introdujo el concepto de enfermedad laboral y extendió la afiliación "
               "obligatoria a los trabajadores independientes. Su desarrollo reglamentario se "
               "encuentra compilado en el Decreto 1072 de 2015, que contiene el Decreto Único "
               "Reglamentario del Sector Trabajo, y en la Resolución 0312 de 2019, que fija los "
               "estándares mínimos del Sistema de Gestión de Seguridad y Salud en el Trabajo."])
    rich(doc, ["El sistema se financia con una cotización a cargo exclusivo del empleador, "
               "equivalente a un porcentaje del IBC que varía según la clase de riesgo (arts. 18 y 19 "
               "del Decreto 1295 de 1994), y cubre dos tipos de prestaciones: asistenciales, que "
               "comprenden la atención médica integral, y económicas, que compensan la pérdida de "
               "capacidad laboral. La afiliación es obligatoria y su omisión genera una "
               "responsabilidad directa del empleador por las prestaciones que la ARL habría "
               "otorgado, además de las sanciones administrativas (art. 13 de la Ley 1562 de 2012)."])

    h(doc, 3, "Definiciones legales de accidente de trabajo y enfermedad laboral")
    rich(doc, ["El artículo 3 de la Ley 1562 de 2012 define el accidente de trabajo como ",
               ("«todo suceso repentino que sobrevenga por causa o con ocasión del trabajo, y que "
                "produzca en el trabajador una lesión orgánica, una perturbación funcional o "
                "psiquiátrica, una invalidez o la muerte»", {"i": True}),
               ". La misma norma precisa que también lo es el ocurrido durante la ejecución de "
               "órdenes del empleador, aun fuera del lugar y horas de trabajo, el sufrido durante el "
               "traslado en los vehículos de la empresa y el que se produce en el tiempo destinado a "
               "la actividad sindical. Por su parte, el artículo 4 define la enfermedad laboral como ",
               ("«la contraída como resultado de la exposición a factores de riesgo inherentes a la "
                "actividad del trabajador o del medio ambiente de trabajo»", {"i": True}),
               ", y la distingue de la enfermedad común, que se rige por el sistema de salud."])
    rich(doc, ["Esta distinción es determinante en la práctica: de ella depende quién asume la "
               "prestación (la ARL o la EPS), cuál es el porcentaje del subsidio (100 % del salario "
               "base de cotización en el origen laboral, frente al 66,67 % en la enfermedad general) "
               "y qué régimen de responsabilidad se activa para el empleador."])

    # ---------- 2.2 clasificación del riesgo
    h(doc, 2, "Clasificación del riesgo y tarifa de cotización")
    rich(doc, ["La clasificación del riesgo se determina con la tabla de actividades económicas "
               "adoptada por el Decreto 768 de 2022, que sustituyó el Decreto 1607 de 2002 y adoptó "
               "la CIIU revisión 4 adaptada para Colombia. El código consta de siete dígitos: el "
               "primero indica la clase de riesgo (I a V) y los restantes identifican la actividad y "
               "la subactividad. El procedimiento es reglado: (i) se identifica el código CIIU de la "
               "actividad principal registrada en el Registro Único Tributario; (ii) se ubica ese "
               "código en la tabla; y (iii) cuando la actividad no figura, la empresa y la ARL la "
               "clasifican por la actividad más afín."])
    tabla(doc, ["Clase", "Nivel de riesgo", "Tarifa inicial (% IBC)", "Aporte mensual por trabajador con 1 SMMLV"],
          [["I", "Mínimo", "0,522 %", money(SMLMV * 0.00522)],
           ["II", "Bajo", "1,044 %", money(SMLMV * 0.01044)],
           ["III", "Medio", "2,436 %", money(SMLMV * 0.02436)],
           ["IV", "Alto", "4,350 %", money(SMLMV * 0.04350)],
           ["V", "Máximo", "6,960 %", money(SMLMV * 0.06960)]],
          [1.8, 4.0, 4.4, 5.8], size=9,
          titulo="Tabla 6. Clases de riesgo y tarifas de cotización al SGRL",
          align_centro=[0, 2, 3],
          fuente="Fuente: artículo 13 del Decreto 1772 de 1994 y artículo 18 del Decreto 1295 de 1994. "
                 "Dentro de cada clase la ARL puede fijar la tarifa entre un mínimo y un máximo, según "
                 "el índice de lesiones incapacitantes y el cumplimiento del SG-SST.")

    h(doc, 3, "Clasificación de EXTINTORES SCM")
    rich(doc, ["La actividad principal registrada corresponde al código CIIU 4752 (comercio al por menor de artículos de ferretería, pinturas y productos de vidrio en establecimientos especializados), que la tabla del Decreto 768 de 2022 ubica en la ",
               ("clase de riesgo I, subactividades 4752-01 y 4752-02", {"b": True}),
               ", con una tarifa inicial del 0,522 %% del IBC. La verificación arroja una "
               "discrepancia relevante: la empresa venía cotizando sobre una base asimilada a la "
               "clase II, lo que generó un sobrecosto mensual de aproximadamente %s." % money(
                   IBC_TOTAL * (TARIFAS['II'] - TARIFAS['I']))])
    rich(doc, ["Sin embargo, la auditoría no puede cerrarse con la sola lectura del código. La "
               "empresa desarrolla actividades conexas de recarga, mantenimiento y prueba "
               "hidrostática de extintores que introducen riesgos (presión, proyección de partes, manejo de agentes químicos y esfuerzo físico) que podrían justificar una "
               "reclasificación si la ARL considera que la actividad real difiere de la declarada. "
               "El artículo 3 del Decreto 768 de 2022 permite clasificar por la actividad más afín "
               "cuando la principal no refleja el riesgo, y el artículo 4 obliga a cotizar por el "
               "valor inicial de la nueva clase cuando la clasificación se modifica. Por ello se "
               "recomienda una consulta formal a la ARL, con la descripción del proceso productivo, "
               "antes de solicitar la reliquidación."])
    caja(doc, "Escenarios de clasificación y efecto económico",
         [[("Escenario A (actividad principal 4752): ", {"b": True}),
           "clase I, tarifa 0,522 %% (aporte mensual sobre la nómina de %s: %s." % ( money(IBC_TOTAL), money(IBC_TOTAL * TARIFAS['I']))], [("Escenario B (reclasificación por actividad conexa a clase II): ", {"b": True}), "tarifa 1,044 %%) aporte mensual: %s. La diferencia anual asciende a %s." % (
               money(IBC_TOTAL * TARIFAS['II']), money(IBC_TOTAL * TARIFAS['II'] * 12))],
          "La decisión debe sustentarse en el proceso real y no en la conveniencia económica: una "
          "clasificación subdeclarada expone a la empresa al cobro retroactivo y, en caso de "
          "siniestro, a que la ARL discuta la cobertura."],
         tipo='nota')

    # ---------- 2.3 SG-SST
    h(doc, 2, "Obligaciones del empleador en materia de seguridad y salud en el trabajo")
    rich(doc, ["El artículo 13 de la Ley 1562 de 2012, que modificó el artículo 13 del Decreto 1295 "
               "de 1994, enumera las obligaciones del empleador: afiliar a todos sus trabajadores al "
               "SGRL, implementar el Sistema de Gestión de Seguridad y Salud en el Trabajo, reportar "
               "los accidentes y enfermedades laborales a la ARL, permitir la participación de los "
               "trabajadores y sus representantes, y adoptar medidas de prevención. El SG-SST debe "
               "ser implementado por todos los empleadores, con independencia de su tamaño, y su "
               "incumplimiento es sancionable por la Dirección Territorial del Ministerio del "
               "Trabajo."])
    rich(doc, ["La Resolución 0312 de 2019 graduó la exigencia según el número de trabajadores. Para "
               "empresas de uno a diez trabajadores se exigen ",
               ("siete (7) estándares mínimos", {"b": True}),
               ", lo que en principio facilita el cumplimiento de EXTINTORES SCM; sin embargo, la "
               "sencillez del estándar no reduce el deber de diligencia, porque la responsabilidad "
               "del artículo 216 del C.S.T. se mide por el resultado y no por el tamaño de la "
               "empresa."])
    tabla(doc, ["Estándar mínimo (Res. 0312 de 2019, Anexo III)", "Evidencia exigible en la empresa", "Estado"],
          [["Responsable de diseñar e implementar el SG-SST", "Designación escrita y acta de "
            "asignación de funciones con recursos", "Cumple parcialmente: designación verbal"],
           ["Afiliación al SGSSS y al SGRL", "Certificados de afiliación y planilla PILA",
            "Cumple"],
           ["Capacitación en SST e inducción", "Registros de asistencia y contenido temático",
            "No cumple: no hay evidencia de capacitación en manejo de cilindros presurizados"],
           ["Plan anual de trabajo en SST", "Documento con actividades, responsables, plazos y "
            "presupuesto", "No cumple: el plan no se actualizó en 2026"],
           ["Identificación de peligros y evaluación de riesgos", "Matriz de peligros actualizada "
            "y valorada", "Cumple parcialmente: matriz desactualizada frente a la recarga"],
           ["Plan de prevención, preparación y respuesta ante emergencias", "Plan de emergencias, "
            "brigada y simulacros", "Cumple parcialmente: sin simulacro documentado"],
           ["Reporte e investigación de incidentes y accidentes", "FURAT y planilla de reporte",
            "No cumple: el caso AT-2026-014 se reportó con 5 días de retardo"]],
          [5.4, 6.3, 4.3], size=8.4,
          titulo="Tabla 7. Verificación de los estándares mínimos del SG-SST")

    # ---------- 2.4 caso práctico
    h(doc, 2, "Caso práctico: accidente de trabajo AT-2026-014")
    h(doc, 3, "Hechos relevantes")
    rich(doc, ["El trabajador EXT-005, técnico de recarga con contrato a término indefinido, "
               "devengando un salario base de cotización de %s, se encontraba el 12 de agosto de 2026 "
               "en la zona de recarga de cilindros efectuando la prueba hidrostática de un extintor "
               "de agente limpio con veinte años de servicio. Al accionar la válvula del cilindro, "
               "esta se desprendió por corrosión y la proyectó contra su mano derecha, causándole "
               "una fractura conminuta de los dedos índice y medio y lesión del tendón flexor. El "
               "trabajador no portaba guantes antifallo, el cilindro no estaba anclado al banco de "
               "prueba, el procedimiento de trabajo no había sido actualizado y la certificación del "
               "trabajador para manipulación de cilindros presurizados estaba vencida desde noviembre "
               "de 2025." % money(IBC_CASO)])
    rich(doc, ["El accidente se reportó a la ARL y a la EPS cinco días después de ocurrido. La "
               "investigación del evento se elaboró cuatro semanas más tarde, sin la participación "
               "del trabajador ni de los representantes del SG-SST, y con la conclusión (insuficiente) de que «el trabajador actuó con exceso de confianza»."])

    h(doc, 3, "Determinación del nexo causal")
    rich(doc, ["La calificación del origen laboral exige acreditar tres elementos: (i) la existencia "
               "de una lesión orgánica o perturbación funcional; (ii) su ocurrencia en relación con "
               "la actividad laboral; y (iii) la relación de causalidad entre el trabajo y el daño. "
               "En el caso examinado los tres están presentes de manera inequívoca y sin necesidad "
               "de acudir a presunciones:"])
    for t in [
        [("Elemento fáctico: ", {"b": True}), "el accidente se produjo en las instalaciones de la "
         "empresa, en jornada laboral, durante la ejecución de una tarea contratada (la prueba "
         "hidrostática). El lugar, el tiempo y la actividad están documentados en el reporte y en "
         "los registros de trabajo del taller."],
        [("Elemento objetivo: ", {"b": True}), "existe un daño demostrado (fractura conminuta y lesión tendinosa) con incapacidad certificada por la ARL y posterior calificación de "
         "pérdida de capacidad laboral."],
        [("Elemento causal: ", {"b": True}), "el daño no habría ocurrido fuera de la ejecución de "
         "la tarea: el desprendimiento de la válvula es un riesgo inherente a la actividad de "
         "recarga y la omisión de anclaje, de guantes y de capacitación ejecutada incrementa la "
         "probabilidad de que el riesgo se materialice."],
    ]:
        bullet(doc, t)
    rich(doc, ["Conviene precisar que, incluso si el origen fuera discutible, el artículo 3 de la "
               "Ley 1562 de 2012 y el principio de favorabilidad que se desprende del artículo 53 "
               "constitucional orientan la resolución de la duda a favor del trabajador cuando el "
               "hecho se produce en el lugar y en el tiempo de la ejecución del trabajo. En el caso, "
               "no hay margen razonable de discusión."])

    h(doc, 3, "Prestaciones asistenciales a cargo de la ARL")
    rich(doc, ["Las prestaciones asistenciales están definidas en el artículo 5 del Decreto 1295 de "
               "1994 y comprenden, entre otras, la asistencia médica, quirúrgica, terapéutica y "
               "farmacéutica; los servicios de hospitalización; el servicio odontológico; el "
               "suministro de medicamentos; los servicios auxiliares de diagnóstico y tratamiento; "
               "las prótesis y aparatos ortopédicos y su reparación; la rehabilitación física y "
               "profesional; el suministro de calzado y overoles; y la orientación ocupacional. La "
               "«reparación tarifada» propia del sistema significa que la ARL asume el costo íntegro "
               "de la atención, sin que la víctima deba asumirlo y con independencia de la culpa del "
               "empleador. El trabajador conserva además el derecho a la reubicación laboral cuando "
               "existe disminución de capacidad (art. 16 de la Ley 776 de 2002)."])

    h(doc, 3, "Prestaciones económicas y liquidación")
    rich(doc, ["Las prestaciones económicas del sistema son el subsidio por incapacidad temporal, la "
               "indemnización por incapacidad permanente parcial, la pensión de invalidez, la pensión "
               "de sobrevivientes y el auxilio funerario (art. 7 del Decreto 1295 de 1994). Sus "
               "porcentajes y montos fueron redefinidos por la Ley 776 de 2002. El subsidio por "
               "incapacidad temporal equivale al ",
               ("100 % del salario base de cotización", {"b": True}),
               " y se reconoce desde el día siguiente al del accidente (art. 1 de la Ley 776 de "
               "2002), lo que lo diferencia del subsidio por enfermedad general, que es del 66,67 % "
               "a cargo de la EPS a partir del tercer día."])
    tabla(doc, ["Concepto", "Fórmula aplicada", "Valor"],
          [["Salario base de cotización (SBC)", "Base de cotización de los seis meses anteriores",
            "$ 1.900.000"],
           ["Incapacidad temporal (42 días)", "SBC ÷ 30 × 42 días",
            money(IBC_CASO / 30 * 42)],
           ["Indemnización por IPP (14 % de pérdida)", "14 % × 24 × SBC",
            money(0.14 * 24 * IBC_CASO)],
           [("Total reparación tarifada a cargo de la ARL", {"b": True}), "",
            (money(IBC_CASO / 30 * 42 + 0.14 * 24 * IBC_CASO), {"b": True})]],
          [6.2, 6.0, 3.8], size=9, align_centro=[2],
          titulo="Tabla 8. Liquidación de las prestaciones económicas del SGRL",
          fuente="Fuente: elaboración propia con base en el artículo 1 de la Ley 776 de 2002 (subsidio "
                 "por incapacidad temporal) y su artículo 5 (indemnización por incapacidad permanente "
                 "parcial, entre 1 y 24 salarios base de liquidación). El porcentaje de pérdida de "
                 "capacidad laboral es hipotético y en la práctica lo determina la ARL con el manual "
                 "único de calificación (Decreto 1507 de 2014).")
    rich(doc, ["La indemnización por incapacidad permanente parcial se calcula en proporción al daño, "
               "con un piso de un salario base de liquidación y un techo de veinticuatro, y se "
               "reajusta anualmente. Su reconocimiento no extingue la posibilidad de reclamar la "
               "indemnización plena de perjuicios cuando concurre culpa del empleador, pues se trata "
               "de títulos distintos: la primera deriva del riesgo profesional asegurado y la segunda "
               "de la responsabilidad contractual."])

    # ---------- 2.5 culpa patronal
    h(doc, 2, "Análisis de la culpa patronal (artículo 216 del C.S.T.)")
    imagen(doc, "assets/f4_ruta_at.png", 14.6,
           "Figura 3. Ruta de análisis de un evento de origen laboral: del nexo causal a la reparación "
           "plena. Fuente: elaboración propia con base en el Decreto 1295 de 1994, la Ley 1562 de "
           "2012, la Ley 776 de 2002, el Decreto 2943 de 2013 y el artículo 216 del C.S.T.")

    h(doc, 3, "Presupuestos de la responsabilidad patronal")
    rich(doc, ["El artículo 216 del Código Sustantivo del Trabajo dispone que ",
               ("«cuando exista culpa suficientemente comprobada del patrono en la ocurrencia del "
                "accidente de trabajo o en la enfermedad profesional, está obligado a la "
                "indemnización total y ordinaria por perjuicios, pero del monto de ella debe "
                "descontarse el valor de las prestaciones en dinero, pagadas en razón de las normas "
                "consagradas en este capítulo»", {"i": True}),
               ". La norma consagra una responsabilidad contractual subjetiva que exige tres "
               "presupuestos: el daño, la culpa del empleador y el nexo causal entre ambos."])
    rich(doc, ["La jurisprudencia de la Sala de Casación Laboral de la Corte Suprema de Justicia ha "
               "precisado que la culpa exigible es la ",
               ("culpa leve", {"b": True}),
               ", definida por el artículo 63 del Código Civil como el descuido ligero o la falta de "
               "diligencia y cuidado que los hombres emplean ordinariamente en sus negocios propios, "
               "por tratarse de un contrato bilateral y conmutativo en el que ambas partes obtienen "
               "beneficio. Esa graduación es decisiva: no se exige culpa grave ni dolo, sino la "
               "omisión de la diligencia ordinaria en la provisión de condiciones seguras de "
               "trabajo."])

    h(doc, 3, "Carga de la prueba y dinámica procesal")
    rich(doc, ["La regla general es que la carga de probar la culpa corresponde a quien la alega: el "
               "trabajador o sus causahabientes (art. 167 del Código General del Proceso). No "
               "obstante, la Corte Suprema ha admitido una dinámica probatoria: acreditada la "
               "ocurrencia del accidente y la omisión de un deber de seguridad, la carga se traslada "
               "al empleador, que debe demostrar diligencia y cuidado (suministro de elementos de protección, capacitación, instrucciones, mantenimiento y supervisión) para exonerarse. "
               "En palabras de la Corporación, «el empleador puede desligarse de ella demostrando "
               "diligencia y cuidado en la realización del trabajo»."])
    rich(doc, ["Ese traslado explica por qué el SG-SST no es un requisito formal sino un medio de "
               "prueba. La sentencia SL2480 de 2025 ilustra el otro extremo: la Sala negó la "
               "indemnización del artículo 216 porque el empleador acreditó la entrega de elementos "
               "de protección, la realización de inducciones y capacitaciones y la impartición de "
               "instrucciones, mientras que la conducta del trabajador (expresamente prohibida y advertida momentos antes del siniestro) configuró culpa exclusiva de la víctima y "
               "rompió el nexo causal. El resultado del proceso, por tanto, depende de la calidad de "
               "la prueba documental que el empleador haya construido antes del accidente."])

    h(doc, 3, "Valoración del caso concreto")
    tabla(doc, ["Deber de seguridad", "Norma o fuente del deber", "Evidencia en el caso", "Valoración"],
          [["Suministro de elementos de protección personal", "Art. 13, Ley 1562 de 2012; "
            "Res. 0312 de 2019", "No hay registro de entrega de guantes antifallo", "Incumplido"],
           ["Capacitación y certificación para tareas críticas", "Art. 13, Ley 1562 de 2012; "
            "Decreto 1072 de 2015", "Certificación vencida desde noviembre de 2025", "Incumplido"],
           ["Procedimientos de trabajo seguro actualizados", "Decreto 1072 de 2015; "
            "Res. 0312 de 2019", "Procedimiento de prueba hidrostática sin actualizar", "Incumplido"],
           ["Identificación de peligros y controles de ingeniería", "Decreto 1072 de 2015",
            "Banco de prueba sin dispositivo de anclaje; matriz desactualizada", "Incumplido"],
           ["Reporte oportuno del accidente a la ARL", "Art. 13, Ley 1562 de 2012", "Reportado "
            "5 días después", "Incumplido"],
           ["Investigación del accidente con participación del trabajador", "Res. 1401 de 2007",
            "Investigación tardía y sin participación", "Incumplido"],
           ["Afiliación y pago de aportes al SGRL", "Arts. 2 y 18, Decreto 1295 de 1994",
            "Afiliación vigente y aportes al día", "Cumplido"]],
          [4.4, 4.4, 4.2, 3.0], size=8.3,
          titulo="Tabla 9. Matriz de valoración de los deberes de seguridad")

    rich(doc, ["La conclusión es desfavorable para la empresa. De los siete deberes verificados, seis "
               "están incumplidos y la omisión es directamente causal: si el cilindro hubiera estado "
               "anclado y el trabajador hubiera portado el elemento de protección adecuado y contado "
               "con certificación vigente, el desprendimiento de la válvula no habría producido la "
               "lesión en la magnitud en que se produjo. No se configura culpa exclusiva de la "
               "víctima, porque el trabajador no ejecutó una maniobra prohibida ni advertida: realizó "
               "la tarea encomendada con los medios deficientes que la empresa le proporcionó."])
    caja(doc, "Conclusión sobre la culpa patronal",
         ["Existe culpa leve suficientemente comprobada del empleador en la ocurrencia del accidente "
          "AT-2026-014, derivada del incumplimiento de seis deberes de seguridad verificables en "
          "documentos. La empresa está expuesta a la indemnización total y ordinaria de perjuicios "
          "del artículo 216 del C.S.T., con descuento de las prestaciones pagadas por la ARL. El "
          "riesgo se agrava porque la investigación interna, al atribuir el hecho al «exceso de "
          "confianza» del trabajador sin prueba, puede convertirse en prueba en contra de la propia "
          "empresa en un eventual proceso."],
         tipo='alerta')

    h(doc, 3, "Cuantificación de la indemnización plena de perjuicios")
    rich(doc, ["La indemnización plena comprende el daño emergente, el lucro cesante consolidado y "
               "futuro y el daño moral, que se presume en el derecho colombiano cuando se afecta la "
               "integridad física o la vida. Para efectos académicos se proyecta una liquidación "
               "ilustrativa, advirtiendo que en un proceso real los rubros deben acreditarse con "
               "prueba pericial."])
    plena = 3200000 + 2700000 + 15000000 + 0.14 * 24 * IBC_CASO
    tarifada = IBC_CASO / 30 * 42 + 0.14 * 24 * IBC_CASO
    tabla(doc, ["Rubro", "Criterio de liquidación", "Valor"],
          [["Daño emergente", "Gastos de transporte, terapias complementarias y adecuación del "
            "puesto de trabajo", money(3200000)],
           ["Lucro cesante", "Diferencia salarial por reubicación (18 meses × %s)" % money(150000),
            money(2700000)],
           ["Perjuicio moral", "Afectación a la integridad física y a la vida de relación "
            "(estimación ilustrativa)", money(15000000)],
           ["Indemnización por IPP reconocida por la ARL", "14 % × 24 × SBC",
            money(0.14 * 24 * IBC_CASO)],
           [("Subtotal reparación plena", {"b": True}), "",
            (money(plena), {"b": True})],
           [("Menos: prestaciones en dinero pagadas por la ARL", {"b": True}),
            "Subsidio por IT + indemnización por IPP (art. 216, inciso final, C.S.T.)",
            "− " + money(tarifada)],
           [("Neto a cargo del empleador", {"b": True}), "",
            (money(plena - tarifada), {"b": True})]],
          [5.4, 6.6, 4.0], size=8.8, align_centro=[2],
          titulo="Tabla 10. Liquidación ilustrativa de la reparación plena (art. 216 C.S.T.)")
    imagen(doc, "assets/f5_reparacion.png", 14.6,
           "Figura 4. Comparación entre la reparación tarifada del SGRL y la reparación plena del "
           "artículo 216 del C.S.T. Fuente: elaboración propia.")
    rich(doc, ["El ejercicio permite dimensionar el impacto económico: la reparación plena supera en "
               "casi tres veces la suma que asume la ARL, y ese excedente sale del patrimonio de la "
               "empresa (que, tratándose de una persona natural, es el patrimonio del propio titular). "
               "La prescripción de la acción es de tres años contados desde la ocurrencia del hecho o "
               "desde la calificación de la pérdida de capacidad laboral, conforme al artículo 151 del "
               "C.P.T.S.S."])

    # ---------- 2.6 diagnóstico
    h(doc, 2, "Diagnóstico del eje y recomendaciones")
    tabla(doc, ["Hallazgo", "Consecuencia jurídica", "Riesgo", "Acción correctiva"],
          [["Clasificación del riesgo inconsistente con la actividad", "Cobro retroactivo de "
            "diferencias y discusión de cobertura ante un siniestro", "Medio",
            "Consulta formal a la ARL y reliquidación"],
           ["SG-SST sin evidencia documental", "Culpa patronal acreditada (art. 216 C.S.T.)",
            "Alto", "Actualizar matriz, procedimientos y registros; asignar responsable con funciones escritas"],
           ["Certificaciones vencidas para tareas críticas", "Nexo causal directo con el daño",
            "Alto", "Programa de recertificación y prohibición de tareas críticas sin certificado"],
           ["Reporte extemporáneo del accidente", "Sanción administrativa y pérdida de prueba",
            "Medio", "Protocolo de reporte dentro de las 24 horas siguientes"],
           ["Investigación del accidente deficiente", "Presunción en contra del empleador",
            "Alto", "Aplicar metodología de investigación con participación del trabajador"]],
          [4.4, 5.2, 1.7, 4.7], size=8.4, align_centro=[2],
          titulo="Tabla 11. Hallazgos y acciones del Eje II")
    salto_pagina(doc)
