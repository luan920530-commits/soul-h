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
          titulo="Tabla 7. Clases de riesgo y tarifas de cotización al SGRL",
          align_centro=[0, 2, 3],
          fuente="Fuente: artículo 13 del Decreto 1772 de 1994 y artículo 18 del Decreto 1295 de 1994. "
                 "Dentro de cada clase la ARL puede fijar la tarifa entre un mínimo y un máximo, según "
                 "el índice de lesiones incapacitantes y el cumplimiento del SG-SST.")

    h(doc, 3, "Clasificación de EXTINTORES S C M")
    rich(doc, ["El acta de inspección sanitaria 028-050164-26 del 09/09/2026 registra el código CIIU ",
               ("4759", {"b": True}),
               " (comercio al por menor de otros artículos domésticos en establecimientos "
               "especializados), que es el mismo que aparece en el Registro Único Tributario. Ese "
               "código no es neutro para el sistema de riesgos laborales: la tabla del Decreto 768 de "
               "2022 ubica la subactividad 4759-01 en ",
               ("clase de riesgo I (riesgo mínimo), código ARL 1475901", {"b": True}),
               ", con tarifa inicial del 0,522 % del IBC, y su descripción incluye expresamente el "
               "«comercio de extintores». La clasificación de la empresa está, por tanto, bien "
               "asignada en el papel."])
    rich(doc, ["El problema jurídico aparece cuando se contrasta ese código con lo que ocurre en el "
               "establecimiento. La actividad de recarga, mantenimiento y prueba hidrostática de "
               "cilindros presurizados desborda el comercio y entra en el terreno del servicio "
               "técnico sobre equipos a presión, con riesgo de proyección de partes, contacto con agentes "
               "extinguentes y esfuerzo físico. La subactividad 4759-01 no la describe, y dentro del "
               "mismo CIIU 4759 las subactividades 4759-01 y 4759-02 corresponden a otras clases. El "
               "artículo 4 del Decreto 768 de 2022 permite clasificar por la actividad más afín "
               "cuando la principal no refleja el riesgo real, atendiendo al peligro, las materias "
               "primas, los medios de producción y los procesos; el artículo 5 obliga a cotizar por "
               "el valor inicial de la nueva clase cuando la clasificación se modifica. La "
               "consecuencia práctica es doble: si la ARL reclasifica, hay cobro retroactivo; y si "
               "ocurre un accidente en la recarga, la ARL puede discutir la cobertura por "
               "subdeclaración del riesgo."])
    rich(doc, ["La recomendación no es reclasificar a ciegas, porque el sobrecosto es real. Es "
               "documentar el proceso, consultar formalmente a la ARL y decidir con respaldo técnico. "
               "El siguiente cuadro compara los tres escenarios posibles sobre el IBC actual de %s."
               % money(IBC_TOTAL)])
    tabla(doc, ["Escenario", "Fundamento", "Tarifa", "Aporte mensual", "Aporte anual", "Diferencia "
                "anual frente al escenario vigente"],
          [["A. Vigente: CIIU 4759-01, clase I", "Lo registrado en el RUT y en el acta sanitaria",
            "0,522 %", money(IBC_TOTAL * TARIFAS["I"]), money(IBC_TOTAL * TARIFAS["I"] * 12),
            "$ 0"],
           ["B. Reclasificación a clase II", "Si la ARL considera que la actividad real es de riesgo "
            "bajo y no mínimo", "1,044 %", money(IBC_TOTAL * TARIFAS["II"]),
            money(IBC_TOTAL * TARIFAS["II"] * 12),
            "+ " + money((IBC_TOTAL * (TARIFAS["II"] - TARIFAS["I"])) * 12)],
           ["C. Clase III por actividad afín de mantenimiento", "Art. 4 del Decreto 768 de 2022: "
            "talleres de mantenimiento de maquinaria y equipo (CIIU 3312-01)", "2,436 %",
            money(IBC_TOTAL * TARIFAS["III"]), money(IBC_TOTAL * TARIFAS["III"] * 12),
            "+ " + money((IBC_TOTAL * (TARIFAS["III"] - TARIFAS["I"])) * 12)]],
          [3.4, 4.4, 1.4, 2.0, 2.0, 2.6], size=7.6, align_centro=[2, 3, 4, 5],
          titulo="Tabla 8. Escenarios de clasificación del riesgo y efecto económico sobre el IBC "
                 "actual",
          fuente="Fuente: elaboración propia con base en el anexo técnico del Decreto 768 de 2022 y "
                 "las tarifas del artículo 13 del Decreto 1772 de 1994. El escenario A es el que "
                 "corresponde al código registrado; los escenarios B y C son hipótesis de "
                 "reclasificación, no una recomendación de subdeclarar el riesgo.")
    caja(doc, "Cómo se decide",
         ["La decisión debe sustentarse en el proceso real y no en la conveniencia económica. "
          "Reclasificar hacia arriba encarece la operación pero blinda la cobertura; mantener la "
          "clase I abarata el costo pero deja abierta la discusión sobre el origen del siniestro. "
          "El camino intermedio, y el que se recomienda en el plan de mejoramiento, es radicar una "
          "consulta escrita ante la ARL con la descripción del proceso de recarga, la ficha técnica "
          "de los equipos y el protocolo de prueba hidrostática, y conservar la respuesta: ese "
          "documento es la prueba de la diligencia del empleador si después se discute la culpa.",
          [("Nivel de riesgo: alto. Responsable: gerencia. Plazo: 60 días.", {"b": True})]],
         tipo='alerta')

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
               ", lo que en principio facilita el cumplimiento de EXTINTORES S C M; sin embargo, la "
               "sencillez del estándar no reduce el deber de diligencia, porque la responsabilidad "
               "del artículo 216 del C.S.T. se mide por el resultado y no por el tamaño de la "
               "empresa. La empresa tiene tres trabajadores y, por tanto, está en el grupo más "
               "pequeño de exigencia: no puede alegar que el estándar le resulta desproporcionado."])
    rich(doc, ["Aquí la auditoría deja de ser teórica, porque existe una prueba documental externa. "
               "El acta 028-050164-26, levantada el 09/09/2026 por la Secretaría de Salud de Soacha, "
               "consignó en la revisión documental que ",
               ("«no se evidencia documento de sistemas de seguridad y salud en el trabajo»", {"b": True}),
               ", con la anotación expresa de que, en caso de no contar con él, debe notificarse a la "
               "Dirección Territorial del Ministerio del Trabajo. Esa frase no proviene de la "
               "auditoría: la escribió una autoridad sanitaria en una visita con número de acta, "
               "funcionaria identificada y constancia de notificación. Su valor probatorio en "
               "un eventual proceso por culpa patronal es alto, porque acredita que el empleador "
               "conocía el déficit y no lo corrigió."])
    tabla(doc, ["Estándar mínimo (Res. 0312 de 2019, Anexo III)",
                "Evidencia exigible", "Lo que encontró el acta 028-050164-26", "Valoración"],
          [["Responsable de diseñar e implementar el SG-SST", "Designación escrita con asignación "
            "de funciones y recursos", "No se evidencia documento del sistema de gestión", "No cumple"],
           ["Afiliación al SGSSS y al SGRL", "Certificados de afiliación y planilla PILA",
            "Los campos «N.° de empleados afiliados a riesgos profesionales» y «N.° de empleados "
            "afiliados al Sistema de Seguridad Social» quedaron en blanco", "Sin verificar"],
           ["Capacitación en SST e inducción", "Registros de asistencia y contenido temático",
            "«No se evidencia programa de capacitación» en prácticas higiénicas, identificación y "
            "reducción del riesgo, almacenamiento seguro de sustancias químicas y uso de EPP "
            "(arts. 12 a 17 de la Ley 55 de 1993)", "No cumple"],
           ["Elementos de protección personal", "Dotación entregada y registrada",
            "El ítem 4.8 se calificó como cumple, pero sin soporte documental de entrega", "Cumple "
            "parcialmente"],
           ["Identificación de peligros y evaluación de riesgos", "Matriz de peligros actualizada",
            "No se evidencia listado de inventario de insumos químicos ni hojas de seguridad "
            "(ítem 4.11)", "No cumple"],
           ["Plan de prevención y respuesta ante emergencias", "Plan de emergencias, brigada y "
            "simulacros", "No se evidencia plan de gestión del riesgo de desastres (art. 2.3.1.5.2.1 "
            "del Decreto 2157 de 2017); sin alarma de emergencia (ítem 4.6) ni señalización de "
            "evacuación (ítem 1.5)", "No cumple"],
           ["Reporte e investigación de incidentes y accidentes", "FURAT y planilla de reporte",
            "El acta no verificó este punto; el caso AT-2026-014 documenta un reporte con cinco días "
            "de retardo", "No cumple"]],
          [3.9, 3.4, 6.4, 2.2], size=7.5, header_size=7.2, align_centro=[3],
          titulo="Tabla 9. Estándares mínimos del SG-SST contrastados con el acta sanitaria",
          fuente="Fuente: elaboración propia con base en la Resolución 0312 de 2019 y el acta "
                 "028-050164-26 del 09/09/2026 (Secretaría de Salud de Soacha).")
    rich(doc, ["El acta añade un dato que las auditorías internas suelen pasar por alto: la "
               "calificación de la visita anterior, del 18/10/2024, fue de ",
               ("57,50 % de cumplimiento", {"b": True}),
               ", con concepto «favorable con requerimientos». Dos años después, los mismos "
               "requerimientos siguen sin resolverse y el acta nueva fija un plazo de 22 días "
               "hábiles que vencía el 09/10/2026. La reincidencia no es un detalle: convierte una "
               "observación administrativa en prueba de un patrón de incumplimiento, que es "
               "exactamente lo que el artículo 216 del C.S.T. exige valorar como culpa."])

    # ---------- 2.4 caso práctico
    h(doc, 2, "Caso práctico: accidente de trabajo AT-2026-014")
    h(doc, 3, "Hechos relevantes")
    rich(doc, ["El trabajador WT-02, Luis Antonio Aguilar Rodríguez (C.C. 1073694256), técnico de "
               "recarga con contrato a término indefinido y único operativo de la planta, devengando "
               "un salario base de cotización de %s, se encontraba el 12 de agosto de 2026 "
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
          [["Salario base de cotización (SBC)", "Base de cotización de los seis meses anteriores "
            "(WT-02)", money(IBC_CASO)],
           ["Incapacidad temporal (42 días)", "SBC ÷ 30 × 42 días",
            money(IBC_CASO / 30 * 42)],
           ["Indemnización por IPP (14 % de pérdida)", "14 % × 24 × SBC",
            money(0.14 * 24 * IBC_CASO)],
           [("Total reparación tarifada a cargo de la ARL", {"b": True}), "",
            (money(IBC_CASO / 30 * 42 + 0.14 * 24 * IBC_CASO), {"b": True})]],
          [6.2, 6.0, 3.8], size=9, align_centro=[2],
          titulo="Tabla 10. Liquidación de las prestaciones económicas del SGRL",
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
               "trabajador o sus causahabientes (art. 167 del Código General del Proceso). La Corte "
               "Suprema ha admitido, sin embargo, una dinámica probatoria: acreditada la "
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
          titulo="Tabla 11. Matriz de valoración de los deberes de seguridad")

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
          titulo="Tabla 12. Liquidación ilustrativa de la reparación plena (art. 216 C.S.T.)")
    imagen(doc, "assets/f5_reparacion.png", 14.6,
           "Figura 4. Comparación entre la reparación tarifada del SGRL y la reparación plena del "
           "artículo 216 del C.S.T. Fuente: elaboración propia.")
    rich(doc, ["El ejercicio permite dimensionar el impacto económico: la reparación plena supera en "
               "casi tres veces la suma que asume la ARL, y ese excedente sale del patrimonio de la "
               "empresa (que, tratándose de una persona natural, es el patrimonio del propio titular). "
               "La prescripción de la acción es de tres años contados desde la ocurrencia del hecho o "
               "desde la calificación de la pérdida de capacidad laboral, conforme al artículo 151 del "
               "C.P.T.S.S."])

    # ---------- 2.6 el acta sanitaria como prueba externa
    h(doc, 2, "El acta sanitaria 028-050164-26 como prueba del estado del SG-SST")
    rich(doc, ["Una auditoría interna puede ser discutida; una actuación administrativa no. El 09 de "
               "septiembre de 2026 la Secretaría de Salud de Soacha practicó al establecimiento una "
               "inspección, vigilancia y control sanitario que quedó consignada en el acta ",
               ("028-050164-26", {"b": True}),
               ", formato SL-FR-028 versión 1, suscrita por una ingeniera química con contrato "
               "3136-2026 y notificada al vendedor Felipe Martínez. Esa acta es el documento más "
               "valioso del expediente, porque reúne, en un solo instrumento público, la fecha, la "
               "autoridad, el hallazgo y el plazo."])
    rich(doc, ["El acta no evalúa seguridad social, pero evalúa las mismas condiciones materiales que "
               "el SG-SST está llamado a controlar. Por eso sus hallazgos se trasladan al análisis de "
               "culpa del artículo 216 del C.S.T.: allí donde el acta dice «no se evidencia», el "
               "empleador no puede sostener después que cumplió, salvo que aporte la prueba en "
               "contrario, que es precisamente la que no tiene."])
    tabla(doc, ["Hallazgo del acta", "Ítem del formato", "Norma incumplida",
                "Efecto en el SG-SST y en el Eje II"],
          [["«No se evidencia documento de sistemas de seguridad y salud en el trabajo»",
            "Revisión documental", "Art. 13 de la Ley 1562 de 2012; Res. 0312 de 2019",
            "Acredita la ausencia total del SG-SST, no un simple defecto documental"],
           ["«No se evidencia programa de capacitación» en prácticas higiénicas, identificación y "
            "reducción del riesgo, almacenamiento seguro de sustancias químicas y uso de EPP",
            "Revisión documental", "Arts. 12 a 17 de la Ley 55 de 1993",
            "Sustenta la falta de capacitación que el caso AT-2026-014 identifica como causa del daño"],
           ["«No se evidencia hojas de seguridad de sustancias químicas» y no hay inventario de "
            "insumos", "ítems 4.11 y revisión documental", "Art. 102 de la Ley 9 de 1979",
            "Impide demostrar que el trabajador conocía el riesgo del cilindro que manipulaba"],
           ["Toma corriente sin recubrimiento", "ítem 1.8",
            "Art. 117 de la Ley 9 de 1979", "Riesgo eléctrico no controlado en zona de trabajo"],
           ["No se evidencia extintor con carga vigente", "ítem 4.7",
            "Art. 52 de la Resolución 14861 de 1985",
            "Paradoja probatoria: la empresa comercializa extintores y no tenía uno cargado"],
           ["No se evidencia alarma de emergencia ni señalización de evacuación", "ítems 4.6, 1.5 y 4.3",
            "Art. 52 de la Resolución 14861 de 1985",
            "Sin plan de respuesta ante emergencias verificable"],
           ["No se evidencia botiquín completo de primeros auxilios", "ítem 4.9",
            "Art. 240 de la Ley 9 de 1979",
            "Refuerza la falta de preparación para atender el accidente que efectivamente ocurrió"]],
          [5.2, 2.4, 3.4, 4.9], size=7.6, header_size=7.2,
          titulo="Tabla 13. Hallazgos del acta sanitaria y su proyección sobre el SG-SST",
          fuente="Fuente: elaboración propia con base en el acta 028-050164-26 del 09/09/2026 "
                 "(Secretaría de Salud de Soacha) y las normas citadas en cada ítem.")
    rich(doc, ["El acta añade tres datos que cambian la estrategia de defensa. Primero, la visita "
               "anterior, del 18 de octubre de 2024, había calificado el establecimiento con ",
               ("57,50 % de cumplimiento", {"b": True}),
               " y concepto «favorable con requerimientos»: los mismos incumplimientos persisten dos "
               "años después. Segundo, el acta concede ",
               ("22 días hábiles", {"b": True}),
               " contados desde el 09/09/2026, de modo que el plazo venció el ",
               ("09/10/2026", {"b": True}),
               ". Tercero, el documento advierte que no constituye permiso de funcionamiento y remite "
               "a los requisitos del artículo 87 de la Ley 1801 de 2016. Un requerimiento vencido y no "
               "atendido deja de ser un hallazgo y pasa a ser una prueba de desatención."])
    caja(doc, "Lectura jurídica del acta",
         ["El acta 028-050164-26 no impone sanciones por sí misma, pero produce tres efectos. "
          "(i) Es prueba documental preconstituida del estado del SG-SST a una fecha cierta, "
          "oponible al empleador en un proceso por culpa patronal. (ii) Activa el deber de "
          "notificación a la Dirección Territorial del Ministerio del Trabajo cuando no existe "
          "documento del sistema de gestión, según la propia anotación del formato. (iii) Si el "
          "requerimiento se incumple, habilita las medidas sanitarias de seguridad y las sanciones "
          "de la Ley 9 de 1979 y de la Ley 1801 de 2016. La recomendación es responder dentro del "
          "término, con un plan de acción fechado y firmado, y conservar la radicación: ese documento "
          "es la única prueba de diligencia posterior al 09/09/2026.",
          [("Nivel de riesgo: alto. Responsable: gerencia. Plazo: inmediato.", {"b": True})]],
         tipo='alerta')

    # ---------- 2.7 diagnóstico
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
          titulo="Tabla 14. Hallazgos y acciones del Eje II")
    salto_pagina(doc)
