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
    rich(doc, ["El sistema de riesgos laborales viene de la Ley 100 de 1993, se organizó con el "
               "Decreto 1295 de 1994 y cambió de fondo con la Ley 1562 de 2012. Ese cambio no fue de "
               "nombre: se pasó de «riesgos profesionales» a «riesgos laborales», se amplió la "
               "definición de accidente de trabajo, apareció la categoría de enfermedad laboral y se "
               "volvió obligatoria la afiliación de los independientes. El desarrollo reglamentario "
               "está compilado en el Decreto 1072 de 2015 y los estándares mínimos del SG-SST en la "
               "Resolución 0312 de 2019."])
    rich(doc, ["Un detalle que conviene retener: este es el único de los tres sistemas que paga "
               "únicamente el empleador. La cotización es un porcentaje del IBC según la clase de "
               "riesgo (artículos 18 y 19 del Decreto 1295 de 1994), sin aporte del trabajador. Y "
               "cubre dos cosas distintas: prestaciones asistenciales, que son la atención médica "
               "integral, y prestaciones económicas, que compensan la pérdida de capacidad. Si la "
               "afiliación se omite, el empleador responde directo por lo que la ARL habría pagado, "
               "más las sanciones del artículo 13 de la Ley 1562 de 2012."])

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
    rich(doc, ["La distinción no es teórica. De ella depende quién paga (ARL o EPS), cuánto paga "
               "(100 % del salario base de cotización si el origen es laboral, 66,67 % si es "
               "enfermedad general) y qué régimen de responsabilidad se activa para el empleador. En "
               "un caso límite, la diferencia entre una calificación y otra son varios millones y un "
               "proceso."])

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
    rich(doc, ["Aquí hay una coincidencia que vale la pena, porque no es frecuente. El acta de "
               "inspección sanitaria 028-050164-26 registra el código CIIU ",
               ("4759", {"b": True}),
               " (comercio al por menor de otros artículos domésticos en establecimientos "
               "especializados), y es exactamente el mismo que aparece en el Registro Único "
               "Tributario. Dos documentos de origen distinto dicen lo mismo sobre la actividad. Para "
               "el sistema de riesgos laborales ese código tiene consecuencia: la tabla del Decreto "
               "768 de 2022 ubica la subactividad 4759-01 en ",
               ("clase de riesgo I (riesgo mínimo), código ARL 1475901", {"b": True}),
               ", con tarifa inicial del 0,522 % del IBC, y su descripción incluye expresamente el "
               "«comercio de extintores». Dicho de otro modo: la clasificación está bien puesta. En "
               "el papel, la empresa no tiene problema de clase de riesgo."])
    rich(doc, ["El problema jurídico empieza justo ahí, cuando uno contrasta el papel con lo que pasa "
               "en el local. La actividad de recarga, mantenimiento y prueba hidrostática de "
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
    rich(doc, ["Y ojo, la recomendación no es reclasificar por si acaso. El sobrecosto es real: pasar "
               "de clase I a clase II casi duplica la tarifa. Lo que corresponde es documentar el "
               "proceso, consultar por escrito a la ARL y decidir con respaldo técnico, no con "
               "intuición ni con conveniencia. El cuadro siguiente compara los tres escenarios sobre "
               "el IBC actual de %s." % money(IBC_TOTAL)])
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
    caja(doc, "Cómo se decide esto",
         ["Hay dos maneras de equivocarse. Una es subir de clase por miedo: encarece la operación y "
          "no resuelve nada, porque el riesgo no lo determina la tarifa que uno pague. La otra es "
          "quedarse en clase I por comodidad: abarata el costo mensual y deja abierta la discusión "
          "sobre el origen del siniestro el día que ocurra. Ninguna de las dos es una decisión "
          "jurídica.",
          "Lo que sí sirve es radicar una consulta escrita ante la ARL con la descripción del "
           "proceso de recarga, la ficha técnica de los equipos y el protocolo de prueba "
           "hidrostática, y guardar la respuesta. Ese papel es la prueba de diligencia si después "
           "alguien discute la culpa.",
          ("Nivel de riesgo: alto. Responsable: gerencia. Plazo: 60 días.", {"b": True})],
         tipo='alerta')

    # ---------- 2.3 SG-SST
    h(doc, 2, "Obligaciones del empleador en materia de seguridad y salud en el trabajo")
    rich(doc, ["El artículo 13 de la Ley 1562 de 2012, que reemplazó el artículo 13 del Decreto 1295 "
               "de 1994, es la lista de lo que el empleador debe hacer: afiliar a todos sus "
               "trabajadores al SGRL, implementar el SG-SST, reportar accidentes y enfermedades a la "
               "ARL, permitir la participación de los trabajadores y adoptar medidas de prevención. "
               "Y son deberes de todos los empleadores, sin importar el tamaño. La sanción la impone "
               "la Dirección Territorial del Ministerio del Trabajo."])
    rich(doc, ["La Resolución 0312 de 2019 sí graduó la exigencia. A las empresas de uno a diez "
               "trabajadores les pide ",
               ("siete estándares mínimos", {"b": True}),
               ". Se podría pensar que eso le facilita las cosas a EXTINTORES S C M, y en parte es "
               "cierto: nadie le está pidiendo un sistema de gestión de planta industrial. Pero el "
               "estándar liviano no reduce el deber de diligencia. La responsabilidad del artículo "
               "216 del C.S.T. se mide por el resultado, no por el número de empleados. Con tres "
               "trabajadores la empresa está en el grupo de menor exigencia, y por eso mismo no "
               "puede decir que el estándar le quedaba grande."])
    rich(doc, ["Y aquí la auditoría deja de ser teoría, porque hay prueba externa. El acta "
               "028-050164-26, levantada el 09/09/2026 por la Secretaría de Salud de Soacha, escribió "
               "en revisión documental: ",
               ("«no se evidencia documento de sistemas de seguridad y salud en el trabajo»", {"b": True}),
               ", y añadió que, de no tenerlo, debe notificarse a la Dirección Territorial del "
               "Ministerio del Trabajo. Esa frase no la escribió esta auditoría. La escribió una "
               "autoridad sanitaria, en una visita con número de acta, funcionaria identificada y "
               "constancia de notificación. En un proceso por culpa patronal vale mucho, porque "
               "prueba dos cosas: que el empleador sabía y que no corrigió."])
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
    rich(doc, ["Son las tres de la tarde de un día de agosto y hay un solo operario en la zona de "
               "recarga. Luis Antonio Aguilar Rodríguez (WT-02, C.C. 1073694256) tiene contrato a "
               "término indefinido, salario base de cotización de %s y es el único técnico de "
               "recarga de la planta. Está haciendo la prueba hidrostática de un extintor de agente "
               "limpio con veinte años de servicio." % money(IBC_CASO),
               " Al accionar la válvula, el cilindro se la desprende por corrosión y se la proyecta "
               "contra la mano derecha. Fractura conminuta de los dedos índice y medio, lesión del "
               "tendón flexor."])
    rich(doc, ["Cuatro condiciones estaban dadas antes de eso, y las cuatro son del empleador: no "
               "había guantes antifallo, el cilindro no estaba anclado al banco de prueba, el "
               "procedimiento de trabajo no se había actualizado y la certificación del trabajador "
               "para manipular cilindros presurizados estaba vencida desde noviembre de 2025. El "
               "trabajador hizo lo que sabía hacer con las condiciones que le dieron."])
    rich(doc, ["El reporte a la ARL y a la EPS se hizo cinco días después. La investigación del "
               "evento se elaboró cuatro semanas más tarde, sin el trabajador y sin los "
               "representantes del SG-SST, y concluyó que «el trabajador actuó con exceso de "
               "confianza». Esa conclusión, además de injusta, es la prueba de que la investigación "
               "se hizo para cerrar un expediente y no para encontrar la causa."])

    h(doc, 3, "Determinación del nexo causal")
    rich(doc, ["Calificar el origen laboral exige tres cosas: una lesión orgánica o perturbación "
               "funcional, su ocurrencia en relación con la actividad y la relación de causalidad "
               "entre el trabajo y el daño. Aquí están las tres, y no hace falta acudir a "
               "presunciones:"])
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
    rich(doc, ["Y aun si alguien quisiera discutir el origen, el artículo 3 de la Ley 1562 de 2012 y "
               "la favorabilidad del artículo 53 constitucional resuelven la duda a favor del "
               "trabajador cuando el hecho ocurre en el lugar y en el tiempo del trabajo. En este "
               "caso no hay margen: la duda no existe."])

    h(doc, 3, "Prestaciones asistenciales a cargo de la ARL")
    rich(doc, ["Las prestaciones asistenciales están en el artículo 5 del Decreto 1295 de 1994, y la "
               "lista es larga a propósito: asistencia médica, quirúrgica, terapéutica y "
               "farmacéutica; hospitalización; odontología; medicamentos; ayudas diagnósticas; "
               "prótesis y aparatos ortopédicos con su reparación; rehabilitación física y "
               "profesional; calzado y overoles; orientación ocupacional. Todo eso lo asume la ARL, "
               "íntegro, y con independencia de si el empleador tuvo o no culpa. Esa es la lógica de "
               "la «reparación tarifada»: la víctima no discute con nadie para que la atiendan. "
               "Además, si queda con capacidad disminuida, tiene derecho a la reubicación laboral "
               "(artículo 16 de la Ley 776 de 2002)."])

    h(doc, 3, "Prestaciones económicas y liquidación")
    rich(doc, ["Las económicas son cinco: subsidio por incapacidad temporal, indemnización por "
               "incapacidad permanente parcial, pensión de invalidez, pensión de sobrevivientes y "
               "auxilio funerario (artículo 7 del Decreto 1295 de 1994). La Ley 776 de 2002 redefinió "
               "porcentajes y montos. La diferencia más visible con el sistema de salud está en el "
               "subsidio por incapacidad temporal: aquí es del ",
               ("100 % del salario base de cotización", {"b": True}),
               " y corre desde el día siguiente al accidente (artículo 1). En enfermedad general es "
               "del 66,67 %, a cargo de la EPS, y desde el tercer día."])
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
    rich(doc, ["La indemnización por incapacidad permanente parcial va en proporción al daño, con "
               "piso de un salario base de liquidación y techo de veinticuatro, y se reajusta cada "
               "año. Cobrarla no cierra la puerta a la indemnización plena si hay culpa del "
               "empleador. Son títulos distintos: una viene del riesgo asegurado, la otra de la "
               "responsabilidad contractual. Y esa distinción es el corazón del Eje II."])

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
               ". Es una responsabilidad contractual subjetiva y pide tres cosas: daño, culpa del "
               "empleador y nexo causal entre los dos."])
    rich(doc, ["La Sala de Casación Laboral ha precisado que la culpa exigible es la ",
               ("culpa leve", {"b": True}),
               ", la del artículo 63 del Código Civil: el descuido ligero, la falta de diligencia "
               "que las personas emplean ordinariamente en sus propios negocios. La razón es que el "
               "contrato de trabajo es bilateral y conmutativo, y de él se benefician las dos "
               "partes. La graduación importa mucho: no hay que probar culpa grave ni dolo. Basta "
               "con acreditar que no se prestaron las condiciones seguras que cualquier empleador "
               "razonable presta."])

    h(doc, 3, "Carga de la prueba y dinámica procesal")
    rich(doc, ["En principio, quien alega la culpa la prueba: el trabajador o sus causahabientes "
               "(artículo 167 del Código General del Proceso). Pero la Corte Suprema ha admitido una "
               "dinámica distinta. Probado el accidente y la omisión de un deber de seguridad, la "
               "carga pasa al empleador, que solo se exonera demostrando diligencia y cuidado: "
               "elementos de protección entregados, capacitación dictada, instrucciones impartidas, "
               "mantenimiento al día, supervisión efectiva. En palabras de la Corporación, «el "
               "empleador puede desligarse de ella demostrando diligencia y cuidado en la "
               "realización del trabajo»."])
    rich(doc, ["Ahí está la explicación del SG-SST. Se construye antes del accidente, no después, y "
               "cuando está bien hecho funciona como medio de prueba. La sentencia "
               "SL2480 de 2025 muestra el otro desenlace. La Sala negó la indemnización porque el "
               "empleador probó entrega de elementos de protección, inducciones, capacitaciones e "
               "instrucciones, y porque el trabajador hizo algo expresamente prohibido y advertido "
               "minutos antes: culpa exclusiva de la víctima, nexo causal roto. Nótese la asimetría. "
               "El mismo artículo 216 absolvió a un empleador y condenaría a EXTINTORES S C M, y la "
               "diferencia no está en la gravedad del accidente sino en la carpeta de soportes."])

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

    rich(doc, ["El resultado no es favorable. De siete deberes verificados, seis están incumplidos, "
               "y la omisión es causal en sentido estricto: con el cilindro anclado, con los guantes "
               "adecuados y con la certificación vigente, el desprendimiento de la válvula no habría "
               "producido la lesión en esa magnitud. Tampoco hay culpa exclusiva de la víctima, que es "
               "la única defensa viable en este escenario. Luis Antonio no hizo una maniobra "
               "prohibida ni advertida: hizo la tarea que le pidieron, con los medios que le dieron."])
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
    rich(doc, ["La indemnización plena tiene tres componentes: daño emergente, lucro cesante "
               "consolidado y futuro, y daño moral, que en Colombia se presume cuando se afecta la "
               "integridad física o la vida. Lo que sigue es una proyección con fines académicos. En "
               "un proceso real cada rubro se prueba con peritaje, y esa diferencia es enorme: aquí "
               "hay una cifra; allá hay un dictamen."])
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
    rich(doc, ["Vale la pena mirar la proporción. La reparación plena triplica lo que asume la ARL, y "
               "ese excedente no lo paga una aseguradora: lo paga la empresa. Y como EXTINTORES S C M "
               "es persona natural, en la práctica lo paga el patrimonio de Julia Inés. Ahí está el "
               "argumento económico del Eje II, y es el que justifica invertir en SG-SST antes de que "
               "ocurra el accidente y no después. El plazo para reclamar es de tres años desde el "
               "hecho o desde la calificación de la pérdida de capacidad (artículo 151 del "
               "C.P.T.S.S.), así que el riesgo no prescribe rápido."])

    # ---------- 2.6 el acta sanitaria como prueba externa
    h(doc, 2, "El acta sanitaria 028-050164-26 como prueba del estado del SG-SST")
    rich(doc, ["Una auditoría interna se puede discutir. Una actuación administrativa, no. El 09 de "
               "septiembre de 2026 la Secretaría de Salud de Soacha hizo una inspección de vigilancia "
               "y control sanitario y la dejó en el acta ",
               ("028-050164-26", {"b": True}),
               ", formato SL-FR-028 versión 1, firmada por una ingeniera química con contrato "
               "3136-2026 y notificada a Felipe Martínez en su calidad de vendedor. En un solo "
               "documento están la fecha, la autoridad, el hallazgo y el plazo. Para efectos "
               "probatorios, eso vale más que veinte folletos internos."])
    rich(doc, ["El acta no evalúa seguridad social; evalúa condiciones materiales, que es justamente "
               "lo que el SG-SST debe controlar. Por eso sus hallazgos se trasladan sin esfuerzo al "
               "análisis de culpa del artículo 216. Donde el acta dice «no se evidencia», el "
               "empleador ya no puede decir que cumplió, salvo que muestre la prueba en contrario. Y "
               "esa prueba es la que no tiene."])
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
    rich(doc, ["Hay tres datos del acta que cambian la estrategia de defensa. El primero es la "
               "visita anterior, del 18 de octubre de 2024, que calificó el establecimiento con ",
               ("57,50 % de cumplimiento", {"b": True}),
               " y concepto «favorable con requerimientos». Han pasado dos años y los requerimientos "
               "siguen ahí, iguales. El segundo es el plazo: ",
               ("22 días hábiles", {"b": True}),
               " desde el 09/09/2026, o sea que venció el ",
               ("09/10/2026", {"b": True}),
               ". El tercero es que el acta advierte que no es permiso de funcionamiento y remite al "
               "artículo 87 de la Ley 1801 de 2016."])
    rich(doc, ["Los tres juntos describen un problema distinto al incumplimiento simple. Un "
               "requerimiento vencido y desatendido deja de ser un hallazgo técnico y se convierte en "
               "prueba de desatención. Para el artículo 216 del C.S.T., eso es exactamente lo que hay "
               "que evitar."])
    caja(doc, "Qué produce el acta, jurídicamente",
         ["Por sí misma no impone sanciones, pero hace tres cosas. Es prueba documental "
          "preconstituida del estado del SG-SST a una fecha cierta, y por tanto oponible al "
          "empleador en un proceso por culpa patronal. Activa el deber de notificar a la Dirección "
          "Territorial del Ministerio del Trabajo cuando no existe documento del sistema, y eso lo "
          "dice la propia anotación del formato. Y si el requerimiento se incumple, habilita las "
          "medidas sanitarias de seguridad y las sanciones de la Ley 9 de 1979 y de la Ley 1801 de "
          "2016. La salida es una sola: responder dentro del término con un plan de acción fechado y "
          "firmado, y guardar la radicación. Es la única prueba de diligencia posterior al "
          "09/09/2026 que la empresa va a poder mostrar.",
          ("Nivel de riesgo: alto. Responsable: gerencia. Plazo: inmediato.", {"b": True})],
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
