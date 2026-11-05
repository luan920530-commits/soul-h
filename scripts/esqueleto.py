# -*- coding: utf-8 -*-
"""Documento reconstruido sobre el esqueleto entregado por el docente.

Estructura obligatoria: portada, introducción (objetivo general), información general de la
empresa, 1. Salud y cotizaciones, 2. Riesgos laborales, 3. Pensiones, 4. Documento jurídico,
conclusiones y referencias. Cada apartado del esqueleto queda resuelto, sin anexos ni
metodología independiente.
"""
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx_engine import (p, rich, h, bullet, tabla, salto_pagina, linea_horizontal,
                         sombrear, bordes_celda, set_col_widths,
                         AZUL, GRIS, GRIS_CLARO, HEX_AZUL, HEX_CELESTE, BLANCO, FUENTE_SANS)
from datos import (EMPRESA, ACTA, TRABAJADORES, NOMINA, IBC_TOTAL, SMLMV, AUX_TRANSPORTE,
                   APORTE_PENSION, APORTE_ARL, APORTE_CAJA, TARIFAS, money,
                   IBC_CASO, TARIFADA, REPARACION_PLENA, NETO_EMPLEADOR, CAPITAL_RAIS,
                   MESADA_RPM, DIF_MENSUAL, MESADA_RAIS_SIN_GARANTIA, IBL_P03,
                   TOPE_CUANTIA, ENTIDADES)

UNIVERSIDAD = "Universidad INCCA de Colombia"
ASIGNATURA = "Derecho de la Seguridad Social"
DOCENTE = "Rafael Felipe Lozano Álvarez"
AUTORA = "Julia Inés Buitrago García"
CC = "39626936"
NIT = "39626936-8"
RAZON = "EXTINTORES S C M"
CIUDAD = "Soacha (Cundinamarca)"
FECHA = "12 de junio de 2026"
EPS = ENTIDADES["eps"]
ARL = ENTIDADES["arl"]
CAJA = ENTIDADES["caja"]
AFP_RPM = ENTIDADES["afp_rpm"]
AFP_RAIS = ENTIDADES["afp_rais"]
WT01, WT02, WT03 = TRABAJADORES
SAL01, SAL02, SAL03 = WT01["ibc"], WT02["ibc"], WT03["ibc"]


# ------------------------------------------------------------------ portada
def portada(doc):
    p(doc, "", space_after=22)
    p(doc, "AUDITORÍA Y DIAGNÓSTICO DEL SISTEMA DE SEGURIDAD SOCIAL DE %s" % RAZON,
      size=20, bold=True, color=AZUL, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
    p(doc, "Salud · Riesgos laborales · Pensiones · Documento jurídico",
      size=11.5, italic=True, color=GRIS, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=16)
    linea_horizontal(doc, color=HEX_AZUL, sz=12, space_after=12)

    filas = [("Universidad", UNIVERSIDAD),
             ("Asignatura", ASIGNATURA),
             ("Docente", DOCENTE),
             ("Integrante", "%s · C.C. %s" % (AUTORA, CC)),
             ("Fecha", FECHA)]
    t = doc.add_table(rows=len(filas), cols=2)
    for i, (k, v) in enumerate(filas):
        c0, c1 = t.rows[i].cells
        c0.text = ""
        c0.paragraphs[0].add_run(k.upper()).bold = True
        c0.paragraphs[0].runs[0].font.size = Pt(9)
        c0.paragraphs[0].runs[0].font.name = FUENTE_SANS
        c0.paragraphs[0].runs[0].font.color.rgb = GRIS
        c1.text = ""
        c1.paragraphs[0].add_run(v).font.size = Pt(11)
        margenes = 90
        for c in (c0, c1):
            from docx_engine import margenes_celda
            margenes_celda(c, top=margenes, bottom=margenes, left=120, right=120)
    set_col_widths(t, [4.4, 11.6])
    p(doc, "", space_after=18)
    p(doc, "%s · C.C./NIT %s · Empresa privada · %s" % (RAZON, NIT, CIUDAD),
      size=10.5, color=GRIS, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    salto_pagina(doc)


# ------------------------------------------------- introducción y empresa
def introduccion(doc):
    h(doc, 1, """
INTRODUCCIÓN""")
    p(doc, "Objetivo general", bold=True, size=12, color=AZUL, space_before=6, space_after=4)
    rich(doc, ["Analizar el cumplimiento de las obligaciones del Sistema de Seguridad Social Integral "
               "en %s, empresa privada de %s dedicada al comercio y recarga de extintores, con el fin "
               "de identificar el estado de la afiliación y las cotizaciones en salud, la exposición "
               "al riesgo laboral y la eventual culpa patronal, la situación pensional de sus tres "
               "trabajadores y los mecanismos jurídicos con los que cuenta la empresa para "
               "controvertir las decisiones de las entidades del sistema." % (RAZON, CIUDAD)])

    rich(doc, ["El trabajo se apoya en dos fuentes. La primera es documental: planilla integrada de "
               "liquidación de aportes, certificados de afiliación y nómina de octubre de 2026. La "
               "segunda es el acta de inspección sanitaria ",
               ("028-050164-26", {"b": True}),
               " del 09/09/2026, levantada por la Secretaría de Salud de Soacha, que no es un "
               "documento interno de la empresa sino una actuación de autoridad y, por eso mismo, "
               "prueba preconstituida de lo que ocurría en el establecimiento esa fecha."])
    rich(doc, ["Conviene decirlo desde el principio, porque ordena todo lo que sigue: la empresa "
               "cumple lo básico. Los tres trabajadores están afiliados, la planilla está al día y la "
               "cotización se hace sobre el ingreso real. Los problemas aparecen en otro lugar, en la "
               "evidencia: hay obligaciones que se cumplen sin dejar rastro documental, y una "
               "obligación que no se cumple en absoluto, que es el sistema de gestión de seguridad y "
               "salud en el trabajo. En una empresa de tres personas eso se arregla en semanas; "
               "ignorarlo, en cambio, puede costar el patrimonio."])

    h(doc, 1, "INFORMACIÓN GENERAL DE LA EMPRESA")
    h(doc, 2, "Datos básicos")
    tabla(doc, ["Variable", "Dato"],
          [["Nombre de la empresa o entidad", "%s (persona natural comerciante, nombre comercial "
            "%s)" % (AUTORA, RAZON)],
           ["Naturaleza", "Privada (empresa unipersonal de persona natural)"],
           ["Actividad económica", "CIIU 4759: comercio al por menor de otros artículos domésticos "
            "en establecimientos especializados. En el mismo local se ejecutan recarga, "
            "mantenimiento y prueba hidrostática de extintores"],
           ["Ciudad", "Soacha (Cundinamarca), comuna 3, barrio La Despensa, carrera 4 n.° 60-17"],
           ["Número aproximado de trabajadores", "Tres (3) dependientes"],
           ["Nivel de riesgo laboral estimado", "Clase I (riesgo mínimo), código de actividad "
            "1475901, tarifa del 0,522 % del IBC. Es la clase declarada; el numeral 2.1 explica por "
            "qué es discutible"],
           ["Representante legal", "%s, C.C. %s" % (AUTORA, CC)],
           ["Afiliaciones", "%s (salud) · %s (riesgos laborales) · %s (caja) · %s para dos "
            "trabajadores y %s para uno" % (EPS, ARL, CAJA, AFP_RPM, AFP_RAIS)],
           ["Identificación tributaria", "C.C./NIT %s · inscripción %s" % (NIT, EMPRESA["inscripcion"])],
           ["Contacto", "Teléfono %s · %s" % (EMPRESA["telefono"], EMPRESA["correo"])]],
          [5.2, 10.8], size=8.8,
          titulo="Tabla 1. Datos básicos de la organización")

    h(doc, 2, "Composición de la planta")
    tabla(doc, ["Código", "Trabajador", "C.C.", "Cargo", "Salario base", "Régimen pensional"],
          [["WT-01", "Julia Inés Buitrago García", "39626936", "Gerente y titular", money(SAL01),
            "Prima Media (%s)" % AFP_RPM],
           ["WT-02", "Luis Antonio Aguilar Rodríguez", "1073694256", "Técnico de recarga",
            money(SAL02), "Ahorro Individual (%s)" % AFP_RAIS],
           ["WT-03", "Felipe Martínez", "8686834", "Vendedor", money(SAL03),
            "Prima Media (%s)" % AFP_RPM],
           [("Total", {"b": True}), "", "", "", (money(IBC_TOTAL), {"b": True}), ""]],
          [1.5, 5.0, 2.3, 2.9, 2.4, 3.6], size=8.6, align_centro=[0, 2, 4],
          titulo="Tabla 2. Trabajadores y salario base de cotización")


# ------------------------------------------------------- 1. salud
def salud(doc):
    salto_pagina(doc)
    h(doc, 1, "1. SALUD Y COTIZACIONES")

    h(doc, 2, "1.1 Descripción del sistema de salud de la empresa")
    rich(doc, ["Los tres trabajadores están afiliados al Régimen Contributivo del Sistema General de "
               "Seguridad Social en Salud a través de ",
               (EPS, {"b": True}),
               ", y la empresa los tiene reportados como cotizantes dependientes en la planilla "
               "integrada de liquidación de aportes. La afiliación es la que corresponde a un "
               "empleador formal: no hay trabajadores por fuera del sistema, ni contratistas "
               "independientes que hayan sustituido una relación laboral."])
    rich(doc, ["En salud, el empleador paga el 8,5 % del ingreso base de cotización y el trabajador "
               "el 4 %. Esa es la regla general. En este caso opera una excepción que conviene "
               "precisar porque cambia el monto: el artículo 114-1 del Estatuto Tributario exime del "
               "aporte patronal a las personas naturales que emplean a menos de diez trabajadores y "
               "que no superan las 10.000 UVT de ingresos. La condición que importa es la primera, y "
               "está redactada al revés de lo que suele creerse: la exoneración cobija a quien tiene ",
               ("dos o más trabajadores", {"b": True}),
               ". Con uno solo no aplica; con tres, sí."])
    rich(doc, ["El efecto práctico es que la empresa no paga el 8,5 % patronal ni los parafiscales de "
               "SENA e ICBF, pero sí sigue pagando el 4 % de la caja de compensación, que nunca está "
               "exonerado, y por supuesto la pensión y los riesgos laborales. El ahorro mensual es de ",
               ("$ 1.100.372", {"b": True}),
               ". Y hay un detalle que la empresa debería tener presente: depende de mantener al menos "
               "dos trabajadores. Si la planta baja a uno, la exoneración desaparece y hay que "
               "reliquidar."])

    h(doc, 2, "1.2 Ingreso Base de Cotización (IBC)")
    p(doc, "¿Qué salario se tomará como referencia?", bold=True, size=11, space_before=6, space_after=3)
    rich(doc, ["Se toma el salario base pactado con cada trabajador, sin incluir el auxilio de "
               "transporte. %s devenga %s; %s, %s; y %s, %s, equivalente a un salario mínimo. El IBC "
               "total de la planta es de %s mensuales."
               % ("Julia Inés", money(SAL01), "Luis Antonio", money(SAL02), "Felipe", money(SAL03),
                  money(IBC_TOTAL))])
    p(doc, "¿Cómo se calcula el IBC?", bold=True, size=11, space_before=6, space_after=3)
    rich(doc, ["El artículo 18 de la Ley 100 de 1993 define el IBC como el total de los factores "
               "salariales, más los pagos que retribuyen el trabajo, con dos límites: no puede ser "
               "inferior a un salario mínimo ni superior a veinticinco. Se excluyen los pagos no "
               "salariales del artículo 128 del Código Sustantivo del Trabajo, y en la práctica el "
               "más importante es el auxilio de transporte, que en 2026 es de %s. Ese es el error más "
               "frecuente en la nómina de las microempresas: sumarlo a la base y cotizar de más, o "
               "restarlo del salario y cotizar de menos." % money(AUX_TRANSPORTE)])
    tabla(doc, ["Concepto", "WT-01", "WT-02", "WT-03"],
          [["Salario base", money(SAL01), money(SAL02), money(SAL03)],
           ["Auxilio de transporte (no integra el IBC)", "No aplica",
            money(AUX_TRANSPORTE), money(AUX_TRANSPORTE)],
           ["Ingreso base de cotización", money(SAL01), money(SAL02), money(SAL03)],
           ["Aporte a pensión, 12 % a cargo del empleador", money(SAL01 * 0.12),
            money(SAL02 * 0.12), money(SAL03 * 0.12)],
           ["Aporte a riesgos laborales, clase I, 0,522 %", money(SAL01 * TARIFAS["I"]),
            money(SAL02 * TARIFAS["I"]), money(SAL03 * TARIFAS["I"])],
           ["Caja de compensación, 4 %", money(SAL01 * 0.04), money(SAL02 * 0.04),
            money(SAL03 * 0.04)],
           [("Total mensual a cargo del empleador", {"b": True}),
            (money(APORTE_PENSION + APORTE_ARL + APORTE_CAJA), {"b": True}), "", ""]],
          [6.4, 3.2, 3.2, 3.2], size=8.6, align_centro=[1, 2, 3],
          titulo="Tabla 3. Liquidación del IBC y de los aportes",
          fuente="Fuente: elaboración propia con la nómina de octubre de 2026 y las tarifas de 2026. "
                 "El aporte a salud patronal no se causa por la exoneración del artículo 114-1 del "
                 "E.T.; SENA e ICBF tampoco.")

    h(doc, 2, "1.3 EPS utilizadas")
    rich(doc, ["La EPS identificada es ", (EPS, {"b": True}),
               ". Para la planta completa se utiliza una sola entidad, lo que simplifica el reporte "
               "de novedades y la verificación de pagos. En el mercado local las alternativas con "
               "presencia en Soacha son las siguientes, y se enuncian porque el esqueleto las pide "
               "como «identificadas o posibles»:"])
    for t in ["EPS Sura: es la que efectivamente cubre a los tres trabajadores.",
              "Nueva EPS y Sanitas: alternativas de mercado, no utilizadas por la empresa.",
              "Compensar EPS: alternativa de mercado en la zona, no utilizada."]:
        bullet(doc, t)
    rich(doc, ["La empresa podría cambiar de EPS en cualquier momento, pero no debería hacerlo sin "
               "razón: cada cambio implica reportar novedad, verificar la continuidad de los "
               "tratamientos en curso y comprobar que no se interrumpa una incapacidad. Con tres "
               "trabajadores el trámite es sencillo; el riesgo es olvidarlo."])

    h(doc, 2, "1.4 Prestaciones reconocidas por la EPS")
    rich(doc, ["El plan obligatorio de salud cubre a los cotizantes y a su núcleo familiar. Aplicado "
               "a la planta de %s, el inventario es el siguiente:" % RAZON])
    tabla(doc, ["Prestación", "Reconocida", "Observación en el caso concreto"],
          [["Atención médica", "Sí", "Consultas de medicina general y especializada por la red de "
            "la EPS"],
           ["Incapacidades", "Sí", "Enfermedad general: 66,67 % del IBC a cargo de la EPS desde el "
            "tercer día; el empleador asume los dos primeros días"],
           ["Licencia de maternidad", "Sí", "100 % del salario durante 18 semanas; no se ha causado "
            "en el periodo auditado"],
           ["Licencia de paternidad", "Sí", "Reconocida a Felipe Martínez (WT-03) sobre una base "
            "inferior al 100 %% del IBC. Diferencia por reliquidar: %s" % money(412000)],
           ["Medicamentos", "Sí", "Incluidos en el plan de beneficios"],
           ["Hospitalización", "Sí", "Incluida, con las coberturas del plan"],
           ["Otras", "Sí", "Atención de urgencias y programas de promoción y prevención"]],
          [4.4, 2.2, 9.4], size=8.6, align_centro=[1],
          titulo="Tabla 4. Prestaciones de la EPS y su estado en la empresa")

    h(doc, 2, "1.5 Caso práctico")
    p(doc, "Situación", bold=True, size=11, space_before=6, space_after=3)
    rich(doc, ["Felipe Martínez (WT-03), vendedor, solicitó el reconocimiento de la licencia de "
               "paternidad por el nacimiento de su hijo. La empresa reportó la novedad, pero la "
               "variación de salario se cargó en la planilla por fuera del término, de modo que la "
               "EPS liquidó la prestación sobre una base inferior al 100 % del IBC."])
    p(doc, "Hechos", bold=True, size=11, space_before=6, space_after=3)
    for i, t in enumerate([
        "El trabajador informó del nacimiento y aportó el registro civil de nacimiento.",
        "La novedad de variación salarial se reportó con extemporaneidad, y la EPS tomó como base el "
        "IBC del periodo anterior.",
        "La diferencia entre lo reconocido y lo que correspondía asciende a %s." % money(412000),
        "La empresa pagó la licencia en la nómina, de modo que la suma quedó a su cargo y no fue "
        "recobrada ante la EPS.",
    ], 1):
        bullet(doc, t)
    p(doc, "Solución jurídica", bold=True, size=11, space_before=6, space_after=3)
    rich(doc, ["La licencia de paternidad es una prestación económica del sistema de salud y equivale "
               "al 100 % del IBC del periodo, conforme a los artículos 236 del Código Sustantivo del "
               "Trabajo y 6 de la Ley 2114 de 2021. Si la base se liquidó mal por una novedad "
               "tardía, el camino no es asumir la diferencia en silencio: se solicita la "
               "reliquidación ante la EPS, se corrige la novedad en la planilla y, si la entidad "
               "niega, se reclama y se agota la vía."])
    rich(doc, ["Nótese dónde está el error. No está en el derecho aplicable, que la empresa conoce y "
               "aplica: pagó la licencia. Está en el procedimiento, en la fecha del reporte. Es el "
               "patrón de todo el diagnóstico: se cumple, pero sin la prueba de haber cumplido a "
               "tiempo."])


# ------------------------------------------------- 2. riesgos laborales
def riesgos(doc):
    salto_pagina(doc)
    h(doc, 1, "2. RIESGOS LABORALES")

    h(doc, 2, "2.1 Identificación de la clase de riesgo")
    rich(doc, ["La actividad declarada ante el Registro Único Tributario y ante la Secretaría de Salud "
               "de Soacha es la ",
               ("CIIU 4759", {"b": True}),
               ", comercio al por menor de otros artículos domésticos en establecimientos "
               "especializados. La tabla del Decreto 768 de 2022 ubica la subactividad 4759-01 en ",
               ("clase I", {"b": True}),
               ", con tarifa inicial del 0,522 % del IBC, y su descripción menciona expresamente el "
               "comercio de extintores. En el papel, la clasificación está bien puesta."])
    rich(doc, ["El problema aparece cuando se compara el papel con el local. En el mismo "
               "establecimiento se recargan cilindros a presión, se les hace prueba hidrostática y se "
               "manipulan agentes limpios; eso no es comercio, es una actividad industrial de riesgo "
               "mayor, y la propia acta sanitaria 028-050164-26 la observó en la visita del "
               "09/09/2026. Mantener la clase I abarata la cotización mensual y, al mismo tiempo, "
               "deja abierta la discusión sobre la cobertura el día del siniestro."])
    p(doc, "Justificación de la clase asignada y de su revisión", bold=True, size=11,
      space_before=6, space_after=3)
    for t in ["Clase declarada: I (mínimo), código de actividad 1475901, tarifa 0,522 % del IBC.",
              "Fundamento de la declaración: CIIU 4759 y descripción expresa del comercio de "
              "extintores en el Decreto 768 de 2022.",
              "Razón para revisarla: la recarga y la prueba hidrostática de cilindros presurizados "
              "no están descritas en la subactividad 4759-01.",
              "Consecuencia económica de un cambio a clase II: la tarifa pasa del 0,522 %% al "
              "1,044 %%, es decir, de %s a %s mensuales sobre el IBC actual."
              % (money(APORTE_ARL), money(IBC_TOTAL * TARIFAS["II"])),
              "Conducta recomendada: consulta escrita a la ARL, con la descripción del proceso, y "
              "decisión documentada. No reclasificar por miedo ni quedarse por comodidad."]:
        bullet(doc, t)

    h(doc, 2, "2.2 ARL")
    rich(doc, ["La administradora de riesgos laborales identificada es ",
               (ARL + ", código de actividad 1475901, clase de riesgo I", {"b": True}),
               ". El sistema se financia solo con el aporte del empleador, que en este caso es de %s "
               "mensuales. La afiliación está vigente y los aportes al día, lo cual es una buena "
               "noticia con una advertencia: en riesgos laborales, estar afiliado no es lo mismo que "
               "estar protegido. La ARL cubre el riesgo; no cubre la culpa del empleador, y esa "
               "diferencia es la que se examina en el numeral 2.4." % money(APORTE_ARL)])
    rich(doc, ["La cobertura es doble: prestaciones asistenciales (atención médica, quirúrgica, "
               "farmacéutica, hospitalización, prótesis, rehabilitación) y económicas (subsidio por "
               "incapacidad temporal, indemnización por incapacidad permanente parcial, pensión de "
               "invalidez, pensión de sobrevivientes y auxilio funerario). Las primeras las presta la "
               "red de la ARL. Las segundas se liquidan sobre el IBC."])

    h(doc, 2, "2.3 Caso de accidente laboral")
    p(doc, "Descripción de los hechos", bold=True, size=11, space_before=6, space_after=3)
    rich(doc, ["El 12 de agosto de 2026, Luis Antonio Aguilar Rodríguez (WT-02, C.C. 1073694256), "
               "técnico de recarga y único operario de la planta, ejecutaba la prueba hidrostática de "
               "un extintor de agente limpio con veinte años de servicio. Al accionar la válvula, el "
               "cilindro se la desprendió por corrosión y se la proyectó contra la mano derecha. "
               "Resultado: fractura conminuta de los dedos índice y medio y lesión del tendón flexor."])
    rich(doc, ["Cuatro condiciones estaban dadas antes del accidente y las cuatro son del empleador: "
               "no había guantes antifallo, el cilindro no estaba anclado al banco de prueba, el "
               "procedimiento de trabajo no se había actualizado y la certificación del trabajador "
               "para manipular cilindros presurizados estaba vencida desde noviembre de 2025."])
    rich(doc, ["El trabajador hizo la tarea que le pidieron, con los medios que le dieron. El reporte "
               "a la ARL se hizo el 17 de agosto, cinco días después."])
    p(doc, "Derechos del trabajador", bold=True, size=11, space_before=6, space_after=3)
    for t in ["Atención médica integral, sin copagos ni cuotas moderadoras, a cargo de la ARL.",
              "Subsidio por incapacidad temporal equivalente al 100 % del IBC, desde el día "
              "siguiente al accidente.",
              "Calificación del origen y del porcentaje de pérdida de capacidad laboral.",
              "Indemnización por incapacidad permanente parcial, si hay pérdida igual o superior al "
              "5 %.",
              "Reubicación laboral cuando exista disminución de capacidad.",
              "Indemnización plena de perjuicios si se acredita culpa del empleador."]:
        bullet(doc, t)
    p(doc, "Prestaciones económicas y asistenciales", bold=True, size=11, space_before=6,
      space_after=3)
    rich(doc, ["Con el salario base de cotización de %s y una incapacidad de 42 días, el subsidio "
               "asciende a %s. Si la calificación reconoce una pérdida del 14 %% de la capacidad "
               "laboral, la indemnización por incapacidad permanente parcial equivale a 14 %% × 24 × "
               "IBC, es decir %s. La reparación tarifada a cargo de la ARL suma %s."
               % (money(WT02["ibc"]), money(WT02["ibc"] / 30 * 42), money(0.14 * 24 * WT02["ibc"]),
                  money(TARIFADA))])
    tabla(doc, ["Concepto", "Fórmula", "Valor"],
          [["Salario base de cotización de WT-02", "Base del periodo", money(IBL_P03)],
           ["Subsidio por incapacidad temporal (42 días)", "IBC ÷ 30 × 42", money(WT02["ibc"] / 30 * 42)],
           ["Indemnización por IPP (14 % de pérdida)", "14 % × 24 × IBC", money(0.14 * 24 * WT02["ibc"])],
           [("Total de la reparación tarifada", {"b": True}), "",
            (money(TARIFADA), {"b": True})]],
          [7.6, 5.0, 3.4], size=8.8, align_centro=[2],
          titulo="Tabla 5. Liquidación de las prestaciones del SGRL")

    h(doc, 2, "2.4 Culpa patronal")
    p(doc, "¿Hubo incumplimiento del empleador?", bold=True, size=11, space_before=6, space_after=3)
    rich(doc, [("Sí.", {"b": True}),
               " De los siete deberes de seguridad verificados en el expediente, seis están "
               "incumplidos: entrega de elementos de protección, capacitación y certificación para "
               "tareas críticas, procedimientos actualizados, controles de ingeniería en el banco de "
               "prueba, reporte oportuno del accidente e investigación con participación del "
               "trabajador. El único cumplido es la afiliación y el pago de aportes."])
    p(doc, "Argumentación", bold=True, size=11, space_before=6, space_after=3)
    rich(doc, ["El artículo 216 del Código Sustantivo del Trabajo obliga al empleador a la "
               "«indemnización total y ordinaria por perjuicios» cuando exista culpa suficientemente "
               "comprobada en la ocurrencia del accidente, con descuento de lo que la ARL haya "
               "pagado. La culpa exigible es la leve, la del artículo 63 del Código Civil: no hace "
               "falta dolo ni culpa grave, basta la omisión de la diligencia ordinaria. Y la "
               "jurisprudencia de la Sala de Casación Laboral admite que, probada la omisión de un "
               "deber de seguridad, la carga de demostrar diligencia pasa al empleador."])
    rich(doc, ["Aquí el empleador no tiene con qué demostrarla, y además la omisión es causal: con el "
               "cilindro anclado, con los guantes adecuados y con la certificación vigente, el "
               "desprendimiento no habría producido esa lesión. Tampoco hay culpa exclusiva de la "
               "víctima, que es la única defensa viable en este escenario: el trabajador no hizo una "
               "maniobra prohibida, hizo la tarea que le pidieron con los medios que le dieron."])
    rich(doc, ["La prueba más pesada no viene de la auditoría sino de la autoridad sanitaria. El acta "
               "028-050164-26 consignó que ",
               ("«no se evidencia documento de sistemas de seguridad y salud en el trabajo»", {"b": True}),
               ", y lo consignó una funcionaria identificada, con número de acta y constancia de "
               "notificación. La ausencia del sistema no hay que demostrarla: está demostrada."])
    rich(doc, ["En dinero, la diferencia es esta: la reparación tarifada del sistema asciende a %s, "
               "mientras que la reparación plena, en liquidación ilustrativa que incluye daño "
               "emergente, lucro cesante y perjuicio moral, llega a %s. El excedente, %s, no lo paga "
               "una aseguradora: lo paga la empresa y, siendo persona natural, el patrimonio de la "
               "titular. Una condena de ese tamaño se come veinte años del ahorro que produce una "
               "tarifa baja de riesgos laborales."
               % (money(TARIFADA), money(REPARACION_PLENA), money(NETO_EMPLEADOR))])


# --------------------------------------------------- 3. pensiones
def pensiones(doc):
    salto_pagina(doc)
    h(doc, 1, "3. PENSIONES")

    h(doc, 2, "3.1 Situación pensional de los trabajadores")
    rich(doc, ["Dos trabajadores están en el Régimen Solidario de Prima Media con Prestación Definida, "
               "administrado por Colpensiones, y uno en el Régimen de Ahorro Individual con "
               "Solidaridad, administrado por una AFP privada. La distribución es la siguiente:"])
    p(doc, "Trabajadores afiliados a Colpensiones", bold=True, size=11, space_before=6, space_after=3)
    for t in ["Julia Inés Buitrago García (WT-01), gerente y titular de la empresa.",
              "Felipe Martínez (WT-03), vendedor.",
              "En ambos casos la historia laboral está pendiente de certificación con corte a la "
              "fecha, y esa verificación es la que determina si completan las semanas exigidas."]:
        bullet(doc, t)
    p(doc, "Trabajadores afiliados a AFP privadas", bold=True, size=11, space_before=6, space_after=3)
    for t in ["Luis Antonio Aguilar Rodríguez (WT-02), técnico de recarga, afiliado a Porvenir en el "
              "Régimen de Ahorro Individual desde 1997.",
              "Es el único caso con traslado de régimen y, por eso, el que se analiza en los "
              "numerales 3.2 y 3.3."]:
        bullet(doc, t)
    rich(doc, ["Sobre la reforma pensional conviene fijar la fecha, porque es la fuente de casi todos "
               "los errores recientes. La Ley 2381 de 2024, que reorganiza el sistema en cuatro "
               "pilares, entra en vigencia el ",
               ("1.º de abril de 2027", {"b": True}),
               ", según el calendario que fijó la Sentencia C-264 de 2026. Hasta entonces rige la Ley "
               "100 de 1993. Un informe que en 2026 liquide una pensión con las reglas de la Ley "
               "2381 está aplicando una ley que todavía no produce efectos."])

    h(doc, 2, "3.2 Caso práctico de pensión")
    tabla(doc, ["Dato del trabajador", "Contenido"],
          [["Nombre", "Luis Antonio Aguilar Rodríguez (WT-02)"],
           ["Edad", "56 años; cumple 57 en 2027"],
           ["Semanas cotizadas", "1.150"],
           ["Régimen y administradora", "Ahorro Individual con Solidaridad, Porvenir"],
           ["Capital acumulado", money(CAPITAL_RAIS)],
           ["Ingreso base de liquidación de referencia", money(IBL_P03)]],
          [6.4, 9.6], size=8.8, titulo="Tabla 6. Datos del caso pensional")
    p(doc, "Posibilidad de acceder a la pensión", bold=True, size=11, space_before=6, space_after=3)
    rich(doc, ["En prima media, la pensión de vejez exige 57 años si es mujer y 62 si es hombre, y "
               "1.300 semanas. Con 1.150 semanas, faltan 150, que a razón de 52 semanas por año son "
               "cerca de tres años de cotización. Es decir: el requisito de semanas no está cumplido "
               "hoy, pero es alcanzable antes de la edad si el trabajador sigue cotizando."])
    rich(doc, ["La comparación económica es la que decide el caso. En prima media, con una tasa de "
               "reemplazo del 65,5 %% sobre un IBL de %s, la mesada sería de %s. En ahorro individual, "
               "el capital de %s financiaría unos %s mensuales, por debajo del salario mínimo, de modo "
               "que el afiliado dependería del Fondo de Garantía de Pensión Mínima y recibiría %s. La "
               "diferencia es de %s mensuales, que a veinte años de disfrute suman %s."
               % (money(IBL_P03), money(MESADA_RPM), money(CAPITAL_RAIS),
                  money(MESADA_RAIS_SIN_GARANTIA), money(SMLMV), money(DIF_MENSUAL),
                  money(DIF_MENSUAL * 12 * 20))])
    tabla(doc, ["Escenario", "Cálculo", "Mesada mensual"],
          [["Prima Media (Colpensiones)", "%s × 65,5 %%" % money(IBL_P03), money(MESADA_RPM)],
           ["Ahorro Individual sin garantía", "%s ÷ (20 años × 12 meses)" % money(CAPITAL_RAIS),
            money(MESADA_RAIS_SIN_GARANTIA)],
           ["Ahorro Individual con garantía estatal", "Pensión mínima (1 SMMLV)", money(SMLMV)],
           [("Pérdida frente a prima media", {"b": True}), "",
            ("− " + money(DIF_MENSUAL), {"b": True})]],
          [7.2, 5.6, 3.2], size=8.8, align_centro=[2],
          titulo="Tabla 7. Comparación de escenarios pensionales")

    h(doc, 2, "3.3 Ineficacia del traslado pensional")
    p(doc, "Situación planteada", bold=True, size=11, space_before=6, space_after=3)
    rich(doc, ["Luis Antonio se trasladó al Régimen de Ahorro Individual en 1997, a los 27 años, "
               "después de una visita de un asesor comercial que le dijo que en el fondo privado se "
               "pensionaría con una mesada mucho mayor. Firmó un formulario que ya traía impresa la "
               "frase «me afilio libre y voluntariamente». No conserva copia, no recibió simulación y "
               "no suscribió acta de doble asesoría."])
    p(doc, "Análisis jurídico", bold=True, size=11, space_before=6, space_after=3)
    rich(doc, ["Aquí no hace falta adjetivar. El trabajador está en el fondo privado desde 1997 y "
               "nadie le explicó lo que firmaba."])
    rich(doc, ["La pregunta es concreta: ese traslado, ¿produjo efectos? La respuesta se construye con "
               "los criterios de la ",
               ("Sentencia SU-107 de 2024", {"b": True}),
               ", que unificó la materia. El deber de información existe desde la creación de las "
               "administradoras; el consentimiento del formulario no basta, porque se exige "
               "consentimiento informado; la carga de la prueba se reparte, de modo que el afiliado "
               "aporta indicios del déficit y la AFP debe demostrar que informó y asesoró; y no hace "
               "falta estar cerca de la pensión ni tener un derecho consolidado, porque el vicio está "
               "en el acto mismo de traslado. La Sala de Casación Laboral había dicho lo mismo con "
               "otras palabras: «es a la AFP a quien le corresponde acreditar el cumplimiento del "
               "deber de información […] pues exigir al afiliado una prueba de este alcance es un "
               "despropósito» (SL3179-2023)."])
    rich(doc, ["Aplicado al caso, el resultado es claro. No hay soporte de asesoría, la AFP no "
               "conserva simulación ni acta, y el único documento que existe es un formulario con "
               "cláusula preimpresa. Eso configura el déficit de información y, con él, la "
               "ineficacia: el traslado se tiene por no celebrado y el afiliado vuelve a prima media "
               "como si nunca se hubiera ido, con las 1.150 semanas computadas íntegramente. La AFP "
               "debe trasladar el capital, los rendimientos, las comisiones y las primas, sin "
               "descontar nada al afiliado."])
    rich(doc, ["Hay una precisión de tiempo que cambia la estrategia. La oportunidad administrativa "
               "de traslado del artículo 76 de la Ley 2381 de 2024 venció en julio de 2026, y la "
               "ventana de la Ley 1748 de 2014 se cerró hace años. Queda la vía judicial, y hoy es "
               "el único camino para volver a prima media antes del 1.º de abril de 2027. Después de "
               "esa fecha el cálculo se complica, porque cambia la base del régimen aplicable."])


# --------------------------------------------- 4. documento jurídico
def documento(doc):
    salto_pagina(doc)
    h(doc, 1, "4. DOCUMENTO JURÍDICO")

    h(doc, 2, "4.1 Tipo de documento y destinatario")
    rich(doc, ["El documento que corresponde es un ",
               ("derecho de petición", {"b": True}),
               " y no una reclamación administrativa, porque el destinatario es una entidad del "
               "sistema de seguridad social y no una entidad de la administración pública en sentido "
               "estricto. La distinción no es formal: la reclamación administrativa previa del "
               "artículo 6 del Código Procesal del Trabajo y de la Seguridad Social es requisito de "
               "procedibilidad frente a entidades públicas, y su omisión es insubsanable. Frente a "
               "una ARL, lo que se radica es un derecho de petición, y su función es fijar fecha "
               "cierta, interrumpir la prescripción y construir la prueba de la reclamación."])
    tabla(doc, ["Campo del esqueleto", "Contenido"],
          [["Tipo de documento", "Derecho de petición de interés particular"],
           ["Destinatario", "%s, oficina de atención al usuario" % ARL],
           ["Remitente", "%s, C.C. %s, representante legal de %s (NIT %s)"
            % (AUTORA, CC, RAZON, NIT)],
           ["Lugar y fecha", "%s, %s" % (CIUDAD, FECHA)],
           ["Notificaciones", "%s, %s · %s · %s"
            % (EMPRESA["direccion"], EMPRESA["municipio"], EMPRESA["correo"],
               EMPRESA["telefono"])]],
          [4.6, 11.4], size=8.8, titulo="Tabla 8. Identificación del documento")

    h(doc, 2, "4.2 Asunto, hechos y solicitudes")
    p(doc, "Asunto", bold=True, size=11, space_before=6, space_after=3)
    rich(doc, ["Solicitud de revisión de la clasificación del riesgo, reconocimiento de las "
               "prestaciones derivadas del accidente de trabajo del 12 de agosto de 2026 y "
               "expedición de copias."])
    p(doc, "Hechos", bold=True, size=11, space_before=6, space_after=3)
    for i, t in enumerate([
        "La empresa %s está afiliada a esa administradora con el código de actividad 1475901, clase "
        "de riesgo I, y tiene a su cargo tres trabajadores dependientes, todos afiliados al Sistema "
        "General de Riesgos Laborales." % RAZON,
        "La actividad declarada es la CIIU 4759, clasificada en clase I por el Decreto 768 de 2022, "
        "pero en el mismo establecimiento se ejecutan tareas de recarga y prueba hidrostática de "
        "cilindros presurizados, advertidas por la Secretaría de Salud de Soacha en el acta "
        "028-050164-26 del 09/09/2026.",
        "El 12 de agosto de 2026 el trabajador Luis Antonio Aguilar Rodríguez, C.C. 1073694256, "
        "sufrió un accidente de trabajo que le produjo fractura conminuta de los dedos índice y medio "
        "de la mano derecha y lesión del tendón flexor. El reporte se radicó el 17 de agosto de 2026.",
        "A la fecha esa administradora reconoció el subsidio por incapacidad temporal, pero no ha "
        "notificado decisión de fondo sobre la calificación de la pérdida de capacidad laboral ni "
        "sobre la indemnización por incapacidad permanente parcial.",
    ], 1):
        p(doc, "%d. %s" % (i, t), size=10.5, space_after=4)
    p(doc, "Solicitudes", bold=True, size=11, space_before=6, space_after=3)
    for i, t in enumerate([
        "Que se revise la clasificación de la actividad y, si corresponde, se ajuste la clase de "
        "riesgo previa evaluación del proceso de recarga y prueba hidrostática.",
        "Que se expida copia de la decisión de clasificación del riesgo y de los soportes técnicos "
        "en que se fundamentó.",
        "Que se confirme por escrito el reconocimiento y pago del subsidio por incapacidad temporal, "
        "con la liquidación aplicada.",
        "Que se practique la calificación del origen y del porcentaje de pérdida de capacidad laboral "
        "conforme al Decreto 1507 de 2014, y se notifique el resultado.",
        "Que se expida copia íntegra del expediente administrativo del evento.",
    ], 1):
        p(doc, "%d. %s" % (i, t), size=10.5, space_after=4)

    h(doc, 2, "4.3 Fundamentos jurídicos")
    for t in ["Artículos 23 y 29 de la Constitución Política: derecho de petición, debido proceso y "
              "respuesta de fondo.",
              "Ley 1755 de 2015, artículos 5, 14 y 16: término de quince días hábiles y efectos de la "
              "falta de respuesta.",
              "Decreto 1295 de 1994, artículos 5, 7 y 13, y Ley 1562 de 2012, artículo 13: "
              "obligaciones de la administradora y deberes del empleador.",
              "Ley 776 de 2002, artículos 1, 5 y 7: subsidio del 100 % del IBC e indemnización por "
              "incapacidad permanente parcial.",
              "Decreto 768 de 2022, artículos 2 a 4, y Decreto 1772 de 1994, artículo 13: "
              "clasificación de la actividad y tarifas por clase de riesgo.",
              "Decreto 1507 de 2014: manual único para la calificación de la pérdida de capacidad "
              "laboral."]:
        bullet(doc, t)

    h(doc, 2, "4.4 Firma")
    rich(doc, ["El escrito se suscribe por la representante legal, con indicación de la cédula, el "
               "carácter con que actúa y los datos de notificación. En el documento anexo al presente "
               "informe se incluye el modelo completo, listo para radicación, con los datos de la "
               "empresa ya diligenciados."])
    p(doc, "", space_after=14)
    linea_horizontal(doc, color="9AA4B1", sz=4, space_after=2)
    p(doc, AUTORA, size=11, bold=True, space_after=0)
    p(doc, "C.C. %s · Representante legal de %s · NIT %s" % (CC, RAZON, NIT), size=9, color=GRIS,
      space_after=1)
    p(doc, "%s · %s" % (EMPRESA["correo"], EMPRESA["telefono"]), size=9, color=GRIS, space_after=2)

    h(doc, 2, "4.5 Juez competente")
    p(doc, "Autoridad competente", bold=True, size=11, space_before=6, space_after=3)
    rich(doc, [("Juzgado Laboral del Circuito de Soacha", {"b": True}),
               ", en primera instancia. La competencia se fija por cuatro factores que se aplican en "
               "orden: materia, territorio, cuantía y función."])
    p(doc, "Justificación", bold=True, size=11, space_before=6, space_after=3)
    rich(doc, ["Por la materia, el artículo 2 del C.P.T.S.S. atribuye a la jurisdicción ordinaria "
               "laboral las controversias del sistema de seguridad social integral «cualquiera que "
               "sea la naturaleza de la relación jurídica y de los actos jurídicos que se "
               "controviertan», y ese giro resuelve casi todo el problema: prestaciones de la ARL, "
               "reliquidación de aportes y traslado pensional van al juez laboral, no al "
               "administrativo."])
    rich(doc, ["Por el territorio, el artículo 11 del C.P.T.S.S. consagra un fuero concurrente: el "
               "lugar del domicilio de la entidad demandada o el lugar donde se surtió la "
               "reclamación, a elección del demandante. Como la reclamación se radicó en Soacha, el "
               "circuito competente es el de Soacha y no el de Bogotá."])
    rich(doc, ["Por la cuantía, el artículo 12 del C.P.T.S.S., modificado por la Ley 1395 de 2010, "
               "traza la raya en veinte salarios mínimos, que para 2026 son %s. Las pretensiones del "
               "caso la superan con holgura: la diferencia de mesadas asciende a %s y la reparación "
               "plena a %s, de modo que el asunto corresponde al juez del circuito en única "
               "instancia." % (money(TOPE_CUANTIA), money(DIF_MENSUAL * 12 * 20),
                               money(REPARACION_PLENA))])
    rich(doc, ["Por la función, con el juzgado del circuito en primera instancia, la apelación va al "
               "Tribunal Superior del Distrito Judicial de Cundinamarca, Sala Laboral, y la casación "
               "a la Sala de Casación Laboral de la Corte Suprema cuando la cuantía lo permita."])


# ------------------------------------------- conclusiones y referencias
def conclusiones(doc):
    salto_pagina(doc)
    h(doc, 1, "CONCLUSIONES")
    rich(doc, ["Tres conclusiones, en orden de urgencia."])

    p(doc, "Conclusión 1 · la afiliación está bien y la evidencia no", bold=True, size=11,
      space_before=8, space_after=3)
    rich(doc, ["Los tres trabajadores están afiliados a %s, a %s y a %s, la planilla está al día y el "
               "IBC se liquida sobre el salario real, sin sumar el auxilio de transporte. Eso se "
               "cumple. Lo que falla es la prueba: la novedad de la licencia de paternidad se reportó "
               "tarde y dejó una diferencia de %s; la exoneración del artículo 114-1 del Estatuto "
               "Tributario aplica pero no hay acta que verifique el umbral; y el acta sanitaria "
               "028-050164-26 encontró en blanco los campos de trabajadores afiliados. En una "
               "empresa de tres personas eso se corrige en una tarde de archivo, y es la corrección "
               "más rentable de todo el plan."
               % (EPS, ARL, CAJA, money(412000))])

    p(doc, "Conclusión 2 · el riesgo real está en el SG-SST, no en la tarifa", bold=True, size=11,
      space_before=8, space_after=3)
    rich(doc, ["La clase de riesgo I es defendible en el papel y discutible en el local, pero la "
               "discusión sobre la tarifa es un problema menor: pasar a clase II costaría %s "
               "mensuales más. El problema mayor es otro. No existe sistema de gestión de seguridad y "
               "salud en el trabajo, y la autoridad sanitaria ya lo dejó por escrito. Con seis de "
               "siete deberes incumplidos y un accidente que ocurrió, la empresa está expuesta a la "
               "indemnización plena del artículo 216 del C.S.T.: %s frente a los %s que asume la ARL, "
               "es decir %s que salen del patrimonio de la titular. Ahí está el argumento financiero "
               "para invertir en prevención antes del próximo accidente y no después."
               % (money(IBC_TOTAL * (TARIFAS["II"] - TARIFAS["I"])), money(REPARACION_PLENA),
                  money(TARIFADA),
                  money(NETO_EMPLEADOR))])

    p(doc, "Conclusión 3 · el caso pensional se gana por prueba, no por conveniencia", bold=True,
      size=11, space_before=8, space_after=3)
    rich(doc, ["Luis Antonio lleva 1.150 semanas en ahorro individual, con una mesada proyectada de "
               "%s frente a los %s que obtendría en prima media: una diferencia de %s mensuales y de "
               "%s en veinte años. La demanda de ineficacia del traslado no se sostiene en esa "
               "diferencia, que es una consecuencia, sino en un hecho simple: no existe acta de doble "
               "asesoría ni simulación entregada. La SU-107 de 2024 pone la carga de probar la "
               "asesoría en la administradora, y si no la acredita, el traslado no produce efectos. "
               "El plazo apremia: la reforma de la Ley 2381 de 2024 rige desde el 1.º de abril de "
               "2027, y después de esa fecha el cálculo cambia."
               % (money(MESADA_RAIS_SIN_GARANTIA), money(MESADA_RPM), money(DIF_MENSUAL),
                  money(DIF_MENSUAL * 12 * 20))])


def referencias(doc):
    h(doc, 1, "REFERENCIAS BIBLIOGRÁFICAS")
    refs = [
        "Congreso de la República de Colombia. Ley 100 de 1993, por la cual se crea el Sistema de "
        "Seguridad Social Integral. Diario Oficial n.° 41.148, 23 de diciembre de 1993.",
        "Congreso de la República de Colombia. Ley 1562 de 2012, por la cual se modifica el Sistema "
        "de Riesgos Laborales. Diario Oficial n.° 48.488, 11 de julio de 2012.",
        "Congreso de la República de Colombia. Ley 2381 de 2024, por medio de la cual se establece la "
        "reforma pensional. Diario Oficial n.° 52.825, 16 de julio de 2024.",
        "Corte Constitucional de Colombia. Sentencia SU-107 de 2024 (M.P. por designar). Bogotá, 9 "
        "de abril de 2024.",
        "Corte Suprema de Justicia, Sala de Casación Laboral. Sentencias SL3179-2023 y SL2480 de "
        "2025. Bogotá.",
        "Presidencia de la República de Colombia. Decreto 768 de 2022, por el cual se adiciona el "
        "Decreto 1072 de 2015 en materia de clasificación de actividades económicas. Bogotá, 2022.",
        "Secretaría de Salud de Soacha. Acta de inspección, vigilancia y control sanitario "
        "028-050164-26, del 9 de septiembre de 2026.",
    ]
    for i, r in enumerate(refs, 1):
        p(doc, "%d. %s" % (i, r), size=10, space_after=6)


def construir(doc):
    portada(doc)
    introduccion(doc)
    salud(doc)
    riesgos(doc)
    pensiones(doc)
    documento(doc)
    conclusiones(doc)
    referencias(doc)
