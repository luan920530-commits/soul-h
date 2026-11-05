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
               "en %s, empresa privada de %s dedicada al comercio y la recarga de extintores, para "
               "establecer cómo están la afiliación y las cotizaciones en salud, qué exposición al "
               "riesgo laboral tiene la planta y si hay culpa patronal, en qué situación pensional "
               "están los tres trabajadores y con qué mecanismos jurídicos cuenta la empresa para "
               "controvertir lo que decidan las entidades." % (RAZON, CIUDAD)])

    rich(doc, ["Trabajé con dos fuentes. La primera son los papeles de la empresa: la planilla "
               "integrada, los certificados de afiliación y la nómina de octubre de 2026. La segunda "
               "es el acta ",
               ("028-050164-26", {"b": True}),
               ", del 9 de septiembre de 2026, que levantó la Secretaría de Salud de Soacha. Esa acta "
               "no la escribió la empresa. La escribió una funcionaria con nombre, cédula y contrato, "
               "y por eso vale más que cualquier informe interno: dice qué había en el local ese día."])

    rich(doc, ["Lo primero que hay que decir es que la empresa cumple. Los tres trabajadores están "
               "afiliados, la planilla está al día y se cotiza sobre el salario real. Si el trabajo "
               "terminara aquí, no habría nada que corregir."])

    rich(doc, ["Pero no termina aquí. Hay obligaciones que se cumplen y no dejan rastro, y hay una "
               "que no se cumple en absoluto: el sistema de gestión de seguridad y salud en el "
               "trabajo."])
    rich(doc, ["Con tres trabajadores, lo primero se arregla en una tarde de archivo. Lo segundo es "
               "distinto, porque ya está probado por una autoridad y eso no se borra reponiendo el "
               "sistema."])

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
    rich(doc, ["Los tres trabajadores cotizan al Régimen Contributivo en salud a través de ",
               (EPS, {"b": True}),
               " y aparecen como cotizantes dependientes en la planilla integrada. No hay nadie por "
               "fuera del sistema. Tampoco hay contratistas independientes que en la práctica hayan "
               "reemplazado una relación laboral, que es el problema que suele encontrarse en "
               "empresas de este tamaño."])
    rich(doc, ["La regla general en salud es 8,5 % a cargo del empleador y 4 % a cargo del "
               "trabajador. Aquí no se paga el 8,5 %, y no es un error de la empresa: el artículo "
               "114-1 del Estatuto Tributario exime del aporte patronal a las personas naturales que "
               "emplean a menos de diez trabajadores."])
    rich(doc, ["La condición está redactada al revés de lo que todo el mundo cree. La exoneración "
               "cobija a quien tiene ",
               ("dos o más trabajadores", {"b": True}),
               ". Con uno no aplica. Con tres, sí."])
    rich(doc, ["El ahorro es de ",
               ("$ 1.100.372", {"b": True}),
               " al mes. De ese beneficio depende de mantener la planta: si la empresa se queda con un "
               "solo trabajador, pierde la exoneración y toca reliquidar. Es una forma curiosa de "
               "premiar la contratación formal, pero así está la norma y hay que saberlo antes de "
               "despedir a alguien para ahorrar."],
         space_after=8)
    rich(doc, ["Queda un detalle que no es menor. La caja de compensación, el 4 %, se paga siempre: "
               "ese rubro nunca está exonerado. Y la pensión y los riesgos laborales tampoco."])

    h(doc, 2, "1.2 Ingreso Base de Cotización (IBC)")
    p(doc, "¿Qué salario se tomará como referencia?", bold=True, size=11, space_before=6, space_after=3)
    rich(doc, ["El salario base pactado, sin auxilio de transporte. Julia Inés devenga %s; Luis "
               "Antonio, %s; y Felipe, %s, que es un mínimo. Suman %s de IBC mensual."
               % (money(SAL01), money(SAL02), money(SAL03), money(IBC_TOTAL))])
    p(doc, "¿Cómo se calcula el IBC?", bold=True, size=11, space_before=6, space_after=3)
    rich(doc, ["El artículo 18 de la Ley 100 de 1993: todos los factores salariales, con dos límites "
               "(no menos de un mínimo, no más de veinticinco) y sin los pagos no salariales del "
               "artículo 128 del Código Sustantivo del Trabajo. El que más pesa es el auxilio de "
               "transporte, que en 2026 vale %s y no se cotiza. Ahí está el error clásico de la "
               "nómina pequeña: sumarlo a la base y pagar de más, o restarlo del salario y pagar de "
               "menos." % money(AUX_TRANSPORTE)])
    rich(doc, ["Las dos cosas salen caras. Una por plata, la otra por sanción."])
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
    rich(doc, ["La EPS es ", (EPS, {"b": True}),
               " y cubre a los tres. Una sola entidad para toda la planta simplifica el reporte de "
               "novedades y la verificación de pagos. Las demás opciones del mercado local son estas, "
               "que el esqueleto pide enunciar aunque no se usen:"])
    for t in ["Nueva EPS.",
              "Sanitas.",
              "Compensar EPS."]:
        bullet(doc, t)
    rich(doc, ["Cambiar de EPS no es difícil con tres trabajadores. El riesgo es hacerlo por "
               "descarte. Cada cambio obliga a reportar novedad, verificar que no se interrumpa una "
               "incapacidad en curso y confirmar que el trabajador siga en tratamiento. Si hay una "
               "incapacidad abierta, el cambio puede costar semanas de pago."])

    h(doc, 2, "1.4 Prestaciones reconocidas por la EPS")
    rich(doc, ["Las que se han causado en la planta son las del cuadro. La única con problema es la "
               "licencia de paternidad, y el problema no es de reconocimiento sino de base."])
    rich(doc, ["La EPS pagó. Pagó sobre una base que no era la que correspondía."])
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
    rich(doc, ["Felipe Martínez, el vendedor, tuvo un hijo y pidió la licencia de paternidad. La "
               "empresa reportó la novedad, pero la variación de salario entró a la planilla después "
               "del término y la EPS liquidó sobre la base del periodo anterior, que era más baja."])
    p(doc, "Hechos", bold=True, size=11, space_before=6, space_after=3)
    for i, t in enumerate([
        "El trabajador avisó del nacimiento y entregó el registro civil.",
        "La novedad de variación salarial se reportó tarde y la EPS tomó el IBC anterior.",
        "La diferencia entre lo reconocido y lo que correspondía es de %s." % money(412000),
        "La empresa pagó la licencia en la nómina, así que el faltante quedó de su bolsillo y nunca "
        "lo recobró.",
    ], 1):
        bullet(doc, t)
    p(doc, "Solución jurídica", bold=True, size=11, space_before=6, space_after=3)
    rich(doc, ["La licencia de paternidad equivale al 100 % del IBC del periodo, según los artículos "
               "236 del Código Sustantivo del Trabajo y 6 de la Ley 2114 de 2021. Si la base quedó "
               "mal por un reporte tardío, no se asume en silencio: se pide la reliquidación a la "
               "EPS, se corrige la novedad en la planilla y, si la entidad dice que no, se reclama."])
    rich(doc, ["El error no está en el derecho. La empresa conoce la norma y pagó la licencia. Está "
               "en la fecha del reporte. Y ese es el patrón de todo el diagnóstico: se cumple, pero "
               "no queda prueba de haber cumplido a tiempo."])


# ------------------------------------------------- 2. riesgos laborales
def riesgos(doc):
    salto_pagina(doc)
    h(doc, 1, "2. RIESGOS LABORALES")

    h(doc, 2, "2.1 Identificación de la clase de riesgo")
    rich(doc, ["Actividad económica: ",
               ("CIIU 4759", {"b": True}),
               ", comercio al por menor de otros artículos domésticos en establecimientos "
               "especializados, que es lo que la empresa declaró en el RUT y lo que reportó a la "
               "Secretaría de Salud de Soacha. La tabla del Decreto 768 de 2022 pone la subactividad "
               "4759-01 en clase I, tarifa inicial del 0,522 %, y su descripción menciona el comercio "
               "de extintores. Con ese papel en la mano, la clasificación está bien puesta."])
    p(doc, "Justificación", bold=True, size=11, space_before=6, space_after=3)
    rich(doc, ["El problema está en el local, no en el código. En el mismo establecimiento se "
               "recargan cilindros a presión, se les hace prueba hidrostática y se manipulan agentes "
               "limpios. Eso es una actividad industrial de riesgo mayor, no comercio, y la propia acta "
               "028-050164-26 lo observó en la visita del 9 de septiembre de 2026."])
    rich(doc, ["La discusión, entonces, es de hecho y no de tarifa."])
    rich(doc, ["Quedarse en clase I ahorra %s al mes. Pasar a clase II costaría %s más. Ninguna de "
               "las dos cifras decide nada por sí sola: lo que decide es qué pasó el día del "
               "siniestro y si la ARL va a discutir la cobertura."
               % (money(APORTE_ARL),
                  money(IBC_TOTAL * (TARIFAS["II"] - TARIFAS["I"])))])
    for t in ["Clase declarada: I, código de actividad 1475901, tarifa 0,522 % del IBC.",
              "Con qué se sostiene: CIIU 4759 y la mención expresa del comercio de extintores en el "
              "Decreto 768 de 2022.",
              "Por qué es discutible: la recarga y la prueba hidrostática no están descritas en la "
              "subactividad 4759-01.",
              "Qué hacer: consulta escrita a la ARL con la descripción del proceso, y decisión "
              "documentada. Ni reclasificar por miedo ni quedarse por comodidad."]:
        bullet(doc, t)

    h(doc, 2, "2.2 ARL")
    rich(doc, ["La ARL es ", (ARL + ", código 1475901, clase I", {"b": True}),
               ", y el aporte es de %s al mes, todo a cargo de la empresa: en riesgos laborales el "
               "trabajador no aporta nada. La afiliación está vigente y los pagos al día. Eso es "
               "bueno y, a la vez, engañoso, porque estar afiliado no es estar protegido. La ARL "
               "cubre el riesgo. No cubre la culpa del empleador. Esa diferencia es la que se ve en "
               "el numeral 2.4." % money(APORTE_ARL)])
    rich(doc, ["Estar afiliado es obligatorio. Estar protegido es otra cosa."])
    rich(doc, ["La cobertura tiene dos mitades que casi nadie separa bien."])
    rich(doc, ["La asistencial es la atención: médica, quirúrgica, farmacéutica, hospitalización, "
               "prótesis y rehabilitación, y la paga la red de la ARL. La económica es el subsidio por "
               "incapacidad temporal, la indemnización por incapacidad permanente parcial, las "
               "pensiones de invalidez y sobrevivientes y el auxilio funerario; todo eso se liquida "
               "sobre el IBC."])

    h(doc, 2, "2.3 Caso de accidente laboral")
    p(doc, "Descripción de los hechos", bold=True, size=11, space_before=6, space_after=3)
    rich(doc, ["El 12 de agosto de 2026, Luis Antonio Aguilar estaba haciendo la prueba hidrostática "
               "de un extintor de agente limpio con veinte años de servicio. Es el único técnico de "
               "recarga de la planta. Al accionar la válvula, el cilindro se la desprendió por "
               "corrosión y se la proyectó contra la mano derecha: fractura conminuta de los dedos "
               "índice y medio y lesión del tendón flexor."])
    rich(doc, ["Antes del accidente ya faltaban cuatro cosas, y las cuatro eran del empleador. No "
               "había guantes antifallo. El cilindro no estaba anclado al banco de prueba. El "
               "procedimiento no se había actualizado. Y la certificación del trabajador para "
               "manipular cilindros presurizados estaba vencida desde noviembre de 2025."])
    rich(doc, ["Él hizo la tarea que le pidieron con los medios que le dieron. El reporte a la ARL se "
               "radicó el 17 de agosto, cinco días después."])
    p(doc, "Derechos del trabajador", bold=True, size=11, space_before=6, space_after=3)
    for t in ["Atención médica integral, sin copagos ni cuotas moderadoras.",
              "Subsidio del 100 % del IBC desde el día siguiente al accidente.",
              "Calificación del origen y del porcentaje de pérdida de capacidad laboral.",
              "Indemnización por incapacidad permanente parcial si la pérdida llega al 5 %.",
              "Reubicación laboral si queda con capacidad disminuida.",
              "Indemnización plena de perjuicios si se prueba culpa del empleador."]:
        bullet(doc, t)
    p(doc, "Prestaciones económicas y asistenciales", bold=True, size=11, space_before=6,
      space_after=3)
    rich(doc, ["Con un salario base de %s y 42 días de incapacidad, el subsidio es de %s. Si la "
               "calificación reconoce un 14 %% de pérdida, la indemnización por incapacidad "
               "permanente parcial es 14 %% × 24 × IBC, o sea %s. Entre las dos, la reparación "
               "tarifada que paga la ARL llega a %s."
               % (money(WT02["ibc"]), money(WT02["ibc"] / 30 * 42), money(0.14 * 24 * WT02["ibc"]),
                  money(TARIFADA))])
    tabla(doc, ["Concepto", "Fórmula", "Valor"],
          [["Salario base de cotización de WT-02", "Base del periodo", money(WT02["ibc"])],
           ["Subsidio por incapacidad temporal (42 días)", "IBC ÷ 30 × 42",
            money(WT02["ibc"] / 30 * 42)],
           ["Indemnización por IPP (14 % de pérdida)", "14 % × 24 × IBC",
            money(0.14 * 24 * WT02["ibc"])],
           [("Total de la reparación tarifada", {"b": True}), "",
            (money(TARIFADA), {"b": True})]],
          [7.6, 5.0, 3.4], size=8.8, align_centro=[2],
          titulo="Tabla 5. Liquidación de las prestaciones del SGRL")

    h(doc, 2, "2.4 Culpa patronal")
    p(doc, "¿Hubo incumplimiento del empleador?", bold=True, size=11, space_before=6, space_after=3)
    rich(doc, [("Sí.", {"b": True}),
               " De siete deberes de seguridad revisados, seis están incumplidos: entrega de "
               "elementos de protección, capacitación y certificación, procedimientos actualizados, "
               "controles en el banco de prueba, reporte oportuno e investigación con el trabajador. "
               "El único que se cumple es la afiliación y el pago de aportes."])
    p(doc, "Argumentación", bold=True, size=11, space_before=6, space_after=3)
    rich(doc, ["El artículo 216 del Código Sustantivo del Trabajo obliga a la «indemnización total y "
               "ordinaria por perjuicios» cuando hay culpa suficientemente comprobada, y ordena "
               "descontar lo que la ARL ya pagó. La culpa que se exige es la leve, la del artículo 63 "
               "del Código Civil: no hace falta dolo ni culpa grave, basta con no haber tenido la "
               "diligencia ordinaria. Y la Sala de Casación Laboral acepta que, probada la omisión de "
               "un deber de seguridad, a quien le toca demostrar diligencia es al empleador."])
    rich(doc, ["Aquí no tiene con qué."])
    rich(doc, ["Y la omisión es causal: con el cilindro anclado, con los guantes y con la "
               "certificación vigente, el desprendimiento no habría producido esa "
               "lesión. Tampoco hay culpa exclusiva de la víctima, que es la única defensa que "
               "quedaría en pie: el trabajador no hizo una maniobra prohibida, hizo su trabajo."])
    rich(doc, ["La prueba más difícil de rebatir no sale de este informe. Sale del acta "
               "028-050164-26, donde una funcionaria identificada, con número de acta y constancia "
               "de notificación, escribió que ",
               ("«no se evidencia documento de sistemas de seguridad y salud en el trabajo»", {"b": True}),
               ". La ausencia del sistema no hay que demostrarla. Está demostrada, con fecha."])
    rich(doc, ["En plata: la reparación tarifada son %s y la reparación plena, en la liquidación "
               "ilustrativa que incluye daño emergente, lucro cesante y perjuicio moral, llega a %s. "
               "El excedente son %s que no paga ninguna aseguradora. Los paga la empresa y, como es "
               "persona natural, los paga el patrimonio de la titular. Esa cifra equivale a más de "
               "veinte años del ahorro que produce una tarifa baja de riesgos laborales."
               % (money(TARIFADA), money(REPARACION_PLENA), money(NETO_EMPLEADOR))])


# --------------------------------------------------- 3. pensiones
def pensiones(doc):
    salto_pagina(doc)
    h(doc, 1, "3. PENSIONES")

    h(doc, 2, "3.1 Situación pensional de los trabajadores")
    rich(doc, ["Dos están en prima media y uno en ahorro individual. Esa es toda la distribución."])
    rich(doc, ["Con tres personas no tiene sentido hablar de promedios: cada historia laboral es un "
               "tercio de la planta, y una semana mal reportada cambia el régimen de alguien."])
    p(doc, "Trabajadores afiliados a Colpensiones", bold=True, size=11, space_before=6, space_after=3)
    for t in ["Julia Inés Buitrago García (WT-01), gerente y titular.",
              "Felipe Martínez (WT-03), vendedor.",
              "Falta certificar las dos historias laborales con corte a la fecha, que es lo que "
              "define si completan las semanas."]:
        bullet(doc, t)
    p(doc, "Trabajadores afiliados a AFP privadas", bold=True, size=11, space_before=6, space_after=3)
    for t in ["Luis Antonio Aguilar Rodríguez (WT-02), técnico de recarga, en Porvenir desde 1997.",
              "Es el único con traslado de régimen y, por eso, el único que se analiza en 3.2 y 3.3."]:
        bullet(doc, t)
    rich(doc, ["Sobre la reforma pensional hay que fijar una fecha, porque de ahí salen casi todos los "
               "errores que circulan. La Ley 2381 de 2024, la de los cuatro pilares, rige desde el ",
               ("1.º de abril de 2027", {"b": True}),
               ", según el calendario de la Sentencia C-264 de 2026. Hasta ese día manda la Ley 100 "
               "de 1993. Un trabajo que en 2026 liquide una pensión con la Ley 2381 está aplicando "
               "una ley que todavía no existe para estos efectos."])

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
    rich(doc, ["En prima media la pensión de vejez pide 62 años a los hombres y 1.300 semanas. Con "
               "1.150 faltan 150, que son unos tres años de cotización. Hoy no cumple; si sigue "
               "cotizando, cumple antes de la edad. Ese no es el problema."])
    rich(doc, ["El problema es la mesada. En prima media, con 65,5 %% de tasa de reemplazo sobre un "
               "IBL de %s, serían %s. En ahorro individual, el capital de %s alcanza para unos %s al "
               "mes, por debajo del mínimo, así que el afiliado terminaría en el Fondo de Garantía de "
               "Pensión Mínima con %s. La diferencia son %s mensuales. A veinte años de disfrute, %s."
               % (money(IBL_P03), money(MESADA_RPM), money(CAPITAL_RAIS),
                  money(MESADA_RAIS_SIN_GARANTIA), money(SMLMV), money(DIF_MENSUAL),
                  money(DIF_MENSUAL * 12 * 20))])
    rich(doc, ["Esa diferencia es la consecuencia, no la causa. El caso no se gana porque el fondo "
               "pague poco. Se gana porque nadie le explicó lo que firmaba, y eso se ve en el numeral "
               "siguiente."])
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
    rich(doc, ["Luis Antonio se pasó al fondo privado en 1997, a los 27 años. Un asesor comercial lo "
               "visitó y le dijo que allá se pensionaría con una mesada mucho mayor. Firmó un "
               "formulario que ya traía impresa la frase «me afilio libre y voluntariamente». No "
               "tiene copia, nunca recibió una simulación y no firmó acta de doble asesoría."])
    rich(doc, ["Veintinueve años después, esa diferencia vale %s al mes." % money(DIF_MENSUAL)])
    p(doc, "Análisis jurídico", bold=True, size=11, space_before=6, space_after=3)
    rich(doc, ["La pregunta es una sola: ese traslado, ¿produjo efectos? Se responde con la ",
               ("Sentencia SU-107 de 2024", {"b": True}),
               ", que unificó la materia y dejó cuatro reglas. Primera: el deber de informar existe "
               "desde que se crearon las administradoras, y con los años se volvió más exigente. "
               "Segunda: el consentimiento del formulario no sirve, tiene que ser informado. Tercera: "
               "la carga se reparte, el afiliado aporta indicios y la AFP demuestra que informó y "
               "asesoró. Cuarta: no hace falta estar cerca de la pensión ni tener un derecho "
               "consolidado, porque el vicio está en el acto mismo de traslado."])
    rich(doc, ["La Corte Suprema había llegado antes al mismo sitio, y con mejores palabras: «es a la "
               "AFP a quien le corresponde acreditar el cumplimiento del deber de información […] "
               "pues exigir al afiliado una prueba de este alcance es un despropósito» "
               "(SL3179-2023). El que tiene los papeles es la administradora. Ese es el detalle que "
               "decide estos casos."])
    rich(doc, ["En el caso no hay soporte de asesoría, no hay simulación y el único documento es un "
               "formulario con cláusula preimpresa. Con eso el traslado es ineficaz: se tiene por no "
               "celebrado, el afiliado vuelve a prima media con las 1.150 semanas completas y el "
               "fondo traslada el capital, los rendimientos, las comisiones y las primas sin "
               "descontarle nada."])
    rich(doc, ["Falta un dato de calendario que cambia la estrategia. La oportunidad administrativa "
               "del artículo 76 de la Ley 2381 de 2024 venció en julio de 2026, y la ventana de la "
               "Ley 1748 de 2014 se cerró hace años."])
    rich(doc, ["Solo queda la vía judicial, y solo sirve hasta el 1.º de abril de 2027."])


# --------------------------------------------- 4. documento jurídico
def documento(doc):
    salto_pagina(doc)
    h(doc, 1, "4. DOCUMENTO JURÍDICO")

    h(doc, 2, "4.1 Tipo de documento y destinatario")
    rich(doc, ["Lo que corresponde aquí es un ",
               ("derecho de petición", {"b": True}),
               ", no una reclamación administrativa. La razón es quién lo recibe. La reclamación "
               "administrativa previa del artículo 6 del Código Procesal del Trabajo y de la "
               "Seguridad Social es requisito de procedibilidad contra entidades públicas, y si se "
               "omite no se sanea. Contra una ARL no aplica: se radica un derecho de petición, que "
               "sirve para tres cosas a la vez: fijar fecha cierta, interrumpir la prescripción y "
               "dejar armada la prueba de que uno reclamó."])
    rich(doc, ["Equivocarse en eso cuesta el caso. Si se radica una «reclamación administrativa "
               "previa» contra un particular, el escrito no suspende la prescripción. Si se demanda a "
               "una entidad pública sin reclamación previa, la demanda se rechaza. Por eso lo primero "
               "que hay que decidir es a quién se le escribe, no qué se alega."])
    rich(doc, ["Con tres trabajadores, un error de ese tipo cuesta el caso entero."])
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
    rich(doc, ["Firma la representante legal, con cédula, carácter y datos de notificación. El texto "
               "completo del escrito va en el archivo aparte que acompaña este trabajo."])
    p(doc, "", space_after=14)
    linea_horizontal(doc, color="9AA4B1", sz=4, space_after=2)
    p(doc, AUTORA, size=11, bold=True, space_after=0)
    p(doc, "C.C. %s · Representante legal de %s · NIT %s" % (CC, RAZON, NIT), size=9, color=GRIS,
      space_after=1)
    p(doc, "%s · %s" % (EMPRESA["correo"], EMPRESA["telefono"]), size=9, color=GRIS, space_after=2)

    h(doc, 2, "4.5 Juez competente")
    p(doc, "Autoridad competente", bold=True, size=11, space_before=6, space_after=3)
    rich(doc, [("Juzgado Laboral del Circuito de Soacha", {"b": True}),
               ", en primera instancia. La competencia se decide con cuatro factores, en este orden: "
               "materia, territorio, cuantía y función."])
    p(doc, "Justificación", bold=True, size=11, space_before=6, space_after=3)
    rich(doc, ["Por materia, el artículo 2 del C.P.T.S.S. le da a la jurisdicción laboral las "
               "controversias de seguridad social integral «cualquiera que sea la naturaleza de la "
               "relación jurídica y de los actos jurídicos que se controviertan». Ese «cualquiera que "
               "sea» resuelve casi todo: prestaciones de la ARL, reliquidación de aportes y traslado "
               "pensional van al juez laboral, no al administrativo."])
    rich(doc, ["Por territorio, el artículo 11 del mismo código consagra un fuero concurrente: el "
               "domicilio de la entidad demandada o el lugar donde se surtió la reclamación, y elige "
               "el demandante. Como la reclamación se radicó en Soacha, el circuito es el de Soacha "
               "y no el de Bogotá."])
    rich(doc, ["Por cuantía, el artículo 12, modificado por la Ley 1395 de 2010, traza la raya en "
               "veinte salarios mínimos, que para 2026 son %s. Las pretensiones la pasan de largo: "
               "la diferencia de mesadas llega a %s y la reparación plena a %s. Es asunto de juez del "
               "circuito en única instancia."
               % (money(TOPE_CUANTIA), money(DIF_MENSUAL * 12 * 20), money(REPARACION_PLENA))])
    rich(doc, ["Por función, la apelación va al Tribunal Superior del Distrito Judicial de "
               "Cundinamarca, Sala Laboral, y la casación a la Sala de Casación Laboral de la Corte "
               "Suprema, cuando la cuantía lo permita."])


# ------------------------------------------- conclusiones y referencias
def conclusiones(doc):
    salto_pagina(doc)
    h(doc, 1, "CONCLUSIONES")

    p(doc, "Conclusión 1 · la afiliación está bien; la prueba no", bold=True, size=11,
      space_before=8, space_after=3)
    rich(doc, ["Los tres trabajadores están afiliados a %s, a %s y a %s, la planilla está al día y el "
               "IBC se liquida sobre el salario real. Eso se cumple. Lo que falla es el soporte. La "
               "novedad de la licencia de paternidad se reportó tarde y dejó %s sin recobrar. La "
               "exoneración del artículo 114-1 del E.T. aplica, pero no hay acta que verifique el "
               "umbral. Y el acta sanitaria dejó en blanco los campos de trabajadores afiliados. Es "
               "la corrección más barata del trabajo: una tarde de archivo."
               % (EPS, ARL, CAJA, money(412000))])

    p(doc, "Conclusión 2 · el riesgo está en el SG-SST, no en la tarifa", bold=True, size=11,
      space_before=8, space_after=3)
    rich(doc, ["Que la clase I sea discutible es un problema menor: pasar a clase II costaría %s al "
               "mes. El problema serio es que no existe sistema de gestión de seguridad y salud en el "
               "trabajo, y no lo digo yo: lo escribió una autoridad sanitaria en un acta. Con seis de "
               "siete deberes incumplidos y un accidente que ya ocurrió, la exposición del artículo "
               "216 del C.S.T. es de %s, frente a los %s que paga la ARL. La diferencia, %s, sale del "
               "patrimonio de la titular. Ese es el argumento para invertir en prevención antes del "
               "próximo accidente."
               % (money(IBC_TOTAL * (TARIFAS["II"] - TARIFAS["I"])), money(REPARACION_PLENA),
                  money(TARIFADA), money(NETO_EMPLEADOR))])

    p(doc, "Conclusión 3 · el caso pensional se gana con papeles", bold=True, size=11,
      space_before=8, space_after=3)
    rich(doc, ["Luis Antonio tiene 1.150 semanas en ahorro individual, con una mesada proyectada de "
               "%s contra %s en prima media. La diferencia son %s al mes, %s en veinte años. Pero la "
               "demanda no se sostiene en esa cifra, que es la consecuencia. Se sostiene en que no "
               "hay acta de doble asesoría ni simulación entregada, y en que la SU-107 de 2024 pone "
               "la carga de probar la asesoría en el fondo. Si el fondo no la acredita, el traslado "
               "no produce efectos. Hay tiempo hasta el 1.º de abril de 2027, y no más."
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
