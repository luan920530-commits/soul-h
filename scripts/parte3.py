# -*- coding: utf-8 -*-
"""EJE III: Diagnóstico pensional e ineficacia del traslado."""
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx_engine import (p, rich, h, bullet, cita, tabla, caja, salto_pagina, imagen,
                         aplicar_numeracion, AZUL, VINO, GRIS, GRIS_CLARO, HEX_VINO)
from datos import SMLMV, money, num, PM, RAIS

IBL = 3200000
TASA = 0.655
MESADA_RPM = IBL * TASA
CAPITAL_RAIS = 268000000
MESADA_RAIS_SIN_GARANTIA = CAPITAL_RAIS / (20 * 12)
DIF = MESADA_RPM - SMLMV
DIF_MENSUAL = MESADA_RPM - MESADA_RAIS_SIN_GARANTIA


def eje_iii(doc):
    salto_pagina(doc)
    h(doc, 1, "EJE III · DIAGNÓSTICO PENSIONAL E INEFICACIA DEL TRASLADO")

    # ---------- 3.1
    h(doc, 2, "Marco normativo del Sistema General de Pensiones")
    rich(doc, ["El Sistema General de Pensiones (SGP) protege las contingencias de vejez, invalidez y "
               "muerte y se estructura, desde la Ley 100 de 1993, en dos regímenes que coexisten y "
               "son excluyentes entre sí: el ",
               ("Régimen Solidario de Prima Media con Prestación Definida", {"b": True}),
               ", administrado hoy por Colpensiones, en el que los aportes conforman un fondo común "
               "de reparto y la prestación se define por ley en función de la edad, las semanas y el "
               "ingreso base de liquidación; y el ",
               ("Régimen de Ahorro Individual con Solidaridad", {"b": True}),
               ", administrado por las sociedades administradoras de fondos de pensiones, en el que "
               "cada afiliado tiene una cuenta individual y la prestación depende del capital "
               "acumulado, de los rendimientos y del bono pensional."])
    rich(doc, ["La libre elección entre regímenes es una manifestación del derecho a la seguridad "
               "social (art. 48 C.P.) y solo se perfecciona cuando el afiliado ha recibido "
               "información suficiente. De esa premisa se deriva todo el régimen de responsabilidad "
               "de las administradoras y el remedio de la ineficacia del traslado, que se estudia en "
               "el numeral 3.4."])

    h(doc, 3, "Requisitos y cálculo de la prestación en cada régimen")
    tabla(doc, ["Variable", "Régimen de Prima Media (Colpensiones)", "Régimen de Ahorro Individual (AFP)"],
          [["Edad", "57 años (mujeres) y 62 años (hombres)", "57 y 62 años, o anticipada si el capital "
            "alcanza para financiar una mesada superior al 110 % de la mínima"],
           ["Semanas", "1.300 semanas (art. 9, Ley 797 de 2003)", "No exige semanas para la pensión, "
            "pero sí 1.150 semanas para acceder a la garantía de pensión mínima"],
           ["Base de cálculo", "Ingreso base de liquidación de los últimos 10 años (art. 21, Ley 100 "
            "de 1993)", "Capital acumulado, bono pensional y rendimientos"],
           ["Monto", "Tasa de reemplazo del 65,5 % del IBL, más 1,5 % por cada 50 semanas "
            "adicionales, hasta el 80 % (art. 34, Ley 100 de 1993)", "Renta vitalicia o retiro "
            "programado según el capital; si es insuficiente, garantía de pensión mínima"],
           ["Garantía estatal", "Pensión mínima y límites legales", "Fondo de Garantía de Pensión "
            "Mínima (arts. 65 y ss., Ley 100 de 1993)"],
           ["Riesgo asumido", "El sistema (reparto)", "El afiliado (mercado y longevidad)"]],
          [2.8, 6.6, 6.6], size=8.4,
          titulo="Tabla 12. Comparación de los dos regímenes del SGP")

    h(doc, 3, "La reforma pensional de la Ley 2381 de 2024 y su estado actual")
    rich(doc, ["La Ley 2381 de 2024 transforma el sistema en un esquema de ",
               ("cuatro pilares", {"b": True}),
               " (solidario, semicontributivo, contributivo y de ahorro voluntario) y establece que "
               "las cotizaciones hasta 2,3 salarios mínimos se dirijan al componente administrado por "
               "Colpensiones, mientras el excedente se canaliza al ahorro individual. Su entrada en "
               "vigencia fue inicialmente suspendida por la Corte Constitucional, que en la ",
               ("Sentencia C-264 de 2026", {"b": True}),
               " declaró exequible la mayor parte del articulado, devolvió varias disposiciones a la "
               "Cámara de Representantes para subsanar vicios de trámite y fijó como fecha de "
               "entrada en vigencia el ",
               ("1.º de abril de 2027", {"b": True}),
               ". Hasta esa fecha continúan aplicándose las reglas de la Ley 100 de 1993."])
    tabla(doc, ["Disposición de la Ley 2381 de 2024", "Contenido", "Estado tras la C-264 de 2026"],
          [["Estructura por pilares", "Solidario, semicontributivo, contributivo y ahorro voluntario",
            "Exequible; rige desde el 1.º de abril de 2027"],
           ["Régimen de transición (art. 75)", "Mujeres con 750 o más semanas y hombres con 900 o "
            "más semanas conservan íntegramente la Ley 100 de 1993",
            "Exequible; el conteo se verifica al 1.º de abril de 2027"],
           ["Requisito de semanas (art. 4 y ss.)", "1.300 semanas para hombres; reducción gradual para "
            "mujeres hasta 1.000 en 2036", "Exequible"],
           ["Beneficio de 50 semanas por hijo (art. 36)", "Reducción adicional de semanas por cada hijo",
            "Devuelto a la Cámara; no rige mientras no se subsane el trámite"],
           ["Oportunidad de traslado (art. 76)", "Dos años para trasladarse, previa doble asesoría",
            "Vencida en julio de 2026"]],
          [4.2, 7.4, 4.4], size=8.3,
          titulo="Tabla 13. Impacto de la reforma pensional en la organización")

    caja(doc, "Advertencia de vigencia",
         ["Cualquier afirmación sobre «pensiones» en Colombia debe precisar la fecha a la que se "
          "refiere. Hasta el 31 de marzo de 2027 rige la Ley 100 de 1993; desde el 1.º de abril de "
          "2027, y siempre que se complete la subsanación ordenada, rige la Ley 2381 de 2024. La "
          "auditoría de EXTINTORES SCM se practicó en el periodo de transición, por lo que combina "
          "ambos marcos y señala expresamente cuál aplica en cada caso. Este detalle es el que "
          "distingue un informe con rigor técnico de uno meramente descriptivo."],
         tipo='alerta')

    # ---------- 3.2
    h(doc, 2, "Mapa pensional de la organización")
    rich(doc, ["La distribución de los trabajadores entre los dos regímenes se obtuvo del certificado "
               "de afiliación y de la planilla integrada de liquidación de aportes. El resultado "
               "muestra una concentración en el Régimen de Prima Media (%d de %d trabajadores, "
               "equivalente al %.0f %%) y una presencia minoritaria en el Régimen de Ahorro "
               "Individual (%d trabajadores, %.0f %%), lo que se explica por la antigüedad de la "
               "planta y por el perfil salarial de la empresa." % (PM, PM + RAIS, 100 * PM / (PM + RAIS),
                                                                   RAIS, 100 * RAIS / (PM + RAIS))])
    imagen(doc, "assets/f3_mapa_pensional.png", 14.6,
           "Figura 5. Mapa pensional de EXTINTORES SCM. Fuente: elaboración propia con base en los "
           "certificados de afiliación y la planilla PILA del periodo auditado.")

    tabla(doc, ["Cód.", "Edad / sexo", "Régimen", "Semanas", "Situación frente a la reforma",
                "Alerta"],
          [["EXT-001", "48 / M", "Prima Media", "980", "Régimen de transición (más de 900 semanas)",
            "Baja"],
           ["EXT-002", "44 / M", "Prima Media", "870", "Entra al nuevo sistema: le faltan 30 semanas "
            "para la transición", "Media"],
           ["EXT-003", "56 / F", "Ahorro Individual", "1.150", "Régimen de transición (más de 750 "
            "semanas) · caso auditado P-03", "Alta"],
           ["EXT-004", "33 / F", "Prima Media", "410", "Sistema de pilares", "Baja"],
           ["EXT-005", "38 / M", "Prima Media", "520", "Sistema de pilares", "Baja"],
           ["EXT-006", "51 / M", "Prima Media", "1.290", "Transición (más de 900 semanas); a 10 "
            "semanas de la pensión", "Media"],
           ["EXT-007", "27 / F", "Ahorro Individual", "230", "Sistema de pilares", "Baja"],
           ["EXT-008", "41 / F", "Prima Media", "760", "Transición (más de 750 semanas)", "Baja"],
           ["EXT-009", "36 / M", "Prima Media", "480", "Sistema de pilares", "Baja"],
           ["EXT-010", "29 / M", "Ahorro Individual", "95", "Sistema de pilares", "Media"]],
          [1.7, 2.0, 3.0, 1.7, 5.2, 1.4], size=8.3, align_centro=[0, 1, 3, 5],
          titulo="Tabla 14. Distribución de los trabajadores por régimen y situación pensional",
          fuente="Fuente: elaboración propia con base en la historia laboral y los certificados de "
                 "afiliación. Las semanas se reportan con corte a 30 de septiembre de 2026.")

    rich(doc, ["Del mapa se derivan tres conclusiones de gestión. Primera, la empresa tiene cuatro "
               "trabajadores en el régimen de transición, cuyas condiciones pensionales dependen "
               "íntegramente de la Ley 100 de 1993 y que, por tanto, no se verán afectados por el "
               "cambio de modelo. Segunda, dos trabajadores están a muy pocas semanas del umbral "
               "(EXT-002 y EXT-006), de modo que la verificación de la historia laboral y el pago "
               "oportuno de aportes adquieren una relevancia extraordinaria: una semana no reportada "
               "puede cambiar el régimen aplicable. Tercera, el trabajador EXT-010 presenta "
               "desviaciones en el reporte de semanas que sugieren periodos sin cotización, lo que "
               "debe corregirse antes de abril de 2027."])

    # ---------- 3.3
    h(doc, 2, "El deber de información y la doble asesoría")
    rich(doc, ["El acto de traslado de régimen no es un acto de adhesión: es una decisión jurídica de "
               "largo plazo cuyos efectos el afiliado solo advierte décadas después. Por esa razón la "
               "jurisprudencia construyó un estándar de información que evolucionó en tres etapas:"])
    for t in [
        [("Deber de información (desde la creación de las AFP): ", {"b": True}),
         "la administradora debía suministrar información clara, cierta, comprensible y oportuna "
         "sobre las características, condiciones, beneficios, diferencias, riesgos y consecuencias "
         "del cambio de régimen."],
        [("Deber de asesoría y buen consejo: ", {"b": True}),
         "la información debía traducirse en una recomendación personalizada, comprensible para el "
         "afiliado, sobre la conveniencia o no del traslado."],
        [("Doble asesoría (Ley 1748 de 2014, art. 3, y Circular Externa 016 de 2016 de la "
          "Superintendencia Financiera): ", {"b": True}),
         "el traslado exige que el afiliado reciba asesoría de representantes de ambos regímenes y "
         "que la administradora conserve el acta y los soportes documentales de esa asesoría."],
    ]:
        bullet(doc, t)

    rich(doc, ["La ",
               ("Sentencia SU-107 de 2024", {"b": True}),
               " de la Corte Constitucional unificó la materia y fijó reglas que hoy son el eje de "
               "cualquier litigio pensional: (i) el deber de información es exigible desde la "
               "creación de las administradoras y su intensidad creció con el tiempo; (ii) el "
               "consentimiento consignado en el formulario de afiliación es insuficiente, porque se "
               "requiere un consentimiento informado; (iii) la carga de la prueba se equilibra: el "
               "afiliado debe aportar indicios del déficit de información y la AFP debe demostrar que "
               "informó y asesoró; y (iv) la ineficacia del traslado no exige que el afiliado esté "
               "próximo a pensionarse ni que tenga un derecho consolidado, porque el vicio se predica "
               "del acto mismo de traslado."])
    cita(doc, "«El simple consentimiento vertido en el formulario de afiliación es insuficiente, por "
              "la necesidad de un consentimiento informado […] la violación del deber de información "
              "se predica frente a la validez del acto jurídico de traslado, considerado en sí "
              "mismo.»",
         "Corte Constitucional, Sentencia SU-107 de 2024 (9 de abril de 2024)")

    rich(doc, ["La Sala de Casación Laboral de la Corte Suprema de Justicia había anticipado esa "
               "regla y sostiene que es la administradora (y no el afiliado) quien debe acreditar el "
               "cumplimiento del deber: «es a la AFP a quien le corresponde acreditar el cumplimiento "
               "del deber de información al momento del traslado de régimen pensional, pues exigir al "
               "afiliado una prueba de este alcance es un despropósito» (CSJ, SL3179-2023). En "
               "materia de efectos económicos, la sentencia SL1179 de 2024 precisó que la "
               "indemnización de perjuicios por el daño causado con ocasión del traslado es una "
               "consecuencia resarcitoria única que se paga por una sola vez y que se causa respecto "
               "de quien ya obtuvo la pensión."])

    # ---------- 3.4
    h(doc, 2, "Auditoría pensional individual: caso P-03")
    h(doc, 3, "Identificación y hechos relevantes")
    rich(doc, ["Se seleccionó a la trabajadora EXT-003, asesora comercial, mujer de 56 años, "
               "vinculada desde 1999 y afiliada al Régimen de Ahorro Individual desde 1997, cuando "
               "tenía 27 años. Cuenta con 1.150 semanas cotizadas y un capital acumulado de %s. "
               "Manifiesta que un asesor comercial de la AFP la visitó en su lugar de trabajo, le "
               "indicó que «en el fondo privado se pensionaría con una mesada mucho mayor» y le hizo "
               "firmar un formulario en el que ya venía impresa la frase «me afilio libre y "
               "voluntariamente». No conserva copia del formulario, no recuerda haber recibido "
               "simulación alguna y no suscribió acta de doble asesoría." % money(CAPITAL_RAIS)])
    caja(doc, "Nota metodológica sobre los casos",
         ["Los casos P-03, AT-2026-014 y las cifras de nómina son hipotéticos y se construyeron con "
          "parámetros reales (edades, semanas, capitales y tarifas de 2026) para que el razonamiento "
          "sea replicable. En la aplicación real deben sustituirse por la historia laboral y los "
          "documentos del expediente de cada trabajador."], tipo='nota')

    h(doc, 3, "Verificación del cumplimiento de requisitos")
    tabla(doc, ["Requisito", "Estándar aplicable", "Situación de la trabajadora", "Resultado"],
          [["Edad", "57 años (mujer)", "56 años; los cumple en 2027", "Cumple próximamente"],
           ["Semanas", "1.300 semanas (Ley 100 de 1993)", "1.150 semanas", "No cumple: faltan 150 semanas"],
           ["Semanas para garantía de pensión mínima", "1.150 semanas (art. 65, Ley 100 de 1993)",
            "1.150 semanas", "Cumple"],
           ["Capital acumulado", "Suficiente para financiar una mesada superior al 110 % de la mínima "
            "para pensión anticipada", money(CAPITAL_RAIS), "No cumple"],
           ["Régimen de transición (Ley 2381 de 2024)", "750 semanas o más a abril de 2027",
            "1.150 semanas", "Cumple: conserva la Ley 100 de 1993"]],
          [3.8, 4.6, 4.2, 3.4], size=8.4,
          titulo="Tabla 15. Verificación de requisitos del caso P-03")

    h(doc, 3, "Comparación económica de los escenarios")
    rich(doc, ["La comparación se hace con un ingreso base de liquidación de %s, equivalente al "
               "promedio de los últimos diez años de cotización, y con la tasa de reemplazo del "
               "65,5 %% que corresponde a 1.300 semanas. En el escenario de prima media la mesada "
               "sería de %s. En el escenario de ahorro individual, el capital de %s financiaría una "
               "mesada aproximada de %s, inferior al salario mínimo, de modo que la trabajadora "
               "dependería del Fondo de Garantía de Pensión Mínima y percibiría %s. La pérdida "
               "patrimonial relevante (la que se reclama en el proceso) es la diferencia entre la "
               "mesada que habría recibido en prima media y el mínimo garantizado: %s mensuales, que "
               "proyectados a veinte años de disfrute representan %s." % (
                   money(IBL), money(MESADA_RPM), money(CAPITAL_RAIS),
                   money(MESADA_RAIS_SIN_GARANTIA), money(SMLMV), money(DIF), money(DIF * 12 * 20))])
    tabla(doc, ["Escenario", "Cálculo", "Mesada mensual", "Diferencia frente al RAIS"],
          [["Prima Media (Colpensiones)", "%s × 65,5 %%" % money(IBL), money(MESADA_RPM),
            money(DIF)],
           ["Ahorro Individual sin garantía", "%s ÷ (20 años × 12 meses)" % money(CAPITAL_RAIS),
            money(MESADA_RAIS_SIN_GARANTIA), "No aplica"],
           ["Ahorro Individual con garantía estatal", "Pensión mínima garantizada (1 SMMLV)",
            money(SMLMV), "$ 0"],
           [("Pérdida frente al escenario de prima media", {"b": True}),
            "Mesada de prima media menos pensión mínima garantizada",
            ("− " + money(DIF), {"b": True}), ""]],
          [4.8, 5.3, 3.0, 2.9], size=8.6, align_centro=[2, 3],
          titulo="Tabla 16. Comparación de mesadas proyectadas del caso P-03",
          fuente="Fuente: elaboración propia. La proyección del RAIS es ilustrativa y depende de la "
                 "tasa de reemplazo ofrecida por la aseguradora y de la esperanza de vida; en un "
                 "proceso real debe aportarse la simulación de la administradora.")

    h(doc, 3, "Procedencia de la demanda de ineficacia del traslado")
    rich(doc, ["La pregunta jurídica es si el traslado de 1997 produjo efectos. La respuesta se "
               "construye con los criterios de la SU-107 de 2024 aplicados al caso:"])
    tabla(doc, ["Criterio jurisprudencial", "Verificación en el caso", "Favorabilidad"],
          [["Información clara, cierta y comprensible sobre diferencias entre regímenes",
            "No obra soporte; la AFP no conserva simulación ni documento explicativo",
            "A favor de la ineficacia"],
           ["Asesoría y buen consejo personalizado", "Relato de una visita comercial con promesa de "
            "mesada superior, sin análisis del caso concreto", "A favor de la ineficacia"],
           ["Doble asesoría documentada (exigible desde 2014)", "No existe acta; el traslado es de "
            "1997, anterior a la Ley 1748 de 2014, pero el deber de información sí era exigible",
            "A favor de la ineficacia"],
           ["Valor probatorio del formulario", "Solo contiene la cláusula preimpresa de afiliación "
            "«libre y voluntaria»", "Insuficiente por sí solo"],
           ["Carga de la prueba", "La AFP deberá acreditar la información; la trabajadora aporta "
            "indicios (testimonio, ausencia de soportes)", "A favor de la ineficacia"],
           ["Proximidad a la pensión o derecho consolidado", "No se exige (SU-107 de 2024)",
            "Indiferente"],
           ["Barrera de los diez años (multivinculación)", "El traslado se produjo a los 27 años; no "
            "se configura", "No aplica"]],
          [4.6, 7.0, 4.4], size=8.3,
          titulo="Tabla 17. Aplicación de los criterios de la SU-107 de 2024 al caso P-03")

    caja(doc, "Conclusión sobre la procedencia",
         [[("La demanda de ineficacia del traslado es viable y altamente probable de éxito. ",
            {"b": True}),
           "El traslado se produjo sin evidencia documental de información suficiente, en un "
           "contexto de asesoría comercial que privilegió la captación del afiliado, y la "
           "administradora no podrá demostrar lo contrario porque no conserva soportes. La "
           "circunstancia de que la trabajadora tenga hoy 1.150 semanas y esté próxima a la edad "
           "refuerza (pero no determina) la procedencia, pues el vicio se predica del acto mismo."]],
         tipo='exito')

    h(doc, 3, "Efectos restitutorios frente al fondo de pensiones")
    for t in [
        [("Ineficacia del acto: ", {"b": True}), "la afiliación al Régimen de Ahorro Individual no "
         "produce efectos jurídicos y se tiene por no celebrada (arts. 271 y 272 de la Ley 100 de 1993)."],
        [("Restablecimiento sin solución de continuidad: ", {"b": True}), "la trabajadora queda "
         "afiliada al Régimen de Prima Media como si nunca se hubiera trasladado, de modo que las "
         "1.150 semanas se computan íntegramente para la pensión de vejez."],
        [("Traslado de la totalidad del ahorro: ", {"b": True}), "la AFP debe trasladar a Colpensiones "
         "las cotizaciones, los rendimientos financieros, los gastos de administración, las primas de "
         "los seguros previsionales y el porcentaje destinado al Fondo de Garantía de Pensión Mínima. "
         "No puede descontar comisiones ni trasladar el costo al afiliado."],
        [("Bono pensional: ", {"b": True}), "cuando exista, debe emitirse o redimirse a favor del "
         "régimen de prima media, con los intereses correspondientes."],
        [("Improcedencia de repetir contra el afiliado: ", {"b": True}), "la administradora no puede "
         "cobrar al afiliado el mayor valor de las mesadas que haya pagado, porque la ineficacia "
         "declarada no le es imputable a este."],
        [("Reparación adicional: ", {"b": True}), "si la trabajadora ya se hubiera pensionado en el "
         "RAIS, la declaratoria de ineficacia habilita la indemnización de perjuicios por la "
         "diferencia de mesadas y por el daño consolidado (CSJ, SL1179 de 2024)."],
    ]:
        bullet(doc, t)
    rich(doc, ["Una precisión temporal es indispensable: la oportunidad administrativa de traslado "
               "prevista en el artículo 76 de la Ley 2381 de 2024 venció en julio de 2026 y la "
               "ventana de la Ley 1748 de 2014 se agotó años atrás. En consecuencia, la vía judicial "
               "de ineficacia es hoy el único mecanismo idóneo para que la trabajadora regrese al "
               "Régimen de Prima Media antes de que el nuevo sistema de pilares entre en vigencia el "
               "1.º de abril de 2027."])

    h(doc, 3, "Estrategia probatoria y procesal")
    tabla(doc, ["Actuación", "Finalidad", "Fundamento"],
          [["Solicitud de copia del formulario de afiliación y del expediente", "Verificar qué "
            "documentos conserva la AFP y desde qué fecha", "Derecho de petición (art. 23 C.P.)"],
           ["Solicitud de la simulación y del acta de doble asesoría", "Acreditar la ausencia de "
            "asesoría documentada", "Ley 1748 de 2014; Circular Externa 016 de 2016"],
           ["Certificación de la historia laboral y del capital", "Cuantificar el perjuicio",
            "Art. 23, Ley 100 de 1993"],
           ["Testimonios de compañeros de trabajo", "Demostrar la dinámica de captación",
            "Art. 165 C.G.P."],
           ["Reclamación previa ante la AFP y Colpensiones", "Agotar la vía y fijar la competencia "
            "territorial", "Art. 11 C.P.T.S.S."],
           ["Demanda ordinaria laboral de ineficacia", "Obtener la declaratoria y el traslado del "
            "ahorro", "Arts. 271 y 272, Ley 100 de 1993"]],
          [5.4, 6.0, 4.6], size=8.4,
          titulo="Tabla 18. Plan probatorio del caso P-03")

    # ---------- 3.5
    h(doc, 2, "Diagnóstico del eje y recomendaciones")
    rich(doc, ["De los tres trabajadores afiliados al Régimen de Ahorro Individual, dos no cuentan "
               "con soporte de asesoría y uno presenta inconsistencia en el reporte de semanas. La "
               "empresa, como empleadora, no es responsable del deber de asesoría (que recae en la administradora), pero sí lo es de la veracidad y oportunidad de las novedades y del "
               "pago de aportes, y esos elementos son los que la jurisprudencia examina para "
               "reconstruir la historia laboral."])
    tabla(doc, ["Hallazgo", "Consecuencia jurídica", "Riesgo", "Acción"],
          [["Dos traslados sin soporte de asesoría", "Demandas de ineficacia con alta probabilidad de "
            "éxito; contingencia patrimonial para la AFP, no para la empresa", "Alto (para la AFP)",
            "Verificar la historia laboral de todos los afiliados al RAIS"],
           ["Semanas no reportadas (EXT-010)", "Afectación del derecho pensional del trabajador y "
            "posible demanda contra la empresa por falta de cotización", "Alto",
            "Corregir la planilla y reportar semanas omitidas"],
           ["Trabajadores en el umbral de la transición (EXT-002 y EXT-006)", "Cambio de régimen "
            "aplicable si no se completan las semanas al 1.º de abril de 2027", "Medio",
            "Revisar la historia laboral y programar cotizaciones"],
           ["Trabajadora EXT-003 sin decisión informada", "Pérdida de la mesada diferencial de %s "
            "mensuales" % money(DIF), "Medio",
            "Acompañar la solicitud de información y el trámite de ineficacia"]],
          [4.2, 6.2, 2.0, 3.6], size=8.3, align_centro=[2],
          titulo="Tabla 19. Hallazgos y acciones del Eje III")

    caja(doc, "Buena práctica preventiva",
         ["Toda decisión pensional de un trabajador debería quedar documentada en la hoja de vida: "
          "simulación entregada, acta de doble asesoría, formulario firmado y constancia de entrega "
          "de copia. Conservar esos documentos durante al menos veinte años no es una carga "
          "burocrática: es la única prueba posible cuando el litigio se presenta décadas después."],
         tipo='info')
    salto_pagina(doc)
