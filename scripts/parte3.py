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
    rich(doc, ["Elegir régimen es ejercer el derecho a la seguridad social (artículo 48 de la "
               "Constitución), pero la elección solo vale si el afiliado estaba informado. De esa "
               "premisa sale todo el régimen de responsabilidad de las administradoras y, con él, el "
               "remedio de la ineficacia del traslado que se estudia en el numeral 3.4."])

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
          titulo="Tabla 15. Comparación de los dos regímenes del SGP")

    h(doc, 3, "La reforma pensional de la Ley 2381 de 2024 y su estado actual")
    rich(doc, ["La Ley 2381 de 2024 reorganiza el sistema en ",
               ("cuatro pilares", {"b": True}),
               " (solidario, semicontributivo, contributivo y ahorro voluntario) y manda que las "
               "cotizaciones hasta 2,3 salarios mínimos vayan al componente que administra "
               "Colpensiones, con el excedente al ahorro individual. La Corte Constitucional la "
               "suspendió primero y luego, en la ",
               ("Sentencia C-264 de 2026", {"b": True}),
               ", declaró exequible la mayor parte: devolvió varias disposiciones a la Cámara de "
               "Representantes por vicios de trámite y fijó la entrada en vigencia para el ",
               ("1.º de abril de 2027", {"b": True}),
               ". Hasta entonces se aplica la Ley 100 de 1993, tal cual."])
    rich(doc, ["Conviene subrayarlo porque es la fuente de casi todos los errores que se ven en "
               "informes pensionales recientes: hoy no rige la reforma. Un documento que en 2026 "
               "liquide una pensión con las reglas de la Ley 2381 está aplicando una ley que todavía "
               "no produce efectos."])
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
          titulo="Tabla 16. Impacto de la reforma pensional en la organización")

    caja(doc, "Sobre la vigencia",
         ["Toda afirmación sobre pensiones en Colombia tiene que decir a qué fecha se refiere. Hasta "
          "el 31 de marzo de 2027 rige la Ley 100 de 1993; desde el 1.º de abril de 2027, y siempre "
          "que se complete la subsanación que ordenó la Corte, rige la Ley 2381 de 2024. Esta "
          "auditoría se practicó en el periodo de transición, así que combina los dos marcos y va "
          "señalando cuál aplica en cada caso. Ahí se juega la utilidad del informe: un documento "
          "que no fija la vigencia envejece en tres meses."],
         tipo='alerta')

    # ---------- 3.2
    h(doc, 2, "Mapa pensional de la organización")
    rich(doc, ["Los datos salen del certificado de afiliación y de la planilla integrada de "
               "liquidación de aportes. El mapa muestra una concentración en el Régimen de Prima "
               "Media (%d de %d trabajadores, "
               "equivalente al %.0f %%) y una presencia minoritaria en el Régimen de Ahorro "
               "Individual (%d trabajadores, %.0f %%), lo que se explica por la antigüedad de la "
               "planta y por el perfil salarial de la empresa." % (PM, PM + RAIS, 100 * PM / (PM + RAIS),
                                                                   RAIS, 100 * RAIS / (PM + RAIS))])
    imagen(doc, "assets/f3_mapa_pensional.png", 14.6,
           "Figura 5. Mapa pensional de EXTINTORES SCM. Fuente: elaboración propia con base en los "
           "certificados de afiliación y la planilla PILA del periodo auditado.")

    tabla(doc, ["Cód.", "Trabajador", "Edad *", "Régimen", "Semanas *",
                "Situación frente a la reforma", "Alerta"],
          [["WT-01", "Julia Inés Buitrago García", "Por verificar", "Prima Media", "Por verificar",
            "Por verificar en la historia laboral", "Media"],
           ["WT-02", "Luis Antonio Aguilar Rodríguez", "56", "Ahorro Individual", "1.150",
            "Régimen de transición (más de 750 semanas) · caso auditado P-03", "Alta"],
           ["WT-03", "Felipe Martínez", "Por verificar", "Prima Media", "Por verificar",
            "Por verificar en la historia laboral", "Media"]],
          [1.4, 4.4, 1.2, 2.6, 1.5, 4.4, 1.3], size=8.2, header_size=7.6,
          align_centro=[0, 2, 4, 6],
          titulo="Tabla 17. Distribución de los trabajadores por régimen y situación pensional",
          fuente="Fuente: elaboración propia con base en los certificados de afiliación y los datos "
                 "aportados por la titular. (*) La edad y las semanas de WT-02 corresponden al "
                 "expediente del caso P-03 y están sujetas a confirmación con la historia laboral; "
                 "para WT-01 y WT-03 deben extraerse de la historia laboral de cada afiliado con "
                 "corte a 30 de septiembre de 2026.")

    rich(doc, ["Con tres personas no se puede hablar de promedio. Cada trabajador es una unidad de "
               "análisis y un error en una historia laboral afecta a un tercio de la planta. De ahí "
               "salen tres cosas."])
    rich(doc, ["La primera es que solo WT-02 está en el régimen de transición, con más de 750 "
               "semanas, así que sus condiciones pensionales dependen de la Ley 100 y no del sistema "
               "de pilares. La segunda es que la titular y WT-03 necesitan verificación individual: "
               "alrededor de los umbrales, una semana no reportada cambia el régimen aplicable y con "
               "él la mesada. La tercera es que la concentración en prima media (dos de tres) hace "
               "que cualquier traslado al RAIS tenga un peso desproporcionado, porque el empleador "
               "responde por la veracidad de la información que sirvió de base a la decisión."])

    # ---------- 3.3
    h(doc, 2, "El deber de información y la doble asesoría")
    rich(doc, ["Un traslado de régimen se firma en una oficina en veinte minutos y se paga durante "
               "treinta años. El afiliado advierte el efecto décadas después, cuando ya no puede "
               "deshacerlo, y por eso la jurisprudencia construyó un estándar de información que fue "
               "subiendo de exigencia en tres etapas:"])
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
               " unificó todo esto y hoy es el punto de partida de cualquier litigio pensional. Fijó "
               "cuatro reglas. El deber de información existe desde la creación de las "
               "administradoras y su exigencia creció con el tiempo. El consentimiento del "
               "formulario no basta: se necesita consentimiento informado. La carga de la prueba se "
               "reparte, porque el afiliado aporta indicios del déficit y la AFP debe demostrar que "
               "informó y asesoró. Y la ineficacia no exige estar cerca de pensionarse ni tener un "
               "derecho consolidado, porque el vicio está en el acto mismo de traslado."])
    cita(doc, "«El simple consentimiento vertido en el formulario de afiliación es insuficiente, por "
              "la necesidad de un consentimiento informado […] la violación del deber de información "
              "se predica frente a la validez del acto jurídico de traslado, considerado en sí "
              "mismo.»",
         "Corte Constitucional, Sentencia SU-107 de 2024 (9 de abril de 2024)")

    rich(doc, ["La Sala de Casación Laboral había llegado antes al mismo punto: «es a la AFP a quien "
               "le corresponde acreditar el cumplimiento del deber de información al momento del "
               "traslado de régimen pensional, pues exigir al afiliado una prueba de este alcance es "
               "un despropósito» (CSJ, SL3179-2023). Ahí está el detalle que decide los casos: quien "
               "tiene los papeles es la administradora."])
    rich(doc, ["Sobre los efectos económicos, la sentencia SL1179 de 2024 precisó que la "
               "indemnización de perjuicios por el traslado es una consecuencia resarcitoria única, "
               "se paga por una sola vez y se causa respecto de quien ya obtuvo la pensión. No es "
               "una suma acumulable ni se reconoce dos veces por el mismo vicio."])

    # ---------- 3.4
    h(doc, 2, "Auditoría pensional individual: caso P-03")
    h(doc, 3, "Identificación y hechos relevantes")
    rich(doc, ["Se seleccionó al trabajador WT-02, Luis Antonio Aguilar Rodríguez (C.C. 1073694256), "
               "técnico de recarga, hombre de 56 años, vinculado desde 1999 y afiliado al Régimen de "
               "Ahorro Individual desde 1997, cuando tenía 27 años. Cuenta con 1.150 semanas "
               "cotizadas y un capital acumulado de %s. "
               "Manifiesta que un asesor comercial de la AFP la visitó en su lugar de trabajo, le "
               "indicó que «en el fondo privado se pensionaría con una mesada mucho mayor» y le hizo "
               "firmar un formulario en el que ya venía impresa la frase «me afilio libre y "
               "voluntariamente». No conserva copia del formulario, no recuerda haber recibido "
               "simulación alguna y no suscribió acta de doble asesoría." % money(CAPITAL_RAIS)])
    caja(doc, "Nota metodológica sobre los casos",
         ["Los casos P-03 y AT-2026-014 se atribuyen a trabajadores reales de la planta, pero los "
          "hechos, edades, semanas y capitales son parámetros de trabajo construidos con las tarifas "
          "de 2026 para que el razonamiento sea replicable. Antes de la presentación definitiva deben "
          "sustituirse por la historia laboral y los documentos del expediente de cada trabajador; el "
          "documento marca con asterisco toda cifra que aún no está respaldada por un soporte."],
         tipo='nota')

    h(doc, 3, "Verificación del cumplimiento de requisitos")
    tabla(doc, ["Requisito", "Estándar aplicable", "Situación del trabajador", "Resultado"],
          [["Edad", "57 años (mujer)", "56 años; los cumple en 2027", "Cumple próximamente"],
           ["Semanas", "1.300 semanas (Ley 100 de 1993)", "1.150 semanas", "No cumple: faltan 150 semanas"],
           ["Semanas para garantía de pensión mínima", "1.150 semanas (art. 65, Ley 100 de 1993)",
            "1.150 semanas", "Cumple"],
           ["Capital acumulado", "Suficiente para financiar una mesada superior al 110 % de la mínima "
            "para pensión anticipada", money(CAPITAL_RAIS), "No cumple"],
           ["Régimen de transición (Ley 2381 de 2024)", "750 semanas o más a abril de 2027",
            "1.150 semanas", "Cumple: conserva la Ley 100 de 1993"]],
          [3.8, 4.6, 4.2, 3.4], size=8.4,
          titulo="Tabla 18. Verificación de requisitos del caso P-03")

    h(doc, 3, "Comparación económica de los escenarios")
    rich(doc, ["La comparación usa un ingreso base de liquidación de %s, que es el promedio de los "
               "últimos diez años de cotización, y la tasa de reemplazo del 65,5 %% de 1.300 semanas. "
               "En prima media la mesada sería de %s. En ahorro individual, el capital de %s "
               "financiaría unos %s mensuales, por debajo del salario mínimo, así que el trabajador "
               "terminaría en el Fondo de Garantía de Pensión Mínima con %s. Ahí está la pérdida: %s "
               "mensuales, que a veinte años de disfrute son %s." % (
                   money(IBL), money(MESADA_RPM), money(CAPITAL_RAIS),
                   money(MESADA_RAIS_SIN_GARANTIA), money(SMLMV), money(DIF), money(DIF * 12 * 20))])
    rich(doc, ["Lo que se reclama es esa diferencia, ni el capital ni la mesada completa. Aquí se "
               "equivocan muchas demandas: el daño no consiste en haber cotizado al fondo, consiste "
               "en haber quedado con una mesada menor por un traslado que nunca se explicó."])
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
          titulo="Tabla 19. Comparación de mesadas proyectadas del caso P-03",
          fuente="Fuente: elaboración propia. La proyección del RAIS es ilustrativa y depende de la "
                 "tasa de reemplazo ofrecida por la aseguradora y de la esperanza de vida; en un "
                 "proceso real debe aportarse la simulación de la administradora.")

    h(doc, 3, "Procedencia de la demanda de ineficacia del traslado")
    rich(doc, ["La pregunta, en concreto, es una sola: ¿el traslado de 1997 produjo efectos? Se "
               "responde con los criterios de la SU-107 de 2024 aplicados al expediente."])
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
           ["Carga de la prueba", "La AFP deberá acreditar la información; el trabajador aporta "
            "indicios (testimonio, ausencia de soportes)", "A favor de la ineficacia"],
           ["Proximidad a la pensión o derecho consolidado", "No se exige (SU-107 de 2024)",
            "Indiferente"],
           ["Barrera de los diez años (multivinculación)", "El traslado se produjo a los 27 años; no "
            "se configura", "No aplica"]],
          [4.6, 7.0, 4.4], size=8.3,
          titulo="Tabla 20. Aplicación de los criterios de la SU-107 de 2024 al caso P-03")

    caja(doc, "Conclusión sobre la procedencia",
         [[("La demanda de ineficacia del traslado es viable y altamente probable de éxito. ",
            {"b": True}),
           "El traslado se produjo sin evidencia documental de información suficiente, en un "
           "contexto de asesoría comercial que privilegió la captación del afiliado, y la "
           "administradora no podrá demostrar lo contrario porque no conserva soportes. La "
           "circunstancia de que el trabajador tenga hoy 1.150 semanas y esté próximo a la edad "
           "refuerza (pero no determina) la procedencia, pues el vicio se predica del acto mismo."]],
         tipo='exito')

    h(doc, 3, "Efectos restitutorios frente al fondo de pensiones")
    for t in [
        [("Ineficacia del acto: ", {"b": True}), "la afiliación al Régimen de Ahorro Individual no "
         "produce efectos jurídicos y se tiene por no celebrada (arts. 271 y 272 de la Ley 100 de 1993)."],
        [("Restablecimiento sin solución de continuidad: ", {"b": True}), "el trabajador queda "
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
        [("Reparación adicional: ", {"b": True}), "si el trabajador ya se hubiera pensionado en el "
         "RAIS, la declaratoria de ineficacia habilita la indemnización de perjuicios por la "
         "diferencia de mesadas y por el daño consolidado (CSJ, SL1179 de 2024)."],
    ]:
        bullet(doc, t)
    rich(doc, ["Hay una precisión temporal que cambia la estrategia. La oportunidad administrativa "
               "de traslado del artículo 76 de la Ley 2381 de 2024 venció en julio de 2026, y la "
               "ventana de la Ley 1748 de 2014 se cerró hace años. Queda la vía judicial. Hoy la "
               "demanda de ineficacia es el único camino para que el trabajador vuelva a prima media "
               "antes del 1.º de abril de 2027, y después de esa fecha el cálculo se complica porque "
               "el régimen aplicable cambia de base."])

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
          titulo="Tabla 21. Plan probatorio del caso P-03")

    # ---------- 3.5
    h(doc, 2, "Diagnóstico del eje y recomendaciones")
    rich(doc, ["De los tres trabajadores, uno está en ahorro individual sin soporte de asesoría y los "
               "otros dos tienen la historia laboral por verificar. Conviene separar responsabilidades, "
               "porque es fácil confundirlas: el deber de asesoría recae en la administradora, no en "
               "la empresa. Lo que sí es de la empresa es la veracidad y la oportunidad de las "
               "novedades y el pago de aportes. Y son justamente esos dos elementos los que la "
               "jurisprudencia examina cuando reconstruye una historia laboral."])
    tabla(doc, ["Hallazgo", "Consecuencia jurídica", "Riesgo", "Acción"],
          [["WT-02 trasladado al RAIS sin soporte de asesoría", "Demanda de ineficacia con alta "
            "probabilidad de éxito; la contingencia patrimonial es de la AFP, no de la empresa",
            "Alto (para la AFP)", "Acompañar la solicitud de información y el trámite de ineficacia"],
           ["Pérdida de mesada diferencial de WT-02 (%s mensuales)" % money(DIF),
            "Daño patrimonial que se reclama en el proceso de ineficacia", "Medio",
            "Cuantificar con la simulación de la administradora y la historia laboral certificada"],
           ["Historia laboral de WT-01 y WT-03 por verificar", "Riesgo de novedades mal reportadas y "
            "de semanas no acreditadas al 1.º de abril de 2027", "Medio",
            "Solicitar la historia laboral con corte a la fecha y cruzar con las planillas"],
           ["Trabajadores cerca del umbral de transición (750 semanas)", "Pérdida del régimen de "
            "transición por semanas no acreditadas", "Medio",
            "Revisar semana por semana y programar las cotizaciones faltantes"]],
          [4.2, 6.2, 2.0, 3.6], size=8.3, align_centro=[2],
          titulo="Tabla 22. Hallazgos y acciones del Eje III")

    caja(doc, "Una práctica que cuesta nada",
         ["Toda decisión pensional de un trabajador debería quedar en la hoja de vida: simulación "
          "entregada, acta de doble asesoría, formulario firmado y constancia de entrega de copia. "
          "Guardar eso veinte años suena exagerado, hasta que el litigio llega dos décadas después y "
          "la única prueba que queda es la que alguien conservó."],
         tipo='info')
    salto_pagina(doc)
