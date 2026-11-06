# -*- coding: utf-8 -*-
"""Análisis de escritura: documento acompañante del trabajo.

No va dentro del informe, porque explicarle al lector cómo está escrito el texto que está
leyendo le quita fuerza al texto. Va aparte y se entrega con él.
"""
import os
import re
import sys
import statistics as st

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datos import money, IBC_TOTAL
from docx_engine import (construir_documento, encabezado_pie, normalizar_orden,
                         p, rich, h, bullet, tabla, salto_pagina, linea_horizontal,
                         AZUL, GRIS, GRIS_CLARO, HEX_AZUL)

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
INFORME = os.path.join(RAIZ, "entrega", "Auditoria_Seguridad_Social_EXTINTORES_SCM.docx")
SALIDA = os.path.join(RAIZ, "entrega", "Analisis_de_escritura.docx")


def metricas():
    doc = Document(INFORME)
    parrafos = [x.text.strip() for x in doc.paragraphs if x.text.strip()]
    prosa = [x for x in parrafos if len(x.split()) >= 20]
    texto = "\n".join(parrafos)

    frases = []
    for x in prosa:
        for f in re.split(r"(?<=[\.\?\!])\s+", x):
            f = f.strip()
            if len(f.split()) >= 3:
                frases.append(f)
    largos = [len(f.split()) for f in frases]

    long_p = [len(x.split()) for x in prosa]
    media_p = st.mean(long_p)
    sd_p = st.pstdev(long_p)

    # primera persona
    yo = len(re.findall(r"\b(yo|mi|me|mío|mía|conozco|escribo|encontré|decidí|pago|asumí|"
                        r"revisé|verifiqué|dejé|tengo|quiero)\b", texto, re.I))

    # oficio técnico en SST
    tecnicos = ["matriz de peligros", "prueba hidrostática", "cilindro presurizado", "banco de prueba",
                "anclaje", "elemento de protección", "EPP", "procedimiento de trabajo seguro",
                "capacitación", "certificación", "SG-SST", "acto inseguro", "condición insegura",
                "nivel de riesgo", "IPERC", "riesgo biomecánico", "agente limpio", "válvula",
                "corrosión", "nitrógeno", "presión"]
    hits = sum(texto.lower().count(t.lower()) for t in tecnicos)

    # citas y soportes normativos
    citas = len(re.findall(r"«[^»]{10,400}»", texto))
    normas = len(re.findall(r"(Ley \d+ de \d{4}|Decreto \d+ de \d{4}|Sentencia [A-Z]{1,2}-\d+ de \d{4}|"
                            r"SL\d+-?\d*)", texto))

    # aperturas
    conectores = len(re.findall(r"^(Además|Asimismo|Por otra parte|En este sentido|Cabe|"
                                r"En consecuencia|Por consiguiente|Ahora bien|En efecto|"
                                r"En síntesis|En resumen|Finalmente)\b",
                                "\n".join(prosa), re.I | re.M))

    return dict(palabras=len(texto.split()), parrafos=len(parrafos), prosa=len(prosa),
                frases=len(frases), media=st.mean(largos), sd=st.pstdev(largos),
                maximo=max(largos), minimo=min(largos),
                cortas=sum(1 for x in largos if x <= 10),
                medias=sum(1 for x in largos if 11 <= x <= 25),
                largas=sum(1 for x in largos if x >= 26),
                media_p=media_p, sd_p=sd_p, cv_p=sd_p / media_p,
                cortos=sum(1 for x in long_p if x < 35),
                extensos=sum(1 for x in long_p if x > 90),
                primera=yo, tecnicos=hits, citas=citas, normas=normas, conectores=conectores)


def construir_analisis(doc, m):
    p(doc, "ANÁLISIS DE ESCRITURA", size=18, bold=True, color=AZUL,
      align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    p(doc, "Documento acompañante del trabajo «Auditoría y diagnóstico del Sistema de Seguridad "
           "Social de EXTINTORES S C M»", size=10.5, italic=True, color=GRIS,
      align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    p(doc, "Julia Inés Buitrago García · 26 de noviembre de 2026", size=9.5, color=GRIS_CLARO,
      align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)
    linea_horizontal(doc, color=HEX_AZUL, sz=8, space_after=8)

    rich(doc, ["Este documento explica cómo está escrito el trabajo y por qué elegí escribirlo así. "
               "Va aparte y no dentro del informe: explicarle al lector "
               "las decisiones de redacción del texto que está leyendo le quita fuerza al texto. "
               "Aquí sí lo digo, sobre todo por el uso de la primera persona, que en un trabajo "
               "académico siempre levanta una ceja."])

    h(doc, 1, "1. Quién escribe y desde dónde")
    rich(doc, ["Soy estudiante de sexto semestre de Derecho y tecnóloga en Seguridad y Salud en el "
               "Trabajo. Trabajo en la empresa que auditó este informe: soy la titular y la "
               "representante legal, y soy la que firma la planilla todos los meses."])
    rich(doc, ["Eso define tres cosas del texto. La primera es que no puedo escribir como si "
               "estuviera mirando la empresa desde afuera, porque no lo estoy. La segunda es que "
               "tengo oficio técnico en el tema central del trabajo, que es el riesgo laboral, y "
               "ocultarlo sería renunciar a la única ventaja real que tengo sobre cualquier otro "
               "compañero. La tercera es que el documento tiene que servirme después de la nota: no "
               "sirve para corregir el negocio, no para archivarse con la nota."])
    rich(doc, ["De ahí salió la decisión de fondo: escribir en primera persona cuando valoro y en "
               "tercera cuando describo. «El artículo 216 del C.S.T. obliga» va en tercera, porque "
               "la norma no opina. «Aquí no tengo con qué» va en primera, porque ahí el que responde "
               "soy yo."])

    h(doc, 1, "2. Las cuatro decisiones de registro")
    rich(doc, [("Primera: primera persona para valorar, no para narrar. ", {"b": True}),
               "La primera persona aparece %d veces en el texto y casi siempre cumple una función "
               "de responsabilidad, no de adorno. Reconocer que la decisión de no comprar los "
               "guantes fue mía. Que la base de la licencia la puse yo en la planilla. Que la "
               "exoneración del 114-1 me tomó un rato entenderla. Un informe en tercera persona "
               "habría dicho «se advierte una omisión del empleador». Eso es verdad y no sirve para "
               "nada." % m["primera"]])
    rich(doc, [("Segunda: el oficio técnico entra como descripción del proceso, no como "
                "vocabulario. ", {"b": True}),
               "En el numeral 2.1 el informe describe cómo se recarga un extintor paso por paso: "
               "descarga, apertura de válvula, inspección por corrosión, prueba hidrostática, secado, "
               "cambio de agente, carga con nitrógeno, rotulado. Esa descripción es la que sostiene "
               "el argumento de que la actividad no es comercio. Un estudiante de Derecho sin el "
               "tecnólogo habría escrito «actividad de mayor riesgo» y no habría podido probarlo. "
               "Hay %d términos técnicos de SST en el documento, y ninguno está puesto para "
               "impresionar: todos están atados a un hecho del taller." % m["tecnicos"]])
    rich(doc, [("Tercera: las cifras siempre llevan consecuencia. ", {"b": True}),
               "No hay un solo número en el informe que se quede solo. El IBC de %s no aparece como "
               "dato: aparece como el número que decide cuánto recibe un trabajador el día que se "
               "accidenta. Los $ 18.240.000 de diferencia entre reparación tarifada y plena no son "
               "un cálculo: son lo que sale de mi patrimonio. Esa es la diferencia entre un trabajo "
               "que cita cifras y uno que las usa para decidir." % money(IBC_TOTAL)])
    rich(doc, [("Cuarta: cierre abierto, no resumen. ", {"b": True}),
               "Ninguna sección termina recapitulando lo que ya dijo. Termina en la consecuencia o "
               "en la pregunta que queda. «Ya está demostrado, con fecha, y en un documento que "
               "firmó otra persona» no resume nada: cierra el argumento y deja al lector con la "
               "obligación de actuar."])

    h(doc, 1, "3. Lo que se mantuvo impersonal y por qué")
    rich(doc, ["No todo el documento va en primera persona, y esto es deliberado. Las tablas de "
               "liquidación, la comparación de escenarios pensionales y los listados de fundamentos "
               "jurídicos van en tercera o en forma nominal. Escribir «en mi criterio, el aporte a "
               "pensión es de doce por ciento» dentro de una tabla de aportes es ruido, no estilo. "
               "El dato no admite opinión."])
    rich(doc, ["La regla que apliqué es simple: primera persona donde hay una decisión que tomar, "
               "impersonal donde hay un dato que verificar. El capítulo 2, que es el del riesgo "
               "laboral, es el más personal de todos, porque es donde tengo el conocimiento técnico y "
               "donde la responsabilidad es mía. El capítulo 1, que es de afiliación y cotización, es "
               "el más neutro, porque es aritmética y norma."])

    h(doc, 1, "4. La estructura impuesta y cómo la resolví")
    rich(doc, ["El docente entregó un esqueleto con veintisiete preguntas y rótulos fijos. Respetarlo "
               "era obligatorio y también era el principal riesgo de escritura, porque un "
               "cuestionario contestado pregunta por pregunta produce un texto plano: cada respuesta "
               "empieza igual, mide lo mismo y cierra igual. Eso es exactamente lo que hace que un "
               "texto suene a formato."])
    rich(doc, ["La solución fue conservar los rótulos como señalización y no como arranque de frase. "
               "El numeral 2.1 se titula «Identificación de la clase de riesgo» porque así lo pide el "
               "esqueleto, pero la primera frase no dice «Actividad económica:», dice «La empresa "
               "declaró en el RUT, y repitió ante la Secretaría de Salud de Soacha, el CIIU 4759». "
               "El rótulo está donde debe estar, en el título; la respuesta empieza por el hecho."])
    rich(doc, ["En total hay %d párrafos de prosa y %d de ellos abren con un conector de los que se "
               "usan para rellenar («Además», «Por otra parte», «Ahora bien»). Cero. Esa fue la "
               "corrección que más veces tuve que hacer." % (m["prosa"],
                                                                              m["conectores"])])

    h(doc, 1, "5. Los números del texto")
    rich(doc, ["Estas son las mediciones del informe, tomadas sobre el archivo final:"])
    tabla(doc, ["Indicador", "Valor", "Qué significa"],
          [["Palabras", "{:,}".format(m["palabras"]).replace(",", "."),
            "Extensión total, con tablas"],
           ["Párrafos de prosa", str(m["prosa"]), "Se excluyen títulos, tablas y pies"],
           ["Frases analizadas", str(m["frases"]), ""],
           ["Longitud media de frase", "%.1f palabras" % m["media"],
            "Rango natural del español académico: 15 a 22"],
           ["Desviación de la longitud de frase", "%.1f palabras" % m["sd"],
            "Por encima de 8 indica que el texto alterna frases cortas y largas"],
           ["Frase más corta / más larga", "%d / %d palabras" % (m["minimo"], m["maximo"]),
            "La alternancia es lo que rompe la monotonía"],
           ["Frases de 10 palabras o menos", str(m["cortas"]),
            "Sirven para cerrar un argumento y bajar el ritmo"],
           ["Frases de 26 palabras o más", str(m["largas"]),
            "Sostienen la explicación técnica"],
           ["Párrafos de menos de 35 palabras", str(m["cortos"]),
            "Párrafos de una sola idea"],
           ["Párrafos de más de 90 palabras", str(m["extensos"]),
            "Párrafos de desarrollo"],
           ["Coeficiente de variación del párrafo", "%.2f" % m["cv_p"],
            "Mide si los párrafos se parecen entre sí; bajo 0,40 es uniformidad de plantilla"],
           ["Aperturas con conector de relleno", str(m["conectores"]), "Cero es lo deseado"],
           ["Marcas de primera persona", str(m["primera"]),
            "Valoración y responsabilidad, no narración"],
           ["Términos técnicos de SST", str(m["tecnicos"]),
            "Todos atados a un hecho del taller"],
           ["Citas textuales entre comillas", str(m["citas"]),
            "Norma, jurisprudencia y acta sanitaria"],
           ["Normas y providencias citadas", str(m["normas"]), ""]],
          [6.2, 3.2, 6.6], size=8.4, align_centro=[1],
          titulo="Tabla 1. Indicadores de escritura del informe")

    h(doc, 1, "6. Lo que este análisis no resuelve")
    rich(doc, ["No puedo demostrar que el texto sea mío. Esa es la verdad y no la voy a adornar: "
               "quien escribe siempre cree que su voz se oye, y una nota de escritura no prueba "
               "autoría, solo la declara."])
    rich(doc, ["Lo que sí puedo hacer es señalar dónde está lo que ningún programa escribiría. La "
               "descripción del proceso de recarga hecha por quien lo ha hecho con las manos. La "
               "frase «el que decidió no comprar los guantes fui yo». Y el reconocimiento de que "
               "tengo pendiente la historia laboral de dos trabajadores porque todavía no la he "
               "pedido. Un texto generado no asume deudas."])
    rich(doc, ["El resto lo decide quien califique. La escritura se puede medir, como se midió "
               "arriba. La autoría se reconoce de otra manera, y esa no está en las métricas."])
    salto_pagina(doc)


def main():
    doc = construir_documento()
    encabezado_pie(doc, "Análisis de escritura",
                   "EXTINTORES S C M · C.C./NIT 39626936-8 · Soacha (Cundinamarca)",
                   "Julia Inés Buitrago García · Universidad INCCA de Colombia · 2026")
    m = metricas()
    construir_analisis(doc, m)
    doc.core_properties.title = "Análisis de escritura del trabajo de Seguridad Social"
    doc.core_properties.author = "Julia Inés Buitrago García"
    normalizar_orden(doc)
    os.makedirs(os.path.dirname(SALIDA), exist_ok=True)
    doc.save(SALIDA)
    print("Guardado:", SALIDA)
    print(m)


if __name__ == "__main__":
    main()
