#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
validar_docx.py — validación de un .docx generado con python-docx.

Comprueba:
  1. Orden de los hijos de <w:tcPr>, <w:tblPr>, <w:rPr>, <w:pPr> según el esquema OOXML
     (Word es tolerante al leer, pero estricto al abrir archivos con orden inválido).
  2. Elementos obligatorios de un .docx (content types, relaciones, documento principal).
  3. Artefactos de plantilla en el texto (placeholders de Python sin sustituir).
  4. Tablas sin encabezado o sin fuente declarada, y figuras sin pie.

Uso:  python3 scripts/validar_docx.py entrega/archivo.docx
Salida: lista de problemas; exit 0 si no hay ninguno, 1 si hay hallazgos.
"""
import sys
import os
import zipfile
import re
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}

# Orden del esquema (ECMA-376) para los contenedores de propiedades que se suelen desordenar.
ORDEN = {
    "tcPr": ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap",
             "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark"],
    "tblPr": ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
              "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders",
              "shd", "tblLayout", "tblCellMar", "tblLook", "tblCaption", "tblDescription"],
    "trPr": ["cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter", "cantSplit",
             "trHeight", "tblHeader", "tblCellSpacing", "jc", "hidden"],
    "pPr": ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl",
            "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens",
            "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN",
            "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing",
            "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment",
            "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"],
    "rPr": ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike",
            "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish",
            "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight",
            "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang",
            "eastAsianLayout", "specVanish", "oMath"],
}

ARTEFACTOS = [r"\{'", r"%s\b", r"%d\b", r"\bNone\b", r"money\(", r"\[nombre", r"\bEXT-0\d",
              r"\{\"b\""]


def validar(ruta):
    problemas = []
    if not os.path.exists(ruta):
        return ["No existe el archivo: %s" % ruta]

    with zipfile.ZipFile(ruta) as z:
        nombres = set(z.namelist())
        for requerido in ["[Content_Types].xml", "_rels/.rels", "word/document.xml",
                          "word/_rels/document.xml.rels"]:
            if requerido not in nombres:
                problemas.append("Falta la parte obligatoria: %s" % requerido)
        xml = z.read("word/document.xml")

    raiz = etree.fromstring(xml)

    # 1) orden de propiedades
    for cont in ORDEN:
        orden = ORDEN[cont]
        for el in raiz.iter("{%s}%s" % (W, cont)):
            hijos = [etree.QName(h).localname for h in el]
            idx = [orden.index(h) for h in hijos if h in orden]
            if idx != sorted(idx):
                problemas.append("Orden inválido en <%s>: %s" % (cont, hijos))
            desconocidos = [h for h in hijos if h not in orden]
            if desconocidos:
                problemas.append("<%s> con hijo no previsto: %s" % (cont, desconocidos))

    # 2) artefactos de plantilla en el texto
    textos = [t.text for t in raiz.iter("{%s}t" % W) if t.text]
    completo = "\n".join(textos)
    for pat in ARTEFACTOS:
        for m in re.finditer(pat, completo):
            ini = max(0, m.start() - 40)
            problemas.append("Artefacto %r en: …%s…" % (pat, completo[ini:m.end() + 40]))

    # 3) tablas: encabezado con texto y número de columnas coherente
    for i, tbl in enumerate(raiz.iter("{%s}tbl" % W), 1):
        filas = list(tbl.iter("{%s}tr" % W))
        if not filas:
            problemas.append("Tabla %d sin filas" % i)
            continue
        celdas = list(filas[0].iter("{%s}tc" % W))
        vacias = [c for c in celdas
                  if not "".join(t.text or "" for t in c.iter("{%s}t" % W)).strip()]
        if celdas and len(vacias) == len(celdas):
            problemas.append("Tabla %d sin encabezado (primera fila vacía)" % i)
        anchos = set()
        for f in filas:
            anchos.add(len(list(f.iter("{%s}tc" % W))))
        if len(anchos) > 1:
            problemas.append("Tabla %d con filas de distinto número de celdas: %s" % (i, sorted(anchos)))

    # 4) imágenes con pie
    imagenes = len(list(raiz.iter("{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}inline")))
    pies = len(re.findall(r"^Figura \d+\.", completo, flags=re.M))
    if imagenes and pies < imagenes:
        problemas.append("Hay %d imágenes y solo %d pies de figura" % (imagenes, pies))

    return problemas


def main():
    if len(sys.argv) < 2:
        print("Uso: python3 scripts/validar_docx.py <archivo.docx>")
        return 2
    ruta = sys.argv[1]
    problemas = validar(ruta)
    print("=" * 70)
    print("VALIDACIÓN OOXML · %s" % os.path.basename(ruta))
    print("=" * 70)
    if not problemas:
        print("PROBLEMAS: 0 — el archivo es válido y está limpio de artefactos.")
        return 0
    for p in problemas:
        print(" ·", p)
    print("\nPROBLEMAS: %d" % len(problemas))
    return 1


if __name__ == "__main__":
    sys.exit(main())
