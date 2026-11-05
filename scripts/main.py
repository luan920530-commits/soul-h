# -*- coding: utf-8 -*-
"""Generador del informe: PROYECTO INTEGRADOR · EXTINTORES SCM."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docx_engine import construir_documento, encabezado_pie, normalizar_orden
import parte1, parte2, parte3, parte4, parte5

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SALIDA = os.path.join(RAIZ, "entrega",
                     "Auditoria_Seguridad_Social_EXTINTORES_SCM.docx")


def main():
    os.chdir(RAIZ)
    doc = construir_documento()
    encabezado_pie(doc,
                   "Proyecto integrador · Auditoría y diagnóstico de seguridad social",
                   "EXTINTORES SCM · NIT 39626936-8",
                   "Julia Inés Buitrago García · Universidad INCCA de Colombia · 2026")

    parte1.portada(doc)
    parte1.contenido(doc)
    parte1.resumen_ejecutivo(doc)
    parte1.metodologia(doc)
    parte1.introduccion(doc)
    parte1.eje_i(doc)
    parte2.eje_ii(doc)
    parte3.eje_iii(doc)
    parte4.eje_iv(doc)
    parte5.capitulo5(doc)
    parte5.anexos(doc)

    # Propiedades del documento
    cp = doc.core_properties
    cp.title = "Auditoría y diagnóstico de seguridad social organizacional: EXTINTORES SCM"
    cp.subject = "Derecho de la Seguridad Social · Proyecto integrador"
    cp.author = "Julia Inés Buitrago García"
    cp.keywords = ("seguridad social; riesgos laborales; pensiones; IBC; culpa patronal; "
                   "ineficacia del traslado; competencia judicial")
    cp.comments = "Plantilla técnico-jurídica para el desarrollo del proyecto integrador (40 %)."

    normalizar_orden(doc)
    os.makedirs(os.path.dirname(SALIDA), exist_ok=True)
    doc.save(SALIDA)
    print("Guardado:", SALIDA)
    print("Tamaño: %.2f MB" % (os.path.getsize(SALIDA) / 1e6))


if __name__ == "__main__":
    main()
