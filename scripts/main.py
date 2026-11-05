# -*- coding: utf-8 -*-
"""Generador del trabajo final, ajustado al esqueleto entregado por el docente."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docx_engine import construir_documento, encabezado_pie, normalizar_orden
import esqueleto

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SALIDA = os.path.join(RAIZ, "entrega",
                      "Auditoria_Seguridad_Social_EXTINTORES_SCM.docx")


def main():
    os.chdir(RAIZ)
    doc = construir_documento()
    encabezado_pie(doc,
                   "Auditoría y diagnóstico del Sistema de Seguridad Social",
                   "EXTINTORES S C M · C.C./NIT 39626936-8 · Soacha (Cundinamarca)",
                   "Julia Inés Buitrago García · Universidad INCCA de Colombia · 2026")

    esqueleto.construir(doc)

    cp = doc.core_properties
    cp.title = "Auditoría y diagnóstico del Sistema de Seguridad Social de EXTINTORES S C M"
    cp.subject = "Derecho de la Seguridad Social · Proyecto integrador"
    cp.author = "Julia Inés Buitrago García"
    cp.keywords = ("seguridad social; riesgos laborales; pensiones; IBC; culpa patronal; "
                   "SG-SST; acta sanitaria 028-050164-26; ineficacia del traslado; "
                   "competencia judicial")
    cp.comments = "Trabajo final elaborado sobre la estructura del esqueleto del curso."

    normalizar_orden(doc)
    os.makedirs(os.path.dirname(SALIDA), exist_ok=True)
    doc.save(SALIDA)
    print("Guardado:", SALIDA)
    print("Tamaño: %.2f MB" % (os.path.getsize(SALIDA) / 1e6))


if __name__ == "__main__":
    main()
