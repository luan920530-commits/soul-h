# -*- coding: utf-8 -*-
"""Genera los dos escritos complementarios como archivos independientes.

El informe (sección 4) reproduce íntegro el derecho de petición ante la ARL. Los otros dos
—petición ante la AFP y solicitud de agotamiento de la vía interna ante el empleador— van en
documentos separados, listos para radicar, para no alargar el trabajo.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docx_engine import construir_documento, normalizar_orden
from docx.shared import Pt
from parte4 import (encabezado_escrito, rotulo, parrafo_doc, numerado_doc, firma, bullet,
                    AUTOR, CC, NIT, DIR, TEL, MAIL, CIUDAD_FECHA, EPS, ARL, AFP_RPM, AFP_RAIS)
from datos import money, EMPRESA

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SALIDA = os.path.join(RAIZ, "entrega")


def _nuevo(titulo):
    doc = construir_documento()
    doc.core_properties.title = titulo
    doc.core_properties.author = AUTOR
    return doc


def _guardar(doc, nombre):
    normalizar_orden(doc)
    os.makedirs(SALIDA, exist_ok=True)
    ruta = os.path.join(SALIDA, nombre)
    doc.save(ruta)
    print("Guardado:", ruta)


# ---------------------------------------------------------------- AFP
def peticion_afp():
    doc = _nuevo("Derecho de petición ante la AFP")
    encabezado_escrito(doc, "%s  ·  OFICINA DE SERVICIO AL CLIENTE" % AFP_RAIS.upper(), CIUDAD_FECHA)
    rotulo(doc, "Referencia")
    parrafo_doc(doc, "Asunto: derecho de petición. Solicitud de copia del expediente de afiliación y "
                     "del acta de doble asesoría de Luis Antonio Aguilar Rodríguez, y reclamación por "
                     "déficit de información en el traslado de régimen pensional.", bold=True,
                size=10.5)
    rotulo(doc, "Identificación")
    parrafo_doc(doc, "%s, identificada con cédula de ciudadanía número %s, en calidad de "
                     "representante legal de EXTINTORES S C M (NIT %s) y empleadora de Luis Antonio "
                     "Aguilar Rodríguez, identificado con cédula de ciudadanía 1073694256, presento "
                     "derecho de petición y formulo reclamación en los siguientes términos."
                % (AUTOR, CC, NIT))
    rotulo(doc, "Hechos")
    for i, t in enumerate([
        "Luis Antonio Aguilar Rodríguez está afiliado a esa administradora en el Régimen de Ahorro "
        "Individual con Solidaridad y cotiza como dependiente a través de EXTINTORES S C M. La fecha "
        "del traslado consta en el expediente que se solicita.",
        "A la fecha cuenta con 1.150 semanas cotizadas y 56 años de edad, y aspira a acceder a la "
        "pensión de vejez en el año 2027.",
        "El afiliado manifiesta que la decisión de traslado se produjo tras una visita de un asesor "
        "comercial, quien le indicó que obtendría una mesada superior, sin que se le entregara "
        "simulación, proyección o documento comparativo de los regímenes.",
        "El afiliado no conserva copia del formulario de afiliación ni recuerda haber suscrito acta "
        "de doble asesoría.",
        "La empresa empleadora no es titular del deber de asesoría, que corresponde a la "
        "administradora conforme a los artículos 20, 48 y 53 de la Constitución Política y a la Ley "
        "1748 de 2014.",
    ], 1):
        numerado_doc(doc, i, t)
    rotulo(doc, "Fundamentos de derecho")
    for t in ["Artículo 23 de la Constitución Política y Ley 1755 de 2015.",
              "Artículos 271 y 272 de la Ley 100 de 1993: ineficacia del traslado por incumplimiento "
              "del deber de información.",
              "Ley 1748 de 2014, artículo 3: deber de doble asesoría como condición previa al "
              "traslado.",
              "Corte Constitucional, Sentencia SU-107 de 2024: estándar de información, "
              "consentimiento informado y equilibrio de la carga probatoria.",
              "Corte Suprema de Justicia, Sala de Casación Laboral, sentencias SL3179-2023 y "
              "SL1179-2024."]:
        bullet(doc, t, size=10)
    rotulo(doc, "Peticiones")
    for i, t in enumerate([
        "Que se expida copia íntegra del expediente de afiliación del solicitante, incluidos el "
        "formulario, las simulaciones, los extractos y las comunicaciones remitidas.",
        "Que se expida copia del acta de doble asesoría o, en su defecto, se certifique que no existe "
        "soporte documental de la asesoría prestada.",
        "Que se certifique el capital acumulado, los rendimientos, las comisiones y las primas "
        "descontadas desde la fecha de afiliación.",
        "Que se revise de oficio la validez del traslado y, si se constata el déficit de información, "
        "se adopten las medidas para restablecer la afiliación al Régimen de Prima Media.",
        "Que se responda de fondo dentro de los quince días hábiles siguientes.",
    ], 1):
        numerado_doc(doc, i, t)
    parrafo_doc(doc, "La presente solicitud constituye la reclamación que se tendrá como surtida para "
                     "efectos de la competencia territorial prevista en el artículo 11 del Código "
                     "Procesal del Trabajo y de la Seguridad Social.", size=10.5)
    parrafo_doc(doc, "Atentamente,", size=10.5, space_after=2)
    firma(doc)
    _guardar(doc, "Peticion_AFP_Porvenir.docx")


# ------------------------------------------------------------- empleador
def solicitud_interna():
    doc = _nuevo("Solicitud de agotamiento de la vía interna")
    encabezado_escrito(doc, "EXTINTORES S C M  ·  GERENCIA GENERAL", CIUDAD_FECHA)
    rotulo(doc, "Referencia")
    parrafo_doc(doc, "Asunto: solicitud de agotamiento de la vía interna y reclamación por "
                     "diferencias en el reconocimiento de prestaciones económicas del Sistema de "
                     "Seguridad Social Integral.", bold=True, size=10.5)
    rotulo(doc, "Identificación")
    parrafo_doc(doc, "%s, identificada con cédula de ciudadanía número %s, en calidad de titular y "
                     "representante legal de EXTINTORES S C M (NIT %s), presento solicitud de "
                     "agotamiento de la vía interna en los siguientes términos."
                % (AUTOR, CC, NIT))
    rotulo(doc, "Hechos")
    for i, t in enumerate([
        "La suscrita es titular y representante legal de la empresa EXTINTORES S C M, empleadora de "
        "tres (3) trabajadores dependientes.",
        "En la auditoría interna practicada en octubre de 2026 se identificaron diferencias en la "
        "liquidación de una licencia de paternidad y en el reporte de novedades de uno de los "
        "trabajadores.",
        "Las diferencias ascienden a %s y afectan el ingreso de Felipe Martínez (C.C. 8686834), por "
        "la licencia de paternidad, así como el reporte de las semanas de los tres trabajadores de "
        "la planta." % money(412000),
        "La empresa reconoce la procedencia de la corrección y requiere formalizar el trámite para "
        "que produzca efectos frente a la EPS %s y a las administradoras de pensiones." % EPS,
    ], 1):
        numerado_doc(doc, i, t)
    rotulo(doc, "Solicitud")
    for i, t in enumerate([
        "Que se reconozca y ordene el pago de la diferencia dejada de percibir por el trabajador "
        "Felipe Martínez, con los ajustes a que haya lugar.",
        "Que se corrija la planilla integrada de liquidación de aportes y se reporten las semanas "
        "omitidas en la historia laboral de los trabajadores de la planta.",
        "Que se expida constancia escrita de la respuesta, con indicación de la fecha y del "
        "responsable de la decisión.",
    ], 1):
        numerado_doc(doc, i, t)
    rotulo(doc, "Fundamentos de derecho")
    for t in ["Artículo 23 de la Constitución Política y Ley 1755 de 2015.",
              "Artículos 236 del Código Sustantivo del Trabajo y 6 de la Ley 2114 de 2021: licencia "
              "de paternidad y base de liquidación.",
              "Artículos 17 y 18 de la Ley 100 de 1993: obligaciones del empleador y base de "
              "cotización.",
              "Decreto 1072 de 2015, parte 2, título 2: reporte de novedades en la planilla "
              "integrada."]:
        bullet(doc, t, size=10)
    parrafo_doc(doc, "La presente solicitud se presenta dentro del término legal y se entiende como "
                     "agotamiento de la vía interna para todos los efectos. Su respuesta deberá "
                     "producirse dentro de los quince (15) días hábiles siguientes, conforme al "
                     "artículo 23 de la Constitución Política y a la Ley 1755 de 2015.", size=10.5)
    parrafo_doc(doc, "Atentamente,", size=10.5, space_after=2)
    firma(doc)
    _guardar(doc, "Solicitud_Agotamiento_Via_Interna.docx")


def main():
    peticion_afp()
    solicitud_interna()


if __name__ == "__main__":
    main()
