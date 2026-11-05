# -*- coding: utf-8 -*-
"""Datos maestros del caso (cifras 2026) y parámetros del documento."""
from decimal import Decimal, ROUND_HALF_UP

SMLMV = 1750905          # Decreto 1469 de 2025 (vigente desde 01/01/2026)
AUX_TRANSPORTE = 249095  # Decreto 1470 de 2025
TOPE_AUX = 2 * SMLMV     # 3.501.810
TOPE_IBC = 25 * SMLMV    # 43.772.625
TOPE_EXO = 10 * SMLMV    # 17.509.050
TOPE_CUANTIA = 20 * SMLMV  # 35.018.100 — art. 12 CPTSS

TARIFAS = {"I": 0.00522, "II": 0.01044, "III": 0.02436, "IV": 0.04350, "V": 0.06960}

# Planta de personal: (código, cargo, tipo de vinculación, IBC/salario, régimen pensional, obs.)
NOMINA = [
    ("EXT-001", "Gerente general (representante legal)", "Término indefinido", 4500000, "Prima Media", "No recibe auxilio de transporte"),
    ("EXT-002", "Coordinador técnico", "Término indefinido", 3500000, "Prima Media", "Supervisa el taller de recarga"),
    ("EXT-003", "Asesor comercial", "Término fijo 1 año", 2400000, "Ahorro Individual", "Caso pensional auditado (P-03)"),
    ("EXT-004", "Técnico de recarga A", "Término indefinido", 2100000, "Prima Media", "Caso de licencia de maternidad"),
    ("EXT-005", "Técnico de recarga B", "Término indefinido", 1900000, "Prima Media", "Caso de AT-2026-014"),
    ("EXT-006", "Conductor-repartidor", "Término fijo 1 año", 2050000, "Prima Media", "Riesgo vial in itinere"),
    ("EXT-007", "Auxiliar de taller", "Término indefinido", 1750905, "Ahorro Individual", "Devenga 1 SMMLV"),
    ("EXT-008", "Auxiliar administrativa", "Término indefinido", 1800000, "Prima Media", "Gestiona novedades de PILA"),
    ("EXT-009", "Auxiliar de bodega", "Término indefinido", 1750905, "Prima Media", "Manipulación de cilindros"),
    ("EXT-010", "Técnico de mantenimiento", "Término fijo 6 meses", 1900000, "Ahorro Individual", "Ingreso 01/07/2026"),
]

CONTRATISTA = {"nombre": "Contador público (prestación de servicios)", "honorarios": 2800000}
PM = sum(1 for r in NOMINA if r[4] == "Prima Media")
RAIS = sum(1 for r in NOMINA if r[4] == "Ahorro Individual")
IBC_TOTAL = sum(r[3] for r in NOMINA)
N_AUX = sum(1 for r in NOMINA if r[3] <= TOPE_AUX)


def cop(v, dec=0):
    """Formato monetario colombiano."""
    if dec:
        s = f"{v:,.{dec}f}"
    else:
        s = f"{v:,.0f}"
    return "$ " + s.replace(",", ".").replace(".", ",", s.count(".") - 1) if False else "$ " + s.replace(",", "X").replace(".", ",").replace("X", ".")


def money(v):
    ent = int(Decimal(str(v)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
    return "$ " + f"{ent:,}".replace(",", ".")


def num(v, dec=2):
    s = f"{v:,.{dec}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def pct(v, dec=3):
    return (num(v * 100, dec) + " %").replace(" %", " %")
