# -*- coding: utf-8 -*-
"""Datos maestros del caso (cifras 2026) y parámetros del documento.

Fuentes primarias de la caracterización:
  * Acta de inspección, vigilancia y control sanitario N.° 028-050164-26 del 09/09/2026,
    Secretaría de Salud de Soacha (formato SL-FR-028 v.1).
  * Registro Único Tributario — cédula/NIT 39626936-8.
Las bases de cotización marcadas con `por_confirmar=True` son valores de referencia mientras
la empresa aporta los comprobantes de nómina; el documento las rotula como tales.
"""
from decimal import Decimal, ROUND_HALF_UP

SMLMV = 1750905          # Decreto 1469 de 2025 (vigente desde 01/01/2026)
AUX_TRANSPORTE = 249095  # Decreto 1470 de 2025
TOPE_AUX = 2 * SMLMV     # 3.501.810
TOPE_IBC = 25 * SMLMV    # 43.772.625
TOPE_EXO = 10 * SMLMV    # 17.509.050
TOPE_CUANTIA = 20 * SMLMV  # 35.018.100 — art. 12 CPTSS

TARIFAS = {"I": 0.00522, "II": 0.01044, "III": 0.02436, "IV": 0.04350, "V": 0.06960}

# ---------------------------------------------------------------- identificación
EMPRESA = {
    "razon_social": "Julia Inés Buitrago García",
    "nombre_comercial": "EXTINTORES S C M",
    "documento": "39626936-8",       # cédula/NIT según acta 028-050164-26
    "cc_titular": "39626936",
    "inscripcion": "25 754 39626936-8",
    "direccion": "KR 4 60 17",
    "municipio": "Soacha",
    "departamento": "Cundinamarca",
    "comuna": "Comuna 3",
    "barrio": "La Despensa",
    "telefono": "3212944603",
    "correo": "julia_ines_buitrago@hotmail.com",
    "ciiu": "4759",
    "ciiu_desc": "Comercio al por menor de otros artículos domésticos en establecimientos "
                 "especializados",
    "horario": "Lunes a viernes, 9:00 a.m. a 9:00 p.m.",
    "codigo_dane": "25 754",
    "matricula_inmobiliaria": "No consta en el acta",
}

ACTA = {
    "numero": "028-050164-26",
    "formato": "SL-FR-028, versión 1 (aprobado el 07/10/2022)",
    "entidad": "Secretaría de Salud de Soacha",
    "fecha": "09/09/2026",
    "dias_plazo": 22,
    "fecha_limite": "09/10/2026",
    "concepto_anterior": {"fecha": "18/10/2024", "concepto": "Favorable con requerimientos",
                          "cumplimiento": 57.50},
    "funcionaria": "Gissel Paola Hurtado Leyton, C.C. 1013644855, ingeniera química, "
                   "contrato 3136-2026",
    "recibio": "Felipe Martínez, C.C. 8686834, vendedor",
    "operativos_declarados": 1,
}

# ---------------------------------------------------------------- planta de personal
# Cada trabajador: código, nombre, documento, cargo, vinculación, IBC, régimen, observación
TRABAJADORES = [
    {
        "codigo": "WT-01",
        "nombre": "Julia Inés Buitrago García",
        "documento": "C.C. 39626936",
        "cargo": "Propietaria y representante legal (gerente)",
        "vinculacion": "Término indefinido",
        "ibc": 4500000,
        "regimen": "Prima Media",
        "obs": "Titular de la empresa unipersonal; no recibe auxilio de transporte",
        "por_confirmar": True,
    },
    {
        "codigo": "WT-02",
        "nombre": "Luis Antonio Aguilar Rodríguez",
        "documento": "C.C. 1073694256",
        "cargo": "Técnico de recarga y mantenimiento (operativo)",
        "vinculacion": "Término indefinido",
        "ibc": 1900000,
        "regimen": "Ahorro Individual",
        "obs": "Único operativo reportado en el acta; caso AT-2026-014 y caso pensional P-03",
        "por_confirmar": True,
    },
    {
        "codigo": "WT-03",
        "nombre": "Felipe Martínez",
        "documento": "C.C. 8686834",
        "cargo": "Vendedor (a) de mostrador",
        "vinculacion": "Término indefinido",
        "ibc": 1750905,
        "regimen": "Prima Media",
        "obs": "Devenga un SMMLV; recibe auxilio de transporte; caso de licencia de paternidad",
        "por_confirmar": True,
    },
]

# Compatibilidad con el motor de figuras y tablas: (cód, cargo, vinculación, IBC, régimen, obs)
NOMINA = [(t["codigo"], t["cargo"], t["vinculacion"], t["ibc"], t["regimen"], t["obs"])
          for t in TRABAJADORES]

CONTRATISTA = {"nombre": "Contador público (prestación de servicios)", "honorarios": 2800000}
PM = sum(1 for r in NOMINA if r[4] == "Prima Media")
RAIS = sum(1 for r in NOMINA if r[4] == "Ahorro Individual")
IBC_TOTAL = sum(r[3] for r in NOMINA)
N_AUX = sum(1 for r in NOMINA if r[3] <= TOPE_AUX)

# ---------------------------------------------------------------- aportes 2026
# Art. 114-1 E.T.: persona natural con dos o más trabajadores, todos por debajo de 10 SMMLV
EXONERADO_114_1 = len(NOMINA) >= 2 and all(t["ibc"] < TOPE_EXO for t in TRABAJADORES)
TASA_SALUD_EMPLEADOR = 0.0 if EXONERADO_114_1 else 0.085
TASA_SALUD_TRABAJADOR = 0.04
TASA_PENSION = 0.12
TASA_CAJA = 0.04
TASA_SENA = 0.0 if EXONERADO_114_1 else 0.02
TASA_ICBF = 0.0 if EXONERADO_114_1 else 0.03

APORTE_SALUD_EMPLEADOR = IBC_TOTAL * TASA_SALUD_EMPLEADOR
APORTE_SALUD_TRABAJADOR = IBC_TOTAL * TASA_SALUD_TRABAJADOR
APORTE_PENSION = IBC_TOTAL * TASA_PENSION
APORTE_ARL = IBC_TOTAL * TARIFAS["I"]
APORTE_CAJA = IBC_TOTAL * TASA_CAJA

# ---------------------------------------------------------------- casos
IBC_CASO = 1900000          # WT-02 · accidente de trabajo AT-2026-014
DIAS_IT = 42
IPP = 0.14                  # 14 % de pérdida de capacidad laboral
SUBSIDIO_IT = IBC_CASO / 30 * DIAS_IT
INDEMNIZACION_IPP = IPP * 24 * IBC_CASO
TARIFADA = SUBSIDIO_IT + INDEMNIZACION_IPP
DANO_EMERGENTE = 3200000
LUCRO_CESANTE = 2700000
PERJUICIO_MORAL = 15000000
REPARACION_PLENA = DANO_EMERGENTE + LUCRO_CESANTE + PERJUICIO_MORAL + INDEMNIZACION_IPP
NETO_EMPLEADOR = REPARACION_PLENA - TARIFADA

# Caso pensional P-03 (WT-02)
IBL_P03 = 3200000
SEMANAS_P03 = 1150
CAPITAL_RAIS = 268000000
MESADA_RPM = IBL_P03 * 0.655
MESADA_RAIS_SIN_GARANTIA = CAPITAL_RAIS / (20 * 12)
DIF_MENSUAL = MESADA_RPM - SMLMV


def cop(v, dec=0):
    """Formato monetario colombiano."""
    return money(v) if not dec else "$ " + num(v, dec)


def money(v):
    ent = int(Decimal(str(v)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
    return "$ " + f"{ent:,}".replace(",", ".")


def num(v, dec=2):
    s = f"{v:,.{dec}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def pct(v, dec=3):
    return num(v * 100, dec) + " %"
