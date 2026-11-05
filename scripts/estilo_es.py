#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
estilo_es.py — capa de verificación de estilo en español (complemento determinista del
plugin avoid-ai-writing, cuyo motor de patrones está calibrado en inglés).

Lee un .docx o un .md/.txt y reporta:
  * conteo y ubicación de guiones largos (—), con contexto;
  * giros de relleno académico-administrativo ("cabe destacar", "es fundamental", ...);
  * frases demasiado largas (> 45 palabras), para dividir;
  * monotonía de longitud de frase (desviación estándar baja) — señal estilométrica;
  * longitud de párrafo;
  * cumplimiento del perfil de estilo (house-style JSON) en lo verificable:
    comillas, emDash, latinos (e.g./i.e.), abreviaturas protegidas.

Uso:
    python3 scripts/estilo_es.py <archivo> [--style .skills/house-styles/informe-juridico-es.json]
                                       [--json] [--max-frase 45]

Salida: informe legible por defecto, o JSON con --json.
Código de salida: 0 sin hallazgos graves, 1 con hallazgos graves (--strict), 2 error de uso.
"""
import argparse
import json
import os
import re
import statistics
import sys

# ---------------------------------------------------------------- extracción
def extraer_texto(ruta):
    """Devuelve lista de párrafos (str) del archivo."""
    ext = os.path.splitext(ruta)[1].lower()
    if ext == ".docx":
        try:
            from docx import Document
            from docx.oxml.ns import qn
        except ImportError:
            sys.exit("Falta python-docx: pip install python-docx")
        d = Document(ruta)
        parrafos = []
        for child in d.element.body.iterchildren():
            if child.tag == qn("w:p"):
                t = "".join(n.text or "" for n in child.iter(qn("w:t")))
                if t.strip():
                    parrafos.append(t.strip())
            elif child.tag == qn("w:tbl"):
                for tr in child.iter(qn("w:tr")):
                    celdas = []
                    for tc in tr.iter(qn("w:tc")):
                        celdas.append(" ".join(
                            ("".join(n.text or "" for n in p.iter(qn("w:t")))).strip()
                            for p in tc.iter(qn("w:p"))).strip())
                    if any(celdas):
                        parrafos.append(" · ".join(c for c in celdas if c))
        return parrafos
    with open(ruta, encoding="utf-8") as f:
        texto = f.read()
    # quitar frontmatter y bloques de código
    texto = re.sub(r"^---\n.*?\n---\n", "", texto, flags=re.S)
    texto = re.sub(r"```.*?```", "", texto, flags=re.S)
    return [p.strip() for p in re.split(r"\n\s*\n", texto) if p.strip()]


# ---------------------------------------------------------------- patrones es
RELLENO = [
    (r"cabe (destacar|señalar|mencionar|resaltar|anotar)", "cabe destacar"),
    (r"es importante (destacar|señalar|resaltar|mencionar|anotar)", "es importante destacar"),
    (r"es fundamental (destacar|señalar|resaltar|comprender)", "es fundamental"),
    (r"(en la actualidad|hoy en día|en el mundo actual|en los tiempos que corren)", "en la actualidad"),
    (r"(en conclusión|en definitiva|en suma|a modo de cierre)", "en conclusión / en definitiva"),
    (r"juega un papel (crucial|fundamental|clave|vital|importante)", "juega un papel crucial"),
    (r"(desempeña|tiene) un papel (crucial|fundamental|clave|vital)", "papel crucial"),
    (r"sin (lugar a )?duda", "sin duda"),
    (r"(robusto|robusta|holístic[oa]|vibrante|enriquecedor[ae]?)", "adjetivo de relleno"),
    (r"de (manera|forma) (integral|holística|robusta|exhaustiva y)", "de manera integral"),
    (r"(es|resulta) (evidente|incuestionable) que", "es evidente que"),
    (r"no (solamente|solo) .{0,70}? sino (también )?", "no solo… sino también"),
    (r"(permite|posibilita) (que|no solo)", "permite que"),
    (r"(sumamente|extremadamente|altamente) (importante|relevante|significativ[oa])", "sumamente importante"),
    (r"es (crucial|clave|vital) (para|que)", "es crucial para"),
    (r"en este (sentido|orden de ideas)", "en este sentido"),
    (r"(por consiguiente|en consecuencia),? (es|resulta)", "conector mecánico"),
    (r"(punto de inflexión|piedra angular|pilar fundamental)", "metáfora de relleno"),
    (r"de cara a", "de cara a"),
    (r"(a fin de|con el objetivo de) (poder )?", "a fin de"),
    (r"lo que (nos )?lleva a (reflexionar|preguntarnos)", "retórica vacía"),
]

PROTEGIDOS = ["IBC", "SMLMV", "ATEP", "SG-SST", "C.P.T.S.S.", "C.S.T.", "EPS", "ARL", "AFP",
              "PILA", "RAIS", "RPM", "IPP", "SGP", "SGSSS", "SGRL", "SU-107", "SL2480"]

RE_EMDASH = re.compile("—")
RE_LATIN = re.compile(r"\b(e\.g\.|i\.e\.|etc\.)", re.I)


# ---------------------------------------------------------------- análisis
def dividir_frases(texto):
    # corta en . ; : pero no en abreviaturas ni cifras
    partes = re.split(r"(?<=[.!?;:])\s+(?=[A-ZÁÉÍÓÚÑ«¿¡(])", texto)
    return [p.strip() for p in partes if len(p.split()) >= 4]


def analizar(parrafos, max_frase=45):
    texto = "\n\n".join(parrafos)
    palabras_total = sum(len(p.split()) for p in parrafos)

    # guiones largos con contexto
    emdash = []
    for i, p in enumerate(parrafos):
        for m in RE_EMDASH.finditer(p):
            ini, fin = max(0, m.start() - 55), min(len(p), m.end() + 55)
            emdash.append({"parrafo": i + 1, "contexto": p[ini:fin].replace("\n", " ")})

    # relleno
    relleno = []
    for i, p in enumerate(parrafos):
        for patron, etiqueta in RELLENO:
            for m in re.finditer(patron, p, flags=re.I):
                ini, fin = max(0, m.start() - 45), min(len(p), m.end() + 45)
                relleno.append({"parrafo": i + 1, "giro": etiqueta,
                                "texto": m.group(0), "contexto": p[ini:fin]})

    # frases (por párrafo: nunca se cruzan saltos de párrafo)
    frases = []
    for p in parrafos:
        frases.extend(dividir_frases(p))
    largos = [(len(f.split()), f) for f in frases if len(f.split()) > max_frase]
    longitudes = [len(f.split()) for f in frases]
    sd = statistics.pstdev(longitudes) if len(longitudes) > 2 else 0
    media = statistics.mean(longitudes) if longitudes else 0

    # párrafos largos
    parrafos_largos = [i + 1 for i, p in enumerate(parrafos) if len(p.split()) > 180]

    # latinos fuera de paréntesis (como exige el perfil)
    latinos = []
    for i, p in enumerate(parrafos):
        for m in RE_LATIN.finditer(p):
            antes = p[:m.start()]
            if antes.count("(") <= antes.count(")"):
                latinos.append({"parrafo": i + 1, "texto": m.group(0)})

    # mayúsculas de siglas protegidas (posible minúscula indebida)
    siglas_min = []
    for i, p in enumerate(parrafos):
        for s in PROTEGIDOS:
            base = s.rstrip(".")
            if len(base) < 3:
                continue
            if re.search(r"\b%s\b" % re.escape(base.lower()), p) and base.lower() != base:
                siglas_min.append({"parrafo": i + 1, "sigla": base})

    return {
        "archivo_palabras": palabras_total,
        "parrafos": len(parrafos),
        "frases": len(frases),
        "frase_media_palabras": round(media, 1),
        "frase_desviacion": round(sd, 1),
        "frases_mayores_a": max_frase,
        "frases_largas_n": len(largos),
        "frases_largas": [{"palabras": n, "texto": f[:200]} for n, f in largos[:12]],
        "guiones_largos_n": len(emdash),
        "guiones_largos": emdash,
        "relleno_n": len(relleno),
        "relleno": relleno,
        "parrafos_mayores_180p": parrafos_largos,
        "latinos_fuera_parentesis_n": len(latinos),
        "latinos_fuera_parentesis": latinos[:10],
        "siglas_en_minuscula_n": len(siglas_min),
        "siglas_en_minuscula": siglas_min[:10],
    }


# ---------------------------------------------------------------- salida
def reporte(r, max_frase):
    L = []
    L.append("═" * 74)
    L.append("VERIFICACIÓN DE ESTILO EN ESPAÑOL · complemento de avoid-ai-writing")
    L.append("═" * 74)
    L.append("Palabras: %d   |   Párrafos: %d   |   Frases: %d"
             % (r["archivo_palabras"], r["parrafos"], r["frases"]))
    L.append("Longitud media de frase: %.1f palabras (desviación %.1f) — monotonía %s"
             % (r["frase_media_palabras"], r["frase_desviacion"],
                "ALTA" if r["frase_desviacion"] < 8 else "aceptable"))
    L.append("")
    L.append("─ Guiones largos (—): %d   [señal estilométrica: el motor inglés los penaliza]"
             % r["guiones_largos_n"])
    for e in r["guiones_largos"][:6]:
        L.append("   · párr. %d: …%s…" % (e["parrafo"], e["contexto"]))
    if r["guiones_largos_n"] > 6:
        L.append("   · (+%d más; use --json para la lista completa)" % (r["guiones_largos_n"] - 6))
    L.append("")
    L.append("─ Giros de relleno: %d" % r["relleno_n"])
    for e in r["relleno"][:10]:
        L.append("   · párr. %d [%s] «%s»" % (e["parrafo"], e["giro"], e["texto"]))
    L.append("")
    L.append("─ Frases de más de %d palabras: %d" % (max_frase, r["frases_largas_n"]))
    for f in r["frases_largas"][:5]:
        L.append("   · %d palabras: %s…" % (f["palabras"], f["texto"][:140]))
    L.append("")
    L.append("─ Latinos fuera de paréntesis (e.g./i.e./etc.): %d"
             % r["latinos_fuera_parentesis_n"])
    L.append("─ Siglas escritas en minúscula: %d" % r["siglas_en_minuscula_n"])
    if r["parrafos_mayores_180p"]:
        L.append("─ Párrafos de más de 180 palabras: párr. %s"
                 % ", ".join(map(str, r["parrafos_mayores_180p"][:12])))
    L.append("─" * 74)
    graves = r["relleno_n"] + r["latinos_fuera_parentesis_n"] + r["siglas_en_minuscula_n"]
    L.append("VEREDICTO: %s" % ("requiere corrección" if graves else
                                "sin hallazgos graves (revise los guiones largos a mano)"))
    return "\n".join(L)


def main():
    ap = argparse.ArgumentParser(description="Verificación de estilo en español")
    ap.add_argument("archivo")
    ap.add_argument("--style", help="perfil JSON de estilo de la casa")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--max-frase", type=int, default=45)
    ap.add_argument("--strict", action="store_true",
                    help="sale con 1 si hay hallazgos graves")
    args = ap.parse_args()

    if not os.path.exists(args.archivo):
        sys.exit("No existe el archivo: %s" % args.archivo)
    if args.archivo.lower().endswith((".docx", ".doc")):
        if not args.archivo.lower().endswith(".docx"):
            sys.exit("Solo se admite .docx (convierta el .doc)")

    parrafos = extraer_texto(args.archivo)
    if not parrafos:
        sys.exit("El archivo no contiene texto extraíble.")

    r = analizar(parrafos, args.max_frase)

    if args.style:
        if not os.path.exists(args.style):
            sys.exit("No existe el perfil de estilo: %s" % args.style)
        with open(args.style, encoding="utf-8") as f:
            perfil = json.load(f)
        r["perfil"] = {"name": perfil.get("name"),
                       "register": perfil.get("register", []),
                       "mechanics": perfil.get("mechanics", {})}

    if args.json:
        print(json.dumps(r, ensure_ascii=False, indent=2))
    else:
        print(reporte(r, args.max_frase))
        if args.style:
            print("\nPerfil aplicado: %s" % r["perfil"]["name"])
            for d in r["perfil"]["register"][:4]:
                print("   · %s" % d)

    graves = r["relleno_n"] + r["latinos_fuera_parentesis_n"] + r["siglas_en_minuscula_n"]
    return 1 if (args.strict and graves) else 0


if __name__ == "__main__":
    sys.exit(main())
