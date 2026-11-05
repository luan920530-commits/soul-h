# soul-h · Proyecto integrador de Seguridad Social (EXTINTORES SCM)

Documento Word técnico-jurídico (plantilla) para el proyecto integrador de la asignatura
**Derecho de la Seguridad Social** (Universidad INCCA de Colombia, 40 % del corte), aplicado a
la empresa privada **EXTINTORES SCM (NIT 39626936-8)**.

## Entregable

`entrega/Auditoria_Seguridad_Social_EXTINTORES_SCM.docx` — informe de ~19.000 palabras, 28 tablas,
5 figuras, 3 escritos procesales modelo y 7 anexos diligenciables.

## Estructura del documento

| Parte | Contenido |
|---|---|
| Preliminares | Portada, tabla de contenido automática, resumen ejecutivo, metodología, introducción |
| Cap. 1 (Eje I) | Caracterización, afiliación y cotización en salud, IBC, cobertura y prestaciones |
| Cap. 2 (Eje II) | SGRL, clasificación del riesgo, caso ATEP AT-2026-014, culpa patronal (art. 216 CST) |
| Cap. 3 (Eje III) | Mapa pensional, Ley 2381 de 2024 / C-264 de 2026, caso P-03, ineficacia del traslado |
| Cap. 4 (Eje IV) | Derecho de petición (ARL y AFP), agotamiento de vía interna, competencia judicial |
| Cap. 5 | Matriz de hallazgos, plan de mejoramiento, conclusiones, sustentación oral, bibliografía |
| Anexos | 7 formatos diligenciables + matriz de correspondencia con la rúbrica |

## Regenerar el .docx

```bash
pip install --break-system-packages python-docx matplotlib
python3 scripts/figuras.py     # genera assets/*.png
python3 scripts/main.py        # compone el .docx
```

- `scripts/docx_engine.py`: estilos, numeración multinivel, tablas, cajas, campos (PAGE/NUMPAGES/TOC)
  y normalización del orden de elementos OOXML (el archivo pasa la validación de esquema).
- `scripts/datos.py`: cifras maestras (SMMLV 2026, topes, nómina, tarifas).
- `scripts/parte1.py` … `parte5.py`: contenido por capítulo.
- `scripts/figuras.py`: figuras vectoriales (IBC, aportes, mapa pensional, ruta ATEP, reparación).

## Nota

Las cifras de nómina y los casos identificados como ilustrativos (AT-2026-014, P-03) deben
sustituirse por los soportes reales de la empresa antes de la presentación definitiva.
