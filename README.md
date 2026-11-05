# soul-h · Proyecto integrador de Seguridad Social (EXTINTORES SCM)

Documento Word técnico-jurídico (plantilla) para el proyecto integrador de la asignatura
**Derecho de la Seguridad Social** (Universidad INCCA de Colombia, 40 % del corte), aplicado a
la empresa privada **EXTINTORES SCM (NIT 39626936-8)**.

## Entregable

`entrega/Auditoria_Seguridad_Social_EXTINTORES_SCM.docx` — informe de ~21.800 palabras, 31 tablas,
5 figuras, 3 escritos procesales modelo y 8 anexos.

Los datos de la organización provienen del **acta de inspección sanitaria 028-050164-26** del
09/09/2026 (Secretaría de Salud de Soacha) y constan en el documento: razón social, C.C./NIT,
CIIU 4759, domicilio en Soacha, planta de tres trabajadores y hallazgos del SG-SST. Las bases de
cotización marcadas con asterisco son valores de referencia mientras se incorporan los
comprobantes de nómina.

## Estructura del documento

| Parte | Contenido |
|---|---|
| Preliminares | Portada, tabla de contenido automática, resumen ejecutivo, metodología, introducción |
| Fuente primaria | Ficha del acta sanitaria 028-050164-26 y su proyección sobre el SG-SST (Anexo 2 y numeral 2.6) |
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
python3 scripts/validar_docx.py entrega/Auditoria_Seguridad_Social_EXTINTORES_SCM.docx
python3 scripts/estilo_es.py entrega/Auditoria_Seguridad_Social_EXTINTORES_SCM.docx \
        --style .skills/house-styles/informe-juridico-es.json
```

- `scripts/docx_engine.py`: estilos, numeración multinivel, tablas, cajas, campos (PAGE/NUMPAGES/TOC)
  y normalización del orden de elementos OOXML (el archivo pasa la validación de esquema).
- `scripts/datos.py`: cifras maestras (SMMLV 2026, topes, nómina, tarifas).
- `scripts/parte1.py` … `parte5.py`: contenido por capítulo.
- `scripts/figuras.py`: figuras vectoriales (IBC, aportes con la exoneración del art. 114-1 E.T.,
  mapa pensional 2+1, ruta ATEP, reparación).
- `scripts/validar_docx.py`: valida el orden de propiedades OOXML, la integridad de partes, los pies
  de figura y la ausencia de marcadores de plantilla sin sustituir.
- `scripts/estilo_es.py`: verificador de estilo en español (ver más abajo).

## Plugin de redacción

El repositorio incluye el skill [`conorbronsdon/avoid-ai-writing`](https://github.com/conorbronsdon/avoid-ai-writing)
(MIT) vendorizado en `.skills/_vendor/`, más una capa propia para español jurídico:

```bash
# señal determinista del plugin (calibrado en inglés: se interpreta con salvedad)
node .skills/_vendor/bin/avoid-ai-writing.js --context technical documento.txt

# verificación en español: relleno, frases largas, monotonía, guiones largos
python3 scripts/estilo_es.py entrega/Auditoria_Seguridad_Social_EXTINTORES_SCM.docx \
        --style .skills/house-styles/informe-juridico-es.json
```

Detalle de la instalación, del perfil de estilo y de la salvedad de idioma:
[`.skills/README.md`](.skills/README.md) y [`.skills/STYLE-PROFILE-es.md`](.skills/STYLE-PROFILE-es.md).

## Nota

Las cifras de nómina y los casos identificados como ilustrativos (AT-2026-014, P-03) deben
sustituirse por los soportes reales de la empresa antes de la presentación definitiva.
