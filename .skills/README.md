# Plugin de redacción · `avoid-ai-writing`

Instalación local del skill de redacción
[`conorbronsdon/avoid-ai-writing`](https://github.com/conorbronsdon/avoid-ai-writing)
(MIT, © 2026 Conor Bronsdon), más una capa propia para escribir en **español jurídico**.

## Qué quedó instalado

```
.skills/
├── _vendor/                        copia íntegra del upstream v3.38.0 (commit ccc8fc4) — no se edita
│   ├── SKILL.md                    contrato de edición (agentskills.io), con nota de instalación al inicio
│   ├── references/patterns.md      catálogo de patrones (calibrado en inglés)
│   ├── detector/                   motor determinista: patterns.js, validate.js, CATEGORIES.md
│   ├── bin/avoid-ai-writing.js     CLI de auditoría (JSON, solo lectura, exit 0/2)
│   ├── bin/avoid-ai-writing-gate.js  modo "gate" para CI o pre-commit
│   ├── LICENSE · NOTICE.md         licencia MIT y atribución
├── avoid-ai-writing/               skills/avoid-ai-writing del upstream (scripts auxiliares)
│   └── scripts/check-style.js      verificación de mecánicas del perfil de estilo JSON
├── house-styles/
│   └── informe-juridico-es.json    perfil de estilo propio (registro + mecánicas) para este repo
└── STYLE-PROFILE-es.md             qué se importa del plugin, qué no, y las reglas locales

scripts/estilo_es.py                verificador en español (relleno, frases largas, monotonía, guiones)
```

## Uso

### 1. Con un agente (modo instrucciones)

El skill se invoca leyendo su `SKILL.md`. En la práctica: pídale al agente que aplique
`.skills/_vendor/SKILL.md` con el perfil `.skills/house-styles/informe-juridico-es.json`, o
que siga `.skills/STYLE-PROFILE-es.md` (que ya contiene las reglas en español).
Tres modos del upstream: **detect** (solo informa), **rewrite** (reescribe), **edit** (edita el
archivo en su sitio). Para un informe jurídico, empiece siempre por **detect**.

### 2. CLI del detector

```bash
# informe completo o un fragmento (máx. 10.000 palabras por llamada)
node .skills/_vendor/bin/avoid-ai-writing.js --context technical documento.txt

# gate para pre-commit / CI (exit != 0 si supera el umbral)
node .skills/_vendor/bin/avoid-ai-writing-gate.js documento.txt
```

Salida: JSON con `score`, `label`, `issues[]` y `stats`. Es **solo lectura**: no modifica el
archivo. El límite es de 10.000 palabras por invocación; para documentos mayores hay que
partir el texto (`split -n l/2`).

### 3. Verificación en español

```bash
python3 scripts/estilo_es.py entrega/Informe.docx \
        --style .skills/house-styles/informe-juridico-es.json
```

Reporta: guiones largos con contexto, giros de relleno («cabe destacar», «es fundamental»,
«en la actualidad»…), frases de más de 45 palabras, monotonía de longitud de frase (desviación
estándar), párrafos de más de 180 palabras, latinos fuera de paréntesis y siglas en minúscula.
`--json` para procesarlo; `--strict` para que salga con código 1.

### 4. Mecánicas del perfil de estilo

```bash
node .skills/avoid-ai-writing/scripts/check-style.js documento.md \
     --config .skills/house-styles/informe-juridico-es.json
```

Verifica de forma determinista comillas, guiones largos y latinos. (El chequeo `latinAbbrev`
del upstream solo conoce `e.g.`/`i.e.`, que no son uso hispano: se aplica como referencia.)

## Advertencia sobre el idioma

Los patrones y el detector del plugin están **calibrados para inglés**; su propio `SKILL.md` lo
dice y remite a adaptaciones por idioma mantenidas por la comunidad. No existe adaptación
publicada para español. Por eso este repositorio:

- usa el plugin para lo que sí es translingüístico (contrato de edición, fidelidad a la fuente,
  contenido protegido, mecánicas tipográficas, señal de monotonía);
- usa `scripts/estilo_es.py` y el perfil JSON para los patrones léxicos en español;
- **no** usa el `score` como prueba de autoría. El upstream lo advierte: *signals, not proof*.

## Requisitos

- Node.js ≥ 18 (probado con v22.22.3). El plugin no tiene dependencias externas.
- Python 3 + `python-docx` solo para `scripts/estilo_es.py` sobre `.docx`.

## Actualizar el plugin

```bash
git clone --depth 1 https://github.com/conorbronsdon/avoid-ai-writing.git /tmp/aaw
rsync -a --delete /tmp/aaw/detector/ /tmp/aaw/bin/ .skills/_vendor/ 2>/dev/null
cp /tmp/aaw/SKILL.md /tmp/aaw/references/patterns.md .skills/_vendor/
# volver a insertar la nota de instalación al inicio de .skills/_vendor/SKILL.md
```

La versión instalada está en el campo `version:` del frontmatter de `SKILL.md`.
