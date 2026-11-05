# Perfil de estilo para este repositorio (es-CO, informe jurídico)

El plugin [`avoid-ai-writing`](https://github.com/conorbronsdon/avoid-ai-writing) (MIT,
© Conor Bronsdon) se diseñó sobre patrones **en inglés**. Su propio `SKILL.md` lo advierte:
*«These rules and the project's detector are calibrated for English; applying them to another
language requires language-specific rules and evidence.»*

Este archivo es esa capa específica para español que el upstream no trae. **No modifica el
código del plugin**: vive al lado y se aplica antes o después de él.

## 1. Qué se toma del plugin

| Componente | Uso en este repositorio |
|---|---|
| `SKILL.md` | Contrato de edición: alcance autorizado, fidelidad a la fuente, contenido protegido, pasadas de iteración. Se aplica íntegro, con las salvedades de idioma del punto 2. |
| `references/patterns.md` | Catálogo de referencia. Los patrones que **sí** son translingüísticos se reutilizan: eliminación de relleno, evitación de metáforas vacías, variedad de longitud de frase, verbos concretos sobre nominalizaciones, no cerrar con ecos. |
| `detector/`, `bin/avoid-ai-writing.js` | Motor determinista. Se usa como **señal**, con `--context technical`, y se interpreta solo después de leer la salvedad de idioma. |
| `scripts/check-style.js` | Verificación de mecánicas del perfil JSON (`quotes`, `emDash`, `latinAbbrev`, `headings`). Con el perfil de este repo la comprobación es parcial: `latinAbbrev` solo reconoce `e.g.`/`i.e.`, que no son uso hispano. |
| `scripts/normalize-quotes.js` | Normalización de comillas antes de entregar. En español se prefiere «» para citas y “” solo como comilla de segundo nivel; el script fija **un** estilo por familia, así que se ejecuta con `--quotes curly` y se revisa a mano el resultado en citas jurídicas. |

## 2. Qué no se puede importar (y por qué)

- **Los puntajes.** El motor mide entropía de trigramas de *function words* inglesas y listas
  léxicas en inglés. Aplicado a español, un texto perfectamente humano puede salir con
  «señales mínimas» o con falsos positivos, y un texto mecánico puede salir limpio. El puntaje
  se usa para **detectar monotonía y densidad de marcas tipográficas**, no como veredicto.
- **El veredicto de autenticidad.** El propio proyecto lo dice: *«signals, not proof»*. Ningún
  puntaje de esta herramienta sirve como prueba de autoría humana o de IA en un contexto
  académico o disciplinario.
- **Las listas de palabras prohibidas en inglés** (`delve`, `tapestry`, `landscape`…). En
  español los equivalentes son otros y están en el archivo
  `.skills/house-styles/informe-juridico-es.json`.

## 2-bis. Perfil de autor (v2 del JSON)

El perfil `informe-juridico-es.json` tiene dos capas y no se mezclan a voluntad:

| Capa | Qué gobierna | Dónde se aplica |
|---|---|---|
| `register` (fija) | Voz impersonal, precisión normativa, prohibición de relleno, adjetivación escasa | Todo el documento, incluidas tablas y matrices |
| `escenas` (marcada) | Primera persona moderada, longitud irregular de párrafo, conectores de matiz, preguntas que abren problema, cierre con idea propia | Resumen ejecutivo, nota de método, comentario jurisprudencial, conclusiones y colofón |

La razón de separarlas es simple: escribir «en mi criterio» dentro de una tabla de aportes no es
estilo, es ruido. La valoración va donde hay algo que valorar.

Lo que el perfil **prohíbe en todo el documento**, y el verificador comprueba:

1. Conectores mecánicos: `Asimismo`, `No obstante`, `Por otro lado`, `En suma`, `En síntesis`.
2. Cierres vacíos: `En conclusión`, `Para resumir`, `Para finalizar`, `Todo lo anterior nos lleva a`.
3. Aperturas de época: `En el mundo actual`, `En la actualidad`, `Hoy en día`, `Vivimos en una era`.
4. Muletillas de énfasis: `Cabe señalar`, `Es importante destacar`, `Vale la pena mencionar`, `Sin duda alguna`.
5. Adverbios huecos: `claramente`, `evidentemente`, `indudablemente`.
6. La estructura `No es X, es Y`: **una sola vez en todo el documento**.
7. Dos párrafos consecutivos que abran con el mismo conector.
8. Un cierre que sea variación de «el futuro dirá».

Y lo que **exige** en las secciones valorativas: al menos un párrafo de dos oraciones o menos y
al menos uno de seis o más; al menos un dato numérico o una escena laboral concreta por sección;
primera persona solo para valorar, nunca para describir.

## 3. Reglas locales (las que sí aplican aquí)

1. **Guion largo (—).** El perfil lo admite «con moderación»: máximo 1 por cada 1.000
   palabras. En este informe la incidencia era de ~4 por 1.000, así que se sustituyó por
   paréntesis o por dos puntos según el caso.
2. **Relleno académico-administrativo prohibido:** *cabe destacar, es importante señalar, es
   fundamental, en la actualidad, hoy en día, en definitiva, en suma, juega un papel crucial,
   sin duda, robusto, holístico, piedra angular, punto de inflexión, de cara a, sumamente
   importante, no solo… sino también* (cuando no aporta contraste real).
3. **Verbo concreto sobre nominalización:** «la ARL pagó» antes que «se efectuó el pago por
   parte de la ARL».
4. **Cierre de párrafo con consecuencia, no con eco.** Nada de «en resumen, podemos decir
   que…».
5. **Contenido protegido:** citas entre comillas, artículos, radicados, cifras y tablas no se
   reescriben. Se reportan, no se editan.
6. **Comillas:** «» en primera instancia; “” para el segundo nivel; nunca comillas rectas en
   el cuerpo del texto.
7. **Latinos:** `e.g.`, `i.e.`, `etc.` solo dentro de paréntesis; en el cuerpo se escribe
   «por ejemplo», «esto es», «y demás».
8. **Siglas:** siempre en mayúscula y sin puntos internos salvo las que legalmente los llevan
   (C.P.T.S.S., C.S.T.).

## 4. Cómo se ejecuta

```bash
# 1) señal determinista del motor del plugin (inglés, se interpreta con la salvedad del punto 2)
node .skills/_vendor/bin/avoid-ai-writing.js --context technical documento.txt

# 2) verificación en español: relleno, frases largas, monotonía de frase, ritmo de párrafo,
#    estructuras binarias repetidas, conectores de apertura repetidos y fórmulas de cierre
python3 scripts/estilo_es.py entrega/Informe.docx \
        --style .skills/house-styles/informe-juridico-es.json

# 3) mecánicas del perfil (comillas, guiones, latinos) sobre un archivo de texto
node .skills/avoid-ai-writing/scripts/check-style.js documento.md \
     --config ../../house-styles/informe-juridico-es.json
```

## 5. Licencia y atribución

El plugin es MIT (© 2026 Conor Bronsdon). El texto completo de la licencia está en
`.skills/_vendor/LICENSE` y la atribución en `.skills/_vendor/NOTICE.md`. Este archivo y
`scripts/estilo_es.py` son propios de este repositorio y no modifican el plugin.
