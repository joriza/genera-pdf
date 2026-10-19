# generar-pdf

Pipeline **canónico** y **determinista** para generar PDFs a partir de archivos
`.md` usando **pandoc + Chrome/Edge headless** (motor Blink). Vive en esta
carpeta: es el único lugar donde está la plantilla y el script, y ningún
proyecto mantiene copias propias.

## Qué es

```
pandoc : <doc>.md ──> _gen.tmp.html  (con plantilla-pdf.css)
chrome  : _gen.tmp.html ──> <doc>.pdf
         normalizar: reemplaza las fechas del PDF por una constante
```

El markdown se lee **sin matemáticas TeX** (`-tex_math_dollars` desactivado):
los `$` del texto son literales y nunca fórmulas. Los `.md` de esta familia
no usan LaTeX, por regla del estándar.

## Requisitos

- **pandoc** en el `PATH` (≥ 2.11; probado con 3.10.1).
- **Chrome** o **Edge** instalados (ambos usan Blink). El script los busca en
  las ubicaciones normales de Windows.
- **Windows PowerShell 5.1+** (`powershell.exe`).
- La plantilla `plantilla-pdf.css` debe estar junto a este script (no la
  muevas).

## Uso

```powershell
# todos los .md de la carpeta actual
.\generar-pdf.ps1

# un solo documento (nombre sin .md; acepta lista separada por comas)
.\generar-pdf.ps1 -Doc clase-1

# otra carpeta, sin moverte de la tuya
.\generar-pdf.ps1 -Dir D:\cursos\excel

# generación 2 veces + comparación de hashes SHA256 (prueba de determinismo)
.\generar-pdf.ps1 -Verify

# ayuda
.\generar-pdf.ps1 -h
```

Los `.md` de entrada y los `.pdf` de salida van a la misma carpeta (la actual,
o la de `-Dir`). Los archivos temporales (`_gen.tmp.html`, `_gen.tmp.css`) se
borran siempre, incluso si el script falla.

## Parámetros

| Parámetro | Valor por defecto | Notas |
| --- | --- | --- |
| `-Doc <nombres>` | todos los `.md` | Sin extensión; acepta lista por coma |
| `-Dir <carpeta>` | carpeta actual | Entrada y salida |
| `-Orientacion vertical\|apaisado` | `vertical` | `apaisado` → `A4 landscape` (trifoliares, tableros) |
| `-Columnas <1\|2\|3>` | `1` | 2 solo a pedido; 3 para trifoliar |
| `-MargenVertical <dist>` | `1.2cm` | Arriba y abajo (`1.2cm`, `15mm`, `1in`…) |
| `-MargenLomo <dist>` | `1.5cm` | Modo espejado (el de fábrica) |
| `-MargenExterno <dist>` | `1cm` | Modo espejado |
| `-MargenLateral <dist>` | `2.5cm` | Pasa a **modo simétrico** (un solo lado); excluyente con `-MargenLomo` |
| `-SinEncabezado` | — | QUITA el nombre del documento de la cabecera |
| `-Verify` | — | Genera 2 veces y compara hashes |
| `-h` / `-Ayuda` | — | Muestra esta ayuda y no genera nada |

### Reglas fijas del estándar

- La plantilla vive aquí: ningún proyecto lleva copias ni CSS propio.
- Los `.md` **no usan LaTeX**: los `$` del texto son literales.
- El pie lleva la fecha del día (ISO-8601, solo fecha): dos corridas del **mismo
  día** dan PDF byte a byte iguales.
- **Trifoliar**: caras con `::: {.cara .primera}` en el `.md` +
  `-Orientacion apaisado`. Un trifoliar apaisado son 2 páginas (frente y verso)
  de 3 columnas: impresas al doble dan los 6 paneles en una sola hoja.

## ¿Por qué Chrome (y no wkhtmltopdf)?

wkhtmltopdf 0.12.6 usa QtWebKit 4.8 (2011) y desobedece la plantilla: escala
las longitudes del CSS ~0,80, ignora `@page`, ignora `line-height`, ignora
`column-count` y no hace fallback por glifo. **Blink** respeta todo. Además,
Blink es determinista byte a byte tras normalizar las fechas de metadatos.

## Determinismo

Chrome escribe `CreationDate` y `ModDate` en los metadatos: dos fechas. Sin
normalizar, dos corridas dan archivos distintos. El paso 3 del script las
reemplaza por una constante de la **misma longitud** (23 caracteres),
preservando los offsets internos del PDF. Cambiar la fecha por una de otra
longitud desplazaría la tabla xref y corrompería el archivo.

Con `-Verify` el script genera cada PDF dos veces y compara los hashes SHA256.

## Estructura

```
generar-pdf.ps1   script principal (pipeline: pandoc → chrome → normalizar)
plantilla-pdf.css plantilla canónica del PDF
README.md        este archivo
```

## Autor

Jorge J. Izaguirre — <joriza@gmail.com>