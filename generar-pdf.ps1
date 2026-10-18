<#
.SYNOPSIS
    Genera PDFs deterministas a partir de los .md de una carpeta, con Chrome
    headless (motor Blink). Pipeline canonico y unico: vive en esta carpeta,
    se invoca desde cualquier proyecto y ningun proyecto mantiene copias.

.DESCRIPTION
    Pipeline:
      1. pandoc : <doc>.md -> _gen.tmp.html (con plantilla-pdf.css)
      2. chrome : _gen.tmp.html -> <doc>.pdf
      3. normalizar: reemplaza las fechas del PDF por una constante

    CONTRATO CON LA PLANTILLA
    Todo lo variable por generacion vive en cuatro parches que el script hace
    sobre plantilla-pdf.css y que la plantilla auto-documenta en su encabezado:
    margenes del @page, [[ENCABEZADO]] con el nombre del archivo, [[ESPEJADO]]
    con los margenes espejados + fecha ISO-8601 en la esquina del lomo, y
    column-count. Cada parche verifica cuantas veces matchea y CORTA CON ERROR
    si no matchea exactamente una: un reemplazo fallido dejaria el PDF con los
    valores de la plantilla y el flag pareceria ignorado en silencio.

    POR QUE CHROME (y no wkhtmltopdf)
    wkhtmltopdf 0.12.6 usa QtWebKit 4.8 (2011) y desobedece la plantilla:
    escala las longitudes del CSS ~0,80, ignora @page, ignora line-height,
    ignora column-count y no hace fallback por glifo. Blink respeta todo.
    Ademas Blink es determinista byte a byte tras normalizar fechas.

    POR QUE ES NECESARIO EL PASO 3
    Chrome escribe CreationDate y ModDate en los metadatos: DOS fechas. Sin
    normalizar, dos corridas dan archivos distintos. El paso 3 las reemplaza
    por una constante de la MISMA longitud (23 caracteres), preservando los
    offsets internos del PDF. No cambiar la fecha por una de otra longitud:
    desplazaria la tabla xref y corromperia el archivo.

    DETERMINISMO CON LA FECHA DEL PIE
    La fecha de generacion va en el pie, en la esquina del lomo (abajo a la
    izquierda en la hoja impar, abajo a la derecha en la par; en modo
    simetrico, abajo a la izquierda), en ISO-8601 y solo la fecha. La inyecta
    el script: Blink no implementa string-set. No rompe el determinismo: dos
    corridas del MISMO dia dan bytes identicos (la fecha es parte de la
    entrada, igual que el nombre del archivo del encabezado).

    CANTIDAD DE COLUMNAS
    UNA columna por defecto en todos los documentos; dos columnas solo a
    pedido explicito con -Columnas 2. El texto corrido gana legibilidad a dos
    columnas (lineas de ~41-46 caracteres), pero una tabla partida o apretada
    en una columna angosta se lee mal. Los elementos de ancho completo (titulo
    principal y tablas) llevan column-span: all en la plantilla.

.PARAMETER Doc
    Nombres de documento sin la extension .md, relativos a -Dir (o a la
    carpeta actual). Si se omite, se generan todos los .md de la carpeta.

.PARAMETER Dir
    Carpeta con los .md de entrada y destino de los PDF. Por defecto, la
    carpeta actual al invocar el script.

.PARAMETER MargenLateral
    Activa el modo SIMETRICO: el mismo margen a los dos costados, para
    imprimir de un solo lado. Con este flag no se generan los bloques de
    margenes espejados y es excluyente con -MargenLomo; -MargenExterno solo
    tiene sentido en modo espejado.

.PARAMETER MargenVertical
    Margen superior e inferior del PDF, con unidad (1.5cm, 15mm, 1in).
    Por defecto: 1.5cm, el valor declarado en el @page de la plantilla.

.PARAMETER MargenLomo
    Margen del lado del lomo (en la hoja impar, la izquierda; en la par, la
    derecha). Por defecto: 1.5cm (estandar desde 2026-10-17). Los margenes
    vienen ESPEJADOS por defecto: los documentos terminan juntos en un
    anillado y el lomo cae siempre del mismo lado, aunque el documento sea de
    una sola hoja. Es la unica forma de regenerar que da un juego
    encuadernable en todos los casos.

.PARAMETER MargenExterno
    Margen del lado externo (el opuesto al lomo) en el modo espejado.
    Por defecto: 1cm. No tiene sentido en modo simetrico.

.PARAMETER Columnas
    Cantidad de columnas del texto corrido, para todos los documentos. Por
    defecto: 1. Dos columnas solo a pedido explicito (-Columnas 2).

.PARAMETER SinEncabezado
    No imprime el encabezado con el nombre del documento en el margen
    superior. Por defecto el encabezado va.

.PARAMETER Verify
    Genera cada PDF dos veces y compara los hashes SHA256 para comprobar el
    determinismo. Tarda el doble.

.EXAMPLE
    generar-pdf.ps1                          # los .md de la carpeta actual
    generar-pdf.ps1 -Verify                  # doble generacion + hashes
    generar-pdf.ps1 -Doc clase-1             # un solo documento
    generar-pdf.ps1 -Dir D:\cursos\excel     # otra carpeta
    generar-pdf.ps1 -Columnas 2              # dos columnas, solo a pedido
    generar-pdf.ps1 -SinEncabezado
    generar-pdf.ps1 -MargenLateral 2.5cm     # modo simetrico (un solo lado)
#>

[CmdletBinding()]
param(
    [string[]]$Doc,
    [string]$Dir,
    [string]$MargenLateral = '2.5cm',
    [string]$MargenVertical = '1.5cm',
    [string]$MargenLomo = '1.5cm',
    [string]$MargenExterno = '1cm',
    [int]$Columnas = 1,
    [switch]$SinEncabezado,
    [switch]$Verify
)

$ErrorActionPreference = 'Stop'

# Carpeta de trabajo: -Dir, o la carpeta actual del que invoca. Se resuelve a
# ruta absoluta SIEMPRE, porque .NET (WriteAllText, ReadAllBytes) resuelve los
# nombres relativos contra el CWD del proceso, que NO sigue a Set-Location
# (defecto medido: el PDF salia sin estilos y sin aviso).
if ($Dir) {
    $Destino = (Resolve-Path -LiteralPath $Dir).Path
} else {
    $Destino = (Get-Location).Path
}

# El modo por defecto son margenes ESPEJADOS: el kit termina junto en un
# anillado, asi que el lomo tiene que caer siempre del mismo lado, aunque
# algun documento sea de una sola hoja. -MargenLateral es la salida explicita
# al modo simetrico, para la impresion de un solo lado.
$simetrico = $PSBoundParameters.ContainsKey('MargenLateral') -and -not $PSBoundParameters.ContainsKey('MargenLomo')

$validos = @(@('MargenLateral', $MargenLateral), @('MargenVertical', $MargenVertical), @('MargenLomo', $MargenLomo), @('MargenExterno', $MargenExterno))
foreach ($par in $validos) {
    if ($par[1] -notmatch '^\d+(\.\d+)?(cm|mm|in|pt|px)$') {
        throw "$($par[0]) '$($par[1])' no tiene un formato valido. Ejemplos: 2.5cm, 25mm, 1in."
    }
}
if ($PSBoundParameters.ContainsKey('MargenLateral') -and $PSBoundParameters.ContainsKey('MargenLomo')) {
    throw "-MargenLateral y -MargenLomo son excluyentes: con margenes espejados el ancho del texto sale de -MargenLomo y -MargenExterno, y -MargenLateral quedaria ignorado en silencio."
}
if ($simetrico -and $PSBoundParameters.ContainsKey('MargenExterno')) {
    throw "-MargenExterno solo tiene sentido con margenes espejados, que es el modo por defecto: para el modo simetrico use solo -MargenLateral."
}
if ($Columnas -lt 1 -or $Columnas -gt 3) {
    throw "Columnas tiene que estar entre 1 y 3. Dos columnas solo a pedido explicito: -Columnas 2."
}

$CSS = Join-Path $PSScriptRoot 'plantilla-pdf.css'
if (-not (Test-Path -LiteralPath $CSS)) {
    throw "No se encontro la plantilla en $CSS. Vive junto a este script."
}

$Docs = if ($PSBoundParameters.ContainsKey('Doc')) {
    @($Doc)
} else {
    @(Get-ChildItem -LiteralPath $Destino -Filter '*.md' | Sort-Object Name | ForEach-Object { [System.IO.Path]::GetFileNameWithoutExtension($_.Name) })
}
if (-not $Docs -or @($Docs).Count -eq 0) {
    throw "No hay .md para generar en $Destino. Pase -Doc <nombre> o ubique la carpeta con las fuentes."
}

# Ruta del ejecutable de Chrome. Se prueban las dos ubicaciones normales y Edge
# como ultimo recurso (Blink en los tres casos).
$Chrome = @(
    "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
    "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
    "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe",
    "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe"
) | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $Chrome) { throw "No se encontró Chrome ni Edge." }

# Perfil de Chrome fuera de las carpetas de trabajo, para no ensuciarlas.
$Perfil = Join-Path $env:TEMP 'gentle-ai-chrome-perfil'
New-Item -ItemType Directory -Path $Perfil -Force | Out-Null

# Fecha fija, 23 caracteres: D:YYYYMMDDHHmmSS + HH'MM'
$FechaFija   = "D:20260101000000+00'00'"
$PatronFecha = "D:\d{14}[+-]\d{2}'\d{2}'"

if (-not (Get-Command 'pandoc' -ErrorAction SilentlyContinue)) {
    throw "Falta 'pandoc' en el PATH."
}

function Set-FechaDeterministica {
    param([Parameter(Mandatory)][string]$Ruta)

    # Latin1 (ISO-8859-1) mapea 1:1 cada byte, asi que sirve para tratar el PDF
    # como texto sin romper los bytes binarios.
    $enc   = [System.Text.Encoding]::GetEncoding(28591)
    $bytes = [System.IO.File]::ReadAllBytes($Ruta)
    $texto = $enc.GetString($bytes)
    $todas = [regex]::Matches($texto, $PatronFecha)

    # Chrome escribe dos: CreationDate y ModDate.
    if ($todas.Count -lt 1) {
        throw "$Ruta`: no se encontro ninguna fecha para normalizar. El PDF no seria determinista."
    }

    $nuevo = $texto
    foreach ($m in ($todas | Sort-Object Index -Descending)) {
        if ($m.Value.Length -ne $FechaFija.Length) {
            throw "$Ruta`: la fecha '$($m.Value)' mide $($m.Value.Length) y la fija mide $($FechaFija.Length). Igualar longitudes antes de continuar: un reemplazo de otra longitud desplaza los offsets del PDF."
        }
        $nuevo = $nuevo.Substring(0, $m.Index) + $FechaFija + $nuevo.Substring($m.Index + $m.Value.Length)
    }
    [System.IO.File]::WriteAllBytes($Ruta, $enc.GetBytes($nuevo))
}

function New-CssTemporal {
    param(
        [Parameter(Mandatory)][int]$Columnas,
        [Parameter(Mandatory)][string]$Nombre,
        [Parameter(Mandatory)][string]$Destino
    )

    $css = [System.IO.File]::ReadAllText($CSS, [System.Text.Encoding]::UTF8)

    # Los parches van por regex y verifican que hayan entrado (contrato con la
    # plantilla, ver .DESCRIPTION).
    $reColumnas = 'column-count:\s*\d+;'
    $n = ([regex]::Matches($css, $reColumnas)).Count
    if ($n -ne 1) {
        throw "En la plantilla se esperaba 1 declaracion 'column-count', se encontraron $n. Revisar la plantilla antes de parchearla."
    }
    $css = [regex]::Replace($css, $reColumnas, "column-count: $Columnas;")

    # El margen se busca DENTRO de @page: hay otros 'margin' en la hoja de
    # estilos (los de cada titulo) que no se tienen que tocar.
    $reMargen = '(?s)(@page\s*\{.*?)margin:\s*[^;]+;'
    $n = ([regex]::Matches($css, $reMargen)).Count
    if ($n -ne 1) {
        throw "En la plantilla se esperaba 1 margen dentro de @page, se encontraron $n. Revisar la plantilla antes de parchearla."
    }
    # El @page base toma la configuracion de la hoja impar (lomo a la izquierda):
    # es la que vale si algun motor no resolviera los bloques :left/:right. En
    # modo simetrico queda el margen de los dos lados igual.
    $margen = if ($simetrico) {
        "$MargenVertical $MargenLateral"
    }
    else {
        "$MargenVertical $MargenExterno $MargenVertical $MargenLomo"
    }
    $css = [regex]::Replace($css, $reMargen, "`${1}margin: $margen;")

    # Encabezado con el nombre del documento. Lo inyecta el script y no la
    # plantilla porque el CSS no puede leer el nombre del archivo: Blink no
    # implementa string-set/string() (medido, el box queda vacio). El nombre es
    # el mismo que va en --variable=pagetitle de pandoc.
    $reEncabezado = '\/\*\[\[ENCABEZADO\]\]\*\/'
    $n = ([regex]::Matches($css, $reEncabezado)).Count
    if ($n -ne 1) {
        throw "En la plantilla se esperaba 1 marca de encabezado, se encontraron $n. Revisar la plantilla antes de parchearla."
    }
    $encabezado = ''
    if (-not $SinEncabezado) {
        $texto = $Nombre -replace '"', '\"'
        $encabezado = "@top-center {`n    content: `"$texto`";`n    font-size: 8pt;`n    color: #555555;`n    font-family: Arial, Helvetica, sans-serif;`n  }"
    }
    $css = [regex]::Replace($css, $reEncabezado, $encabezado)

    # Margenes espejados: se escriben salvo que se pida el modo simetrico.
    $reEspejado = '\/\*\[\[ESPEJADO\]\]\*\/'
    $n = ([regex]::Matches($css, $reEspejado)).Count
    if ($n -ne 1) {
        throw "En la plantilla se esperaba 1 marca de margenes espejados, se encontraron $n. Revisar la plantilla antes de parchearla."
    }

    # Fecha ISO-8601 (solo la fecha) en la esquina del lomo del pie. La inyecta
    # el script y no la plantilla porque el CSS no puede leer el reloj, igual
    # que pasa con el nombre del documento en el encabezado. Dos corridas el
    # MISMO dia producen los mismos bytes: el determinismo se conserva, la fecha
    # es parte de la entrada, igual que el nombre del archivo.
    $fechaPie = Get-Date -Format 'yyyy-MM-dd'
    $estiloFecha = "`n    font-size: 8pt;`n    color: #555555;`n    font-family: Arial, Helvetica, sans-serif;`n  }"
    $bloqueFechaImpar = "@bottom-left {`n    content: `"$fechaPie`";$estiloFecha"
    $bloqueFechaPar   = "@bottom-right {`n    content: `"$fechaPie`";$estiloFecha"

    $espejado = ''
    if (-not $simetrico) {
        $espejado = @"
/* Hoja impar (la 1 lo es): lomo a la izquierda. Hoja par: lomo a la derecha.
   La fecha va en la esquina del lomo: abajo a la izquierda en la impar y
   abajo a la derecha en la par. */
@page :right {
  margin: $MargenVertical $MargenExterno $MargenVertical $MargenLomo;
  $bloqueFechaImpar
}
@page :left  {
  margin: $MargenVertical $MargenLomo $MargenVertical $MargenExterno;
  $bloqueFechaPar
}
"@
    }
    else {
        # Modo simetrico (impresion de un solo lado): no hay lomo, la fecha va
        # abajo a la izquierda como en una hoja impar. Es un @page adicional y
        # Blink lo fusiona con el @page base de la plantilla.
        $espejado = @"
/* Modo simetrico: sin espejado. La fecha queda abajo a la izquierda. */
@page {
  $bloqueFechaImpar
}
"@
    }
    $css = [regex]::Replace($css, $reEspejado, $espejado)

    # Con una sola columna, llenar en orden evita el hueco que deja el balanceo.
    if ($Columnas -eq 1) {
        $css = [regex]::Replace($css, 'column-gap:\s*[^;]+;', "column-gap: 0.9cm;`n  column-fill: auto;")
    }

    # Sin BOM: el CSS se enlaza tal cual y un BOM al principio de un archivo CSS
    # no aporta nada.
    [System.IO.File]::WriteAllText($Destino, $css, (New-Object System.Text.UTF8Encoding($false)))
}

function Invoke-GenerarPdf {
    param(
        [Parameter(Mandatory)][string]$Nombre,
        [Parameter(Mandatory)][string]$Salida
    )

    if (-not (Test-Path (Join-Path $Destino "$Nombre.md"))) { throw "No existe $Nombre.md en $Destino" }

    # El HTML y el CSS temporales van en la carpeta de trabajo para que el CSS
    # relativo resuelva (Chrome resuelve el href contra la ruta del HTML). Con
    # ruta ABSOLUTA se escriben y con nombre relativo se enlazan: ver la trampa
    # del CWD en la resolucion de $Destino.
    $tmpCssNombre  = '_gen.tmp.css'
    $tmpHtmlNombre = '_gen.tmp.html'
    $tmpCss  = Join-Path $Destino $tmpCssNombre
    $tmpHtml = Join-Path $Destino $tmpHtmlNombre
    try {
        New-CssTemporal -Columnas $Columnas -Nombre $Nombre -Destino $tmpCss

        pandoc (Join-Path $Destino "$Nombre.md") `
            --from=markdown+task_lists+hard_line_breaks `
            --standalone `
            --css=$tmpCssNombre `
            --variable="pagetitle:$Nombre" `
            -o $tmpHtml

        $url  = 'file:///' + ((Resolve-Path -LiteralPath $tmpHtml).Path -replace '\\', '/')
        $dest = Join-Path $Destino $Salida

        & $Chrome --headless=new --disable-gpu --no-pdf-header-footer `
            --user-data-dir="$Perfil" --print-to-pdf="$dest" $url | Out-Null

        if (-not (Test-Path -LiteralPath $dest)) {
            throw "Chrome no genero $Salida (revisar la salida de Chrome)."
        }

        Set-FechaDeterministica -Ruta $dest
    }
    finally {
        if (Test-Path $tmpHtml) { Remove-Item $tmpHtml -Force }
        if (Test-Path $tmpCss)  { Remove-Item $tmpCss  -Force }
    }
}

function Get-HashArchivo {
    param([Parameter(Mandatory)][string]$Ruta)
    (Get-FileHash -LiteralPath $Ruta -Algorithm SHA256).Hash
}

$fallos = 0
foreach ($nombre in $Docs) {
    Write-Host "==> $nombre.pdf  ($Columnas columna(s))" -ForegroundColor Cyan
    Invoke-GenerarPdf -Nombre $nombre -Salida "$nombre.pdf"
    $hash = Get-HashArchivo -Ruta (Join-Path $Destino "$nombre.pdf")

    if ($Verify) {
        $tmp   = "_$nombre.verify.pdf"
        Invoke-GenerarPdf -Nombre $nombre -Salida $tmp
        $hash2 = Get-HashArchivo -Ruta (Join-Path $Destino $tmp)
        Remove-Item (Join-Path $Destino $tmp) -Force

        if ($hash -eq $hash2) {
            Write-Host "    determinista OK  ($($hash.Substring(0,16)))" -ForegroundColor Green
        }
        else {
            Write-Host "    NO DETERMINISTA  $($hash.Substring(0,16)) != $($hash2.Substring(0,16))" -ForegroundColor Red
            $fallos++
        }
    }
    else {
        Write-Host "    sha256 $($hash.Substring(0,16))" -ForegroundColor DarkGray
    }
}

if ($fallos -gt 0) {
    Write-Host ""
    Write-Host "Atencion: $fallos PDF(s) no deterministas." -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Listo. $(@($Docs).Count) PDF(s) generados." -ForegroundColor Green
