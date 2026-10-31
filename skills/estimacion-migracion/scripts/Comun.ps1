# Funciones compartidas por los scripts del arnés (PowerShell 5.1, Windows, Excel COM).
# Se carga con:  . (Join-Path $PSScriptRoot 'Comun.ps1')


$ErrorActionPreference = 'Stop'
$Inv = [Globalization.CultureInfo]::InvariantCulture

# ---------- JSON (UTF-8, sin BOM) ----------

function Read-JsonFile([string]$Ruta) {
    if (-not (Test-Path -LiteralPath $Ruta)) { return $null }
    $txt = [IO.File]::ReadAllText($Ruta, [Text.Encoding]::UTF8)
    if ([string]::IsNullOrWhiteSpace($txt)) { return $null }
    return ConvertTo-Hash ($txt | ConvertFrom-Json)
}

function Write-JsonFile([string]$Ruta, $Objeto) {
    $dir = Split-Path -Parent $Ruta
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
    $json = ConvertTo-Json -InputObject $Objeto -Depth 12
    [IO.File]::WriteAllText($Ruta, $json, (New-Object Text.UTF8Encoding $false))
}

# ConvertFrom-Json de PS 5.1 devuelve PSCustomObject; se pasa a hashtables ordenados y listas.
function ConvertTo-Hash($o) {
    if ($null -eq $o) { return $null }
    if ($o -is [System.Management.Automation.PSCustomObject]) {
        $h = [ordered]@{}
        foreach ($p in $o.PSObject.Properties) { $h[$p.Name] = ConvertTo-Hash $p.Value }
        return $h
    }
    if (($o -is [System.Collections.IEnumerable]) -and -not ($o -is [string])) {
        $l = New-Object System.Collections.Generic.List[object]
        foreach ($x in $o) { $l.Add((ConvertTo-Hash $x)) }
        return , $l
    }
    return $o
}

function Get-Val($Hash, [string]$Clave, $Defecto = $null) {
    if ($null -eq $Hash) { return $Defecto }
    if ($Hash.Contains($Clave)) {
        $v = $Hash[$Clave]
        if ($null -eq $v) { return $Defecto }
        if (($v -is [string]) -and [string]::IsNullOrWhiteSpace($v)) { return $Defecto }
        return $v
    }
    return $Defecto
}

function To-Num($v, [double]$Defecto = 0) {
    if ($null -eq $v) { return $Defecto }
    if ($v -is [double] -or $v -is [int] -or $v -is [long] -or $v -is [decimal]) { return [double]$v }
    $s = ([string]$v).Trim().Replace('%', '')
    if ($s -eq '') { return $Defecto }
    $d = 0.0
    if ([double]::TryParse($s.Replace(',', '.'), [Globalization.NumberStyles]::Float, $Inv, [ref]$d)) { return $d }
    return $Defecto
}

function Is-Si($v) { return (([string]$v).Trim().ToLower() -in @('sí', 'si', 's', 'yes', 'true', '1')) }

function Ceil-Int([double]$x) { return [int][Math]::Ceiling([Math]::Round($x, 6)) }

# Horas finales = ROUNDUP(ROUND(base x factor, 1), 0), mínimo $Minimo.
function Get-HorasFinales([double]$Base, [double]$Factor, [int]$Minimo = 1) {
    $h = [int][Math]::Ceiling([Math]::Round($Base * $Factor, 1))
    if ($h -lt $Minimo) { $h = $Minimo }
    return $h
}

# ---------- Fechas y días hábiles ----------

function ConvertTo-IsoDate($v) {
    if ($null -eq $v) { return $null }
    if ($v -is [double] -or $v -is [int]) { return [DateTime]::FromOADate([double]$v).ToString('yyyy-MM-dd') }
    if ($v -is [DateTime]) { return $v.ToString('yyyy-MM-dd') }
    $s = ([string]$v).Trim()
    if ($s -eq '') { return $null }
    $d = [DateTime]::MinValue
    foreach ($f in @('yyyy-MM-dd', 'dd/MM/yyyy', 'd/M/yyyy', 'dd-MM-yyyy')) {
        if ([DateTime]::TryParseExact($s, $f, $Inv, [Globalization.DateTimeStyles]::None, [ref]$d)) { return $d.ToString('yyyy-MM-dd') }
    }
    return $null
}

function Get-IsoDate([string]$Iso) { return [DateTime]::ParseExact($Iso, 'yyyy-MM-dd', $Inv) }

# Lista de fechas hábiles (ISO) desde $Inicio, saltando fines de semana y feriados.
function Get-DiasHabiles([string]$Inicio, [int]$Cantidad, [string[]]$Feriados, [bool]$Sabado = $false) {
    $fer = @{}
    foreach ($f in $Feriados) { if ($f) { $fer[$f] = $true } }
    $l = New-Object System.Collections.Generic.List[string]
    $d = Get-IsoDate $Inicio
    while ($l.Count -lt $Cantidad) {
        $dow = [int]$d.DayOfWeek
        $esLaborable = ($dow -ge 1 -and $dow -le 5) -or ($Sabado -and $dow -eq 6)
        $iso = $d.ToString('yyyy-MM-dd')
        if ($esLaborable -and -not $fer.ContainsKey($iso)) { $l.Add($iso) }
        $d = $d.AddDays(1)
    }
    return , $l
}

# ---------- Excel COM ----------

function Open-Excel {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.ScreenUpdating = $false
    return $xl
}

function Close-Excel($xl) {
    if ($null -eq $xl) { return }
    try { $xl.Quit() } catch { }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

# Reintenta una acción COM cuando Excel está ocupado (0x800AC472).
function Invoke-Com([scriptblock]$Accion) {
    for ($i = 0; $i -lt 10; $i++) {
        try { return & $Accion }
        catch {
            if ($_.Exception.HResult -eq -2146777998 -or $_.Exception.Message -match '800AC472') { Start-Sleep -Milliseconds 300 }
            else { throw }
        }
    }
    return & $Accion
}

# Lee una hoja como tabla: encabezados en $FilaEnc, datos desde la fila siguiente hasta la primera fila vacía.
function Read-SheetTable($ws, [int]$FilaEnc = 3) {
    $data = $ws.UsedRange.Value2
    $filas = New-Object System.Collections.Generic.List[object]
    if ($null -eq $data -or -not ($data -is [Array])) { return , $filas }
    $r0 = $ws.UsedRange.Row; $c0 = $ws.UsedRange.Column
    $nR = $data.GetLength(0); $nC = $data.GetLength(1)
    $ri = $FilaEnc - $r0 + 1
    if ($ri -lt 1 -or $ri -gt $nR) { return , $filas }
    $enc = @{}
    for ($c = 1; $c -le $nC; $c++) {
        $t = $data[$ri, $c]
        if ($null -ne $t -and ([string]$t).Trim() -ne '') { $enc[$c] = ([string]$t).Trim() }
    }
    for ($r = $ri + 1; $r -le $nR; $r++) {
        $fila = [ordered]@{ '_fila' = $r + $r0 - 1 }
        $vacia = $true
        foreach ($c in ($enc.Keys | Sort-Object)) {
            $v = $data[$r, $c]
            if ($null -ne $v -and ([string]$v).Trim() -ne '') { $vacia = $false }
            if ($v -is [string]) { $v = $v.Trim() }
            $fila[$enc[$c]] = $v
        }
        if (-not $vacia) { $filas.Add($fila) }
    }
    return , $filas
}

function ConvertTo-BGR([string]$Hex) {
    $h = $Hex.TrimStart('#')
    $r = [Convert]::ToInt32($h.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($h.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($h.Substring(4, 2), 16)
    return $r + 256 * $g + 65536 * $b
}

# Escribe un valor en una celda según las reglas de excel-com.md.
function Set-Cell($ws, [int]$Fila, [int]$Col, $Valor) {
    $c = $ws.Cells.Item($Fila, $Col)
    if ($null -eq $Valor) { return }
    if ($Valor -is [string]) {
        if ($Valor.StartsWith('=')) { $c.Formula = $Valor } else { $c.Value2 = [string]$Valor }
    }
    elseif ($Valor -is [bool]) { $c.Value2 = [string]$Valor }
    else { $c.Formula = ([double]$Valor).ToString($Inv) }
}

function Set-Row($ws, [int]$Fila, [int]$ColIni, [object[]]$Valores) {
    for ($i = 0; $i -lt $Valores.Count; $i++) { Set-Cell $ws $Fila ($ColIni + $i) $Valores[$i] }
}

function Get-ColLetter([int]$n) {
    $s = ''
    while ($n -gt 0) { $m = ($n - 1) % 26; $s = [char](65 + $m) + $s; $n = [int][Math]::Floor(($n - $m) / 26) }
    return $s
}

function Get-FileHashSafe([string]$Ruta) {
    if (-not (Test-Path -LiteralPath $Ruta)) { return $null }
    return (Get-FileHash -LiteralPath $Ruta -Algorithm SHA256).Hash
}

# Nombres de las hojas estándar del Excel de salida (orden obligatorio).
function Get-HojasEstandar([string]$TipoTrabajo) {
    return @('Resumen ejecutivo', 'Estrategia de oleadas', 'Matriz de entregables', "Planning - $TipoTrabajo",
        'Capacidad por recurso', 'Equipo', 'Detalle de horas', 'Supuestos', 'Riesgos', 'Validaciones')
}

# Zonas del Excel de salida que el usuario puede editar a mano y que se conservan al regenerar.
# Cualquier otra celda se recalcula; las hojas propias del usuario (no estándar) se copian tal cual.
function Get-TablasEditables {
    return [ordered]@{
        'Estrategia de oleadas' = @{ filaEnc = 3; clave = 'Código'; editables = @('Motivo de la secuencia', 'Riesgos / dependencias', 'Información faltante o por validar') }
        'Matriz de entregables' = @{ filaEnc = 3; clave = 'Entregable'; editables = @('Objetivo del entregable', 'Criterio de salida', 'Riesgos o dependencias') }
        'Equipo'                = @{ filaEnc = 3; clave = 'Recurso'; editables = @('Rol en el proyecto', 'Racional de la asignación') }
        'Detalle de horas'      = @{ filaEnc = 3; clave = 'ID origen'; editables = @('Observación') }
        'Validaciones'          = @{ filaEnc = 3; clave = '#'; editables = @('Detalle / justificación') }
    }
}
# Hojas cuyas filas completas (por ID) puede agregar o modificar el usuario.
function Get-HojasFilasManuales { return @('Supuestos', 'Riesgos') }

# Lee una tabla desde $FilaEnc hasta la primera fila cuya columna clave esté vacía.
function Read-TablaHasta($ws, [int]$FilaEnc, [string]$ColClave) {
    $filas = New-Object System.Collections.Generic.List[object]
    $data = $ws.UsedRange.Value2
    if ($null -eq $data -or -not ($data -is [Array])) { return , $filas }
    $r0 = $ws.UsedRange.Row; $c0 = $ws.UsedRange.Column
    $ri = $FilaEnc - $r0 + 1; if ($ri -lt 1) { return , $filas }
    $enc = [ordered]@{}; $ck = 0
    for ($c = 1; $c -le $data.GetLength(1); $c++) { $t = $data[$ri, $c]; if ($null -ne $t -and ([string]$t).Trim() -ne '') { $enc[[string]$c] = ([string]$t).Trim(); if (([string]$t).Trim() -eq $ColClave) { $ck = $c } } }
    if ($ck -eq 0) { return , $filas }
    for ($r = $ri + 1; $r -le $data.GetLength(0); $r++) {
        $k = $data[$r, $ck]; if ($null -eq $k -or ([string]$k).Trim() -eq '') { break }
        $f = [ordered]@{ '_fila' = $r + $r0 - 1 }
        foreach ($c in $enc.Keys) { $v = $data[$r, [int]$c]; $f[$enc[$c]] = $(if ($null -eq $v) { '' } else { ([string]$v).Trim() }) }
        $filas.Add($f)
    }
    return , $filas
}

# Lee las filas de tarea de la hoja Planning: encabezado = fila con "ID" en A y "Recurso asignado" en I.
function Read-PlanningTareas($ws) {
    $res = [ordered]@{ tareas = [ordered]@{}; filasSinId = New-Object System.Collections.Generic.List[int] }
    $data = $ws.UsedRange.Value2
    if ($null -eq $data -or -not ($data -is [Array])) { return $res }
    $r0 = $ws.UsedRange.Row; $c0 = $ws.UsedRange.Column
    $nR = $data.GetLength(0); $nC = $data.GetLength(1)
    function V([int]$fila, [int]$col) { $i = $fila - $r0 + 1; $j = $col - $c0 + 1; if ($i -lt 1 -or $j -lt 1 -or $i -gt $nR -or $j -gt $nC) { return $null }; return $data[$i, $j] }
    $hr = 0
    for ($f = $r0; $f -lt $r0 + $nR; $f++) { if ([string](V $f 1) -eq 'ID' -and [string](V $f 9) -eq 'Recurso asignado') { $hr = $f; break } }
    if ($hr -eq 0) { return $res }
    $res['filaEnc'] = $hr
    $colDia = @{}
    for ($c = 11; $c -lt $c0 + $nC; $c++) { $v = V $hr $c; if ($null -ne $v -and ([string]$v) -match '^\d+$') { $colDia[$c] = [int]$v } }
    for ($f = $hr + 1; $f -lt $r0 + $nR; $f++) {
        $id = [string](V $f 1)
        $hayAlgo = $false; for ($c = 1; $c -le 10; $c++) { if ($null -ne (V $f $c) -and ([string](V $f $c)).Trim() -ne '') { $hayAlgo = $true; break } }
        if ($id -notmatch '^T\d+$') { if ($hayAlgo) { $res.filasSinId.Add($f) }; continue }
        $dias = [ordered]@{}
        foreach ($c in ($colDia.Keys | Sort-Object)) { $v = V $f $c; if ($null -ne $v -and ([string]$v).Trim() -ne '' -and (To-Num $v) -ne 0) { $dias[[string]$colDia[$c]] = To-Num $v } }
        $res.tareas[$id] = [ordered]@{ fila = $f; recurso = [string](V $f 9); tarea = [string](V $f 8); dias = $dias }
    }
    return $res
}
