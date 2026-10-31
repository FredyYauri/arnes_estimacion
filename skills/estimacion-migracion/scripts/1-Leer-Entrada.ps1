<#
.SYNOPSIS
  Paso 2-3 del procedimiento: lee la entrada (solo lectura), la vuelca a _trabajo/entrada.json
  y ejecuta los controles de references/entradas.md.
.OUTPUTS
  _trabajo/entrada.json, _trabajo/faltantes.json
  Código de salida 0 = entrada completa; 2 = hay faltantes (el agente se detiene y los pide).
.EXAMPLE
  .\1-Leer-Entrada.ps1 -Entrada 'Estimaciones\MigraMacro\entrada_MigraMacro.xlsx'
#>
param(
    [Parameter(Mandatory = $true)][string]$Entrada,
    [string]$Trabajo
)
. (Join-Path $PSScriptRoot 'Comun.ps1')

$Entrada = (Resolve-Path -LiteralPath $Entrada).Path
if (-not $Trabajo) { $Trabajo = Join-Path (Split-Path -Parent $Entrada) '_trabajo' }

# ---------- Lectura ----------
$xl = Open-Excel
try {
    $wb = Invoke-Com { $xl.Workbooks.Open($Entrada, 0, $true) }
    $hojas = @{}
    foreach ($ws in $wb.Worksheets) { $hojas[$ws.Name] = $ws }

    function Read-Pares($ws) {
        $h = [ordered]@{}
        foreach ($f in (Read-SheetTable $ws 3)) {
            $vals = @($f.Values)
            $campo = [string]$vals[1]; $valor = $vals[2]
            if ($campo) { $h[$campo.Trim()] = $valor }
        }
        return $h
    }

    $entradaJson = [ordered]@{
        archivo      = $Entrada
        leido        = (Get-Date).ToString('yyyy-MM-dd HH:mm')
        proyecto     = Read-Pares $hojas['Proyecto']
        parametros   = Read-Pares $hojas['Parametros']
        inventario   = Read-SheetTable $hojas['Inventario'] 3
        recursos     = Read-SheetTable $hojas['Recursos'] 3
        calendario   = Read-SheetTable $hojas['Calendario'] 3
        restricciones = Read-SheetTable $hojas['Restricciones'] 3
        riesgos      = Read-SheetTable $hojas['Riesgos'] 3
        catalogos    = [ordered]@{}
    }
    # Catalogos: encabezados en la fila 1, una lista por columna.
    foreach ($f in (Read-SheetTable $hojas['Catalogos'] 1)) {
        foreach ($k in $f.Keys) {
            if ($k -eq '_fila') { continue }
            $v = $f[$k]
            if ($null -eq $v -or ([string]$v).Trim() -eq '') { continue }
            if (-not $entradaJson.catalogos.Contains($k)) { $entradaJson.catalogos[$k] = New-Object System.Collections.Generic.List[object] }
            $entradaJson.catalogos[$k].Add($v)
        }
    }
    # Capacidad planificable por rol (columnas "Roles" y "Capacidad planificable (h/día)").
    $capRol = [ordered]@{}
    $roles = @(Get-Val $entradaJson.catalogos 'Roles' @())
    $caps = @(Get-Val $entradaJson.catalogos 'Capacidad planificable (h/día)' @())
    for ($i = 0; $i -lt [Math]::Min($roles.Count, $caps.Count); $i++) { $capRol[[string]$roles[$i]] = To-Num $caps[$i] }
    $entradaJson['capacidadRol'] = $capRol
    $wb.Close($false)
}
finally { Close-Excel $xl }

# Fechas OLE -> ISO
foreach ($c in @('Fecha tentativa de inicio', 'Fecha objetivo (comprometida o deseada)')) {
    if ($entradaJson.proyecto.Contains($c)) { $entradaJson.proyecto[$c] = ConvertTo-IsoDate $entradaJson.proyecto[$c] }
}
foreach ($r in $entradaJson.recursos) { foreach ($c in @('Disponible desde', 'Disponible hasta')) { if ($r.Contains($c)) { $r[$c] = ConvertTo-IsoDate $r[$c] } } }
foreach ($r in $entradaJson.calendario) { $r['Fecha'] = ConvertTo-IsoDate $r['Fecha'] }
foreach ($r in $entradaJson.restricciones) { if ($r.Contains('Fecha')) { $r['Fecha'] = ConvertTo-IsoDate $r['Fecha'] } }

# ---------- Controles ----------
$faltantes = New-Object System.Collections.Generic.List[object]
$avisos = New-Object System.Collections.Generic.List[string]
function Add-Faltante($Hoja, $Fila, $Campo, $Problema) {
    $faltantes.Add([ordered]@{ hoja = $Hoja; fila = [string]$Fila; campo = $Campo; problema = $Problema })
}
function Test-Lista($Hoja, $Fila, $Campo, $Valor, $Catalogo) {
    if ($null -eq $Valor -or ([string]$Valor).Trim() -eq '') { return }
    $lista = @(Get-Val $entradaJson.catalogos $Catalogo @()) | ForEach-Object { [string]$_ }
    if ($lista.Count -gt 0 -and -not ($lista -contains [string]$Valor)) {
        Add-Faltante $Hoja $Fila $Campo "Valor '$Valor' no está en Catalogos ($Catalogo)"
    }
}

$p = $entradaJson.proyecto
$bloqProyecto = @('Nombre del proyecto', 'Tipo de trabajo', 'Objetivo', 'Alcance (resumen)', 'Ambientes disponibles',
    'Fecha tentativa de inicio', 'Claridad de requerimientos', 'Datos sensibles o exposición a internet')
$polProyecto = @('La estimación llega hasta', 'Jornada (h/día)', 'Días laborables', 'Uso de asistente IA en el desarrollo',
    'Modelo de entrega', 'Agrupación de certificación y despliegue', 'Certificación a cargo de',
    'Duración mínima de certificación (días hábiles)', 'Duración de hypercare (días hábiles)', 'Versión de la estimación')
foreach ($c in $bloqProyecto) { if ($null -eq (Get-Val $p $c)) { Add-Faltante 'Proyecto' '' $c 'Bloqueante vacío' } }
foreach ($c in $polProyecto) { if ($null -eq (Get-Val $p $c)) { Add-Faltante 'Proyecto' '' $c 'Política vacía (se volvió bloqueante)' } }
$tarifas = Is-Si (Get-Val $p 'Tarifas por rol disponibles (costeo)' 'No')
if ($tarifas -and $null -eq (Get-Val $p 'Moneda')) { Add-Faltante 'Proyecto' '' 'Moneda' 'Obligatoria si hay tarifas' }
Test-Lista 'Proyecto' '' 'Tipo de trabajo' (Get-Val $p 'Tipo de trabajo') 'Tipo de trabajo'
Test-Lista 'Proyecto' '' 'La estimación llega hasta' (Get-Val $p 'La estimación llega hasta') 'La estimación llega hasta'
Test-Lista 'Proyecto' '' 'Modelo de entrega' (Get-Val $p 'Modelo de entrega') 'Modelo de entrega'
Test-Lista 'Proyecto' '' 'Agrupación de certificación y despliegue' (Get-Val $p 'Agrupación de certificación y despliegue') 'Agrupación cert./despliegue'
Test-Lista 'Proyecto' '' 'Certificación a cargo de' (Get-Val $p 'Certificación a cargo de') 'Certificación a cargo de'
if ((Get-Val $p 'Modelo de entrega') -eq 'Oleadas') {
    $no = Get-Val $p 'Número de olas'
    if ($null -eq $no -or (To-Num $no 0) -lt 2) { $avisos.Add('Modelo de entrega = Oleadas sin "Número de olas" (≥ 2): tomarlo del pedido o confirmar con el usuario; si no se indica, es una sola entrega.') }
}
$fi = Get-Val $p 'Fecha tentativa de inicio'; $fo = Get-Val $p 'Fecha objetivo (comprometida o deseada)'
if ($fi -and $fo -and ($fo -lt $fi)) { Add-Faltante 'Proyecto' '' 'Fecha objetivo' 'Es anterior a la fecha de inicio' }

foreach ($k in $entradaJson.parametros.Keys) {
    if ($null -eq (Get-Val $entradaJson.parametros $k)) { Add-Faltante 'Parametros' '' $k 'Parámetro vacío' }
}

$ids = @{}
$bloqInv = @('ID', 'Nombre', 'Aplicativo origen', 'Tecnología origen', 'Aplicativo destino', 'Destino existente o nuevo', 'Plataforma destino', 'Estrategia')
$bloqInvEstrella = @('LOC', 'Pantallas / formularios', 'Reportes / salidas', 'Integraciones (n)', 'Tablas / hojas / fuentes de datos (n)',
    'Migración de datos', 'Reglas de negocio', 'Riesgo técnico', 'Dependencia externa', 'Reutilización', 'Requiere permisos y auditoría')
if ($entradaJson.inventario.Count -eq 0) { Add-Faltante 'Inventario' '' '' 'Sin componentes' }
foreach ($r in $entradaJson.inventario) {
    $fila = $r['_fila']; $id = [string](Get-Val $r 'ID' '')
    foreach ($c in $bloqInv) { if ($null -eq (Get-Val $r $c)) { Add-Faltante 'Inventario' "$fila / $id" $c 'Bloqueante vacío' } }
    $est = [string](Get-Val $r 'Estrategia' '')
    if ($est -notin @('Retirar', 'Mantener')) {
        foreach ($c in $bloqInvEstrella) { if ($null -eq (Get-Val $r $c)) { Add-Faltante 'Inventario' "$fila / $id" $c 'Bloqueante vacío (escriba 0 si no aplica)' } }
        if ((To-Num (Get-Val $r 'Integraciones (n)' 0)) -gt 0 -and $null -eq (Get-Val $r 'Detalle de integraciones')) {
            Add-Faltante 'Inventario' "$fila / $id" 'Detalle de integraciones' 'Obligatorio si Integraciones > 0'
        }
    }
    if ($id) { if ($ids.ContainsKey($id)) { Add-Faltante 'Inventario' "$fila / $id" 'ID' 'ID duplicado' } else { $ids[$id] = $true } }
    Test-Lista 'Inventario' "$fila / $id" 'Plataforma destino' (Get-Val $r 'Plataforma destino') 'Plataformas destino'
    Test-Lista 'Inventario' "$fila / $id" 'Estrategia' $est 'Estrategias'
    Test-Lista 'Inventario' "$fila / $id" 'Destino existente o nuevo' (Get-Val $r 'Destino existente o nuevo') 'Destino'
    Test-Lista 'Inventario' "$fila / $id" 'Migración de datos' (Get-Val $r 'Migración de datos') 'Migración de datos'
    Test-Lista 'Inventario' "$fila / $id" 'Dependencia externa' (Get-Val $r 'Dependencia externa') 'Dependencia externa'
}

$nRec = 0
foreach ($r in $entradaJson.recursos) {
    $fila = $r['_fila']; $nom = [string](Get-Val $r 'Recurso (nombre o alias)' '')
    $nRec++
    foreach ($c in @('Recurso (nombre o alias)', 'Rol', 'Dedicación %')) { if ($null -eq (Get-Val $r $c)) { Add-Faltante 'Recursos' "$fila / $nom" $c 'Bloqueante vacío' } }
    $rol = [string](Get-Val $r 'Rol' '')
    Test-Lista 'Recursos' "$fila / $nom" 'Rol' $rol 'Roles'
    if (($rol -match 'Developer|Data Engineer') -and $null -eq (Get-Val $r 'Plataformas / skills')) { Add-Faltante 'Recursos' "$fila / $nom" 'Plataformas / skills' 'Obligatorio para roles de desarrollo' }
    if ($tarifas -and $null -eq (Get-Val $r 'Tarifa por hora (opcional)')) { Add-Faltante 'Recursos' "$fila / $nom" 'Tarifa por hora' 'Obligatoria si hay tarifas' }
    $d1 = Get-Val $r 'Disponible desde'; $d2 = Get-Val $r 'Disponible hasta'
    if ($d1 -and $d2 -and $d2 -lt $d1) { Add-Faltante 'Recursos' "$fila / $nom" 'Disponible hasta' 'Anterior a Disponible desde' }
}

# Cobertura de recursos (el agente pregunta si esperar o dimensionar; no es un dato que se suponga).
$cobertura = New-Object System.Collections.Generic.List[string]
$plataformas = @($entradaJson.inventario | Where-Object { (Get-Val $_ 'Estrategia' '') -notin @('Retirar', 'Mantener') } | ForEach-Object { [string](Get-Val $_ 'Plataforma destino' '') } | Where-Object { $_ } | Sort-Object -Unique)
foreach ($pl in $plataformas) {
    $ok = @($entradaJson.recursos | Where-Object { ([string](Get-Val $_ 'Plataformas / skills' '')).ToLower().Contains($pl.ToLower()) }).Count -gt 0
    if (-not $ok) { $cobertura.Add("Ningún recurso cubre la plataforma '$pl'") }
}
$cert = [string](Get-Val $p 'Certificación a cargo de' '')
if ($cert -in @('QA proveedor', 'Mixta') -and @($entradaJson.recursos | Where-Object { (Get-Val $_ 'Rol' '') -match 'Calidad|Tester' }).Count -eq 0) { $cobertura.Add('Falta un rol de calidad (Certificación a cargo de = ' + $cert + ')') }
if ((Is-Si (Get-Val $p 'Datos sensibles o exposición a internet' 'No')) -and @($entradaJson.recursos | Where-Object { (Get-Val $_ 'Rol' '') -match 'hacking' }).Count -eq 0) { $cobertura.Add('Falta un rol de seguridad (datos sensibles = Sí)') }
if ($nRec -eq 0) { $cobertura.Add('La hoja Recursos está vacía') }
if ($entradaJson.calendario.Count -eq 0) { $avisos.Add('Calendario vacío: confirmar con el usuario "sin feriados ni ausencias en el periodo" antes de programar.') }

Write-JsonFile (Join-Path $Trabajo 'entrada.json') $entradaJson
Write-JsonFile (Join-Path $Trabajo 'faltantes.json') ([ordered]@{ faltantes = $faltantes; coberturaRecursos = $cobertura; avisos = $avisos })

Write-Host "Entrada leída: $($entradaJson.inventario.Count) componentes, $nRec recursos."
if ($faltantes.Count -gt 0) {
    Write-Host "FALTANTES ($($faltantes.Count)) — Hoja · Fila/ID · Campo · Problema:"
    foreach ($f in $faltantes) { Write-Host (" - {0} · {1} · {2} · {3}" -f $f.hoja, $f.fila, $f.campo, $f.problema) }
}
foreach ($c in $cobertura) { Write-Host " - Recursos: $c" }
foreach ($a in $avisos) { Write-Host " - Aviso: $a" }
if ($faltantes.Count -gt 0) { exit 2 }
exit 0
