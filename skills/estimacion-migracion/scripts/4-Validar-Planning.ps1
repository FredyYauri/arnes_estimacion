<#
.SYNOPSIS
  Paso 10 del procedimiento: pre-valida planning.json con los controles de references/validaciones.md.
.OUTPUTS
  _trabajo/validaciones.json y _trabajo/validaciones.txt. Código de salida 4 si algún control queda en REVISAR.
.EXAMPLE
  .\4-Validar-Planning.ps1 -Trabajo 'Estimaciones\MigraMacro\_trabajo'
#>
param([Parameter(Mandatory = $true)][string]$Trabajo)
. (Join-Path $PSScriptRoot 'Comun.ps1')

$PLAN = Read-JsonFile (Join-Path $Trabajo 'planning.json')
$WBS = Read-JsonFile (Join-Path $Trabajo 'wbs.json')
if ($null -eq $PLAN) { throw 'Falta planning.json. Ejecute 3-Programar.ps1.' }
$TAREAS = @($PLAN.tareas); $NDIAS = [int]$PLAN.ultimoDia; $jornada = [int]$PLAN.jornada
$porId = @{}; foreach ($t in $TAREAS) { $porId[$t.id] = $t }
$capDe = @{}; foreach ($r in $PLAN.recursos) { $capDe[$r.nombre] = $r }
$res = New-Object System.Collections.Generic.List[object]
function Add-Control([int]$n, [string]$control, $valor, [string]$esperado, [bool]$ok, [string]$detalle = '') {
    $res.Add([ordered]@{ n = $n; control = $control; valor = $valor; esperado = $esperado; estado = $(if ($ok) { 'OK' } else { 'REVISAR' }); detalle = $detalle })
}
function Get-Dias($t) { $l = @(); foreach ($k in $t.dias.Keys) { if ([int]$t.dias[$k] -ne 0) { $l += [int]$k } }; return , @($l | Sort-Object) }
function Cap($rec, [int]$d) { $r = $capDe[$rec]; if ($null -eq $r -or $d -lt 1 -or $d -gt $r.capDia.Count) { return 0 }; return [int]$r.capDia[$d - 1] }

# Carga por recurso y día
$carga = @{}
foreach ($r in $PLAN.recursos) { $carga[$r.nombre] = New-Object int[] ($NDIAS + 2) }
foreach ($t in $TAREAS) { foreach ($k in $t.dias.Keys) { $d = [int]$k; if ($d -le $NDIAS -and $carga.ContainsKey($t.recurso)) { $carga[$t.recurso][$d] += [int](To-Num $t.dias[$k]) } } }

# 1-2 Cobertura de componentes
$comps = @($WBS.componentes)
$faltan = @(); foreach ($c in $comps) { if ($c.tieneTareas -and @($TAREAS | Where-Object { $_.componente -eq $c.id }).Count -eq 0) { $faltan += $c.id } }
Add-Control 1 'Todas las funcionalidades están incluidas' ($comps.Count - $faltan.Count) "$($comps.Count)" ($faltan.Count -eq 0) ($faltan -join ', ')
$sinFase = @()
foreach ($c in $comps) {
    if (-not $c.tieneTareas -or $c.estrategia -eq 'Retirar') { continue }
    foreach ($f in @('02', '03', '04', '05')) { if (@($TAREAS | Where-Object { $_.componente -eq $c.id -and $_.fase -like "$f*" }).Count -eq 0) { $sinFase += "$($c.id) sin $f" } }
}
Add-Control 2 'Cada funcionalidad tiene análisis, diseño, construcción y pruebas unitarias' $sinFase.Count '0' ($sinFase.Count -eq 0) ($sinFase -join '; ')

# 3-6 Datos de cada tarea
$x = @($TAREAS | Where-Object { -not $_.recurso }).Count
Add-Control 3 'Cada tarea tiene responsable' $x '0' ($x -eq 0)
$x = @($TAREAS | Where-Object { [int]$_.horas -le 0 -and $_.tipo -ne 'relleno' }).Count
Add-Control 4 'Cada tarea tiene horas > 0' $x '0' ($x -eq 0)
$malas = @()
foreach ($t in $TAREAS) {
    $s = 0; foreach ($k in $t.dias.Keys) { $v = To-Num $t.dias[$k]; if ($v -lt 0 -or $v -ne [Math]::Floor($v)) { $malas += "$($t.id) día $k" }; $s += $v }
    if ($s -ne [int]$t.horas) { $malas += "$($t.id) horas $($t.horas) ≠ Σ días $s" }
}
Add-Control 5 'No hay horas negativas, decimales ni inconsistentes' $malas.Count '0' ($malas.Count -eq 0) (($malas | Select-Object -First 10) -join '; ')
$x = @($TAREAS | Where-Object { -not $_.ola -or -not $_.aplicativoOrigen -or -not $_.aplicativo }).Count
Add-Control 6 'Toda tarea tiene ola, aplicativo origen y aplicativo/entregable destino' $x '0' ($x -eq 0)

# 7-8 Totales y origen
$tot = 0; $totDias = 0; foreach ($t in $TAREAS) { $tot += [int]$t.horas; $totDias += [int]$t.horasProgramadas }
Add-Control 7 'Total Planning = Total Detalle de horas' ($tot - $totDias) '0' ($tot -eq $totDias)
$asum = 0; $cont = 0; foreach ($t in $TAREAS) { if ($t.origen -eq 'Asumida') { $asum += $t.horas }; if ($t.origen -eq 'Contingencia') { $cont += $t.horas } }
Add-Control 8 'Horas agregadas identificadas (Asumida + Contingencia)' ($asum + $cont) 'Marcadas con prefijo' $true "Asumida $asum h · Contingencia $cont h"

# 9-10 Capacidad
$sobre = @(); $maxDia = 0
foreach ($n in $carga.Keys) { for ($d = 1; $d -le $NDIAS; $d++) { $h = $carga[$n][$d]; if ($h -gt $maxDia) { $maxDia = $h }; if ($h -gt (Cap $n $d)) { $sobre += "$n día $d ($h > $(Cap $n $d))" } } }
Add-Control 9 'No hay recursos sobre su capacidad diaria' $sobre.Count '0' ($sobre.Count -eq 0) (($sobre | Select-Object -First 10) -join '; ')
Add-Control 10 'Máximo de horas de un recurso en un día' $maxDia "≤ $jornada" ($maxDia -le $jornada)

# 11-12 Transversales y preparación
$x = @($TAREAS | Where-Object { $_.ola -eq 'Transversal' }).Count
Add-Control 11 'Actividades transversales identificadas' $x '≥ 8' ($x -ge 8)
$p1 = @($TAREAS | Where-Object { $_.fase -like '01.1*' }).Count; $p2 = @($TAREAS | Where-Object { $_.fase -like '01.2*' }).Count
Add-Control 12 'Preparación presente (01.1 y 01.2)' ($p1 * $p2) '> 0' (($p1 * $p2) -gt 0)

# 13 Certificación mínima por bloque
$malB = @(); $nB = 0
foreach ($b in $PLAN.bloques) {
    $ce = @($TAREAS | Where-Object { $_.bloque -eq $b.nombre -and $_.codigo -in @('CE-02', 'CE-03', 'CE-04') })
    if ($ce.Count -eq 0) { continue }
    $nB++; $dias = @{}; foreach ($t in $ce) { foreach ($k in $t.dias.Keys) { $dias[$k] = $true } }
    if ($dias.Count -lt [int]$b.diasCert) { $malB += "$($b.nombre): $($dias.Count) < $($b.diasCert) días" }
    $pu = @($TAREAS | Where-Object { $b.componentes -contains $_.componente -and $_.codigo -eq 'PU-02' } | ForEach-Object { $_.fin } | Measure-Object -Maximum).Maximum
    $ini = @($ce | ForEach-Object { $_.inicio } | Measure-Object -Minimum).Minimum
    if ($pu -and $ini -le $pu) { $malB += "$($b.nombre): certificación inicia el día $ini, antes del fin de PU-02 ($pu)" }
}
Add-Control 13 'Certificación por bloque con duración mínima' ($nB - $malB.Count) "$nB" ($malB.Count -eq 0) ($malB -join '; ')

# 14 Olas: tantas como se pidieron (1 = entrega única), cada una con componentes
$nOlas = [int](To-Num $WBS.proyecto.numeroOlas 1)
$olas = @($TAREAS | Where-Object { $_.ola -match '^Ola' } | ForEach-Object { $_.ola } | Sort-Object -Unique)
$faltan = @(1..$nOlas | Where-Object { $olas -notcontains "Ola $_" } | ForEach-Object { "Ola $_ sin tareas" })
Add-Control 14 'Las olas coinciden con lo pedido' $olas.Count "$nOlas" ($olas.Count -eq $nOlas -and $faltan.Count -eq 0) ($faltan -join '; ')

# 15 Días sin carga de desarrolladores dentro de su ventana
$sinCarga = 0; $ventana = 0; $det = @()
foreach ($r in $PLAN.recursos) {
    if ($r.rol -notmatch 'Developer|Data Engineer') { continue }
    $ds = @(); foreach ($t in $TAREAS) { if ($t.recurso -eq $r.nombre -and $t.inicio -gt 0) { $ds += $t.inicio; $ds += $t.fin } }
    if ($ds.Count -eq 0) { continue }
    $ini = ($ds | Measure-Object -Minimum).Minimum; $fin = ($ds | Measure-Object -Maximum).Maximum; $sc = 0
    for ($d = $ini; $d -le $fin; $d++) { if ((Cap $r.nombre $d) -gt 0) { $ventana++; if ($carga[$r.nombre][$d] -eq 0) { $sc++ } } }
    $sinCarga += $sc; if ($sc -gt 0) { $det += "$($r.nombre): $sc" }
}
$lim = [Math]::Floor(0.1 * $ventana)
Add-Control 15 'Días sin carga de desarrolladores dentro de su ventana' $sinCarga "≤ $lim (10 % de $ventana días-ventana)" ($sinCarga -le $lim) ($det -join '; ')

# 16 Dependencias
$dep = @()
foreach ($t in $TAREAS) {
    if ($t.tipo -in @('distribuida', 'relleno') -or $t.inicio -eq 0) { continue }
    foreach ($pk in $t.preds) {
        $q = $porId[$pk]; if ($null -eq $q -or $q.tipo -in @('distribuida', 'relleno')) { continue }
        $ok = if ($q.recurso -eq $t.recurso) { $t.inicio -ge $q.fin } else { $t.inicio -gt $q.fin }
        if (-not $ok) { $dep += "$($t.id) inicia $($t.inicio) antes del fin de $pk ($($q.fin))" }
    }
}
Add-Control 16 'Dependencias incumplidas' $dep.Count '0' ($dep.Count -eq 0) (($dep | Select-Object -First 10) -join '; ')

# 17 Orden de fases por componente
$ordF = @()
foreach ($c in $comps) {
    $ts = @($TAREAS | Where-Object { $_.componente -eq $c.id -and $_.inicio -gt 0 })
    $prevIni = 0; $prevFin = 0; $prevF = ''
    foreach ($f in @('02', '03', '04', '05', '06', '07', '08')) {
        $tf = @($ts | Where-Object { $_.fase -like "$f*" }); if ($tf.Count -eq 0) { continue }
        $ini = ($tf | ForEach-Object { $_.inicio } | Measure-Object -Minimum).Minimum; $fin = ($tf | ForEach-Object { $_.fin } | Measure-Object -Maximum).Maximum
        if ($prevF -and ($ini -lt $prevIni -or $fin -lt $prevFin)) { $ordF += "$($c.id): $f antes que $prevF" }
        $prevIni = $ini; $prevFin = $fin; $prevF = $f
    }
}
Add-Control 17 'Orden de fases por componente (02 → 08)' $ordF.Count '0' ($ordF.Count -eq 0) (($ordF | Select-Object -First 10) -join '; ')

# 18 Continuidad
$cong = @{}; foreach ($d in $PLAN.congelamiento) { $cong[[int]$d] = $true }
$ven = @{}; foreach ($d in $PLAN.ventanas) { $ven[[int]$d] = $true }
$huecos = @()
foreach ($t in $TAREAS) {
    if ($t.tipo -in @('distribuida', 'relleno') -or $t.inicio -eq 0) { continue }
    for ($d = [int]$t.inicio; $d -le [int]$t.fin; $d++) {
        if ((Cap $t.recurso $d) -le 0) { continue }
        if ($t.fase -match '^0[67]' -and $cong.ContainsKey($d)) { continue }
        if ($t.esPase -and $ven.Count -gt 0 -and -not $ven.ContainsKey($d)) { continue }
        if (-not $t.dias.Contains([string]$d)) { $huecos += "$($t.id) día $d"; break }
    }
}
Add-Control 18 'Tareas con huecos en su ejecución (continuidad)' $huecos.Count '0' ($huecos.Count -eq 0) (($huecos | Select-Object -First 10) -join '; ')

# 19 Construcción antes de cerrar los transversales prerrequisito
$pre = @($TAREAS | Where-Object { $_.codigo -in @('TRV-01', 'TRV-02', 'TRV-04', 'TRV-05', 'TRV-06', 'TRV-07') })
$finPre = ($pre | ForEach-Object { $_.fin } | Measure-Object -Maximum).Maximum
$x = @($TAREAS | Where-Object { ($_.codigo -like 'CO-*' -or $_.codigo -eq 'BC-03') -and $_.inicio -gt 0 -and $_.inicio -lt $finPre }).Count
Add-Control 19 'Construcciones antes de cerrar los transversales prerrequisito' $x '0' ($x -eq 0)

# 20 Fecha objetivo
$finF = $PLAN.fechas[$NDIAS - 1]; $obj = $WBS.proyecto.fechaObjetivo
if ($obj) { Add-Control 20 'Fecha fin ≤ fecha objetivo' $finF "≤ $obj" ($finF -le $obj) $(if ($finF -gt $obj) { 'Brecha: justificar o recomendar recursos adicionales' } else { '' }) }
else { Add-Control 20 'Fecha fin ≤ fecha objetivo' 'N/A' 'Sin fecha objetivo' $true }

# 22 Leyenda
$conColor = @{}; foreach ($l in $PLAN.leyenda) { $conColor[$l.aplicativoOrigen] = $true }
$x = @($TAREAS | ForEach-Object { $_.aplicativoOrigen } | Sort-Object -Unique | Where-Object { -not $conColor.ContainsKey($_) }).Count
Add-Control 22 'Cada aplicativo origen tiene color y entrada en la leyenda' $x '0' ($x -eq 0)
$x = @($PLAN.noProgramadas).Count
Add-Control 23 'Tareas sin programar' $x '0' ($x -eq 0) (@($PLAN.noProgramadas) -join ', ')

Write-JsonFile (Join-Path $Trabajo 'validaciones.json') $res
$lineas = @("Validaciones de planning.json — $((Get-Date).ToString('yyyy-MM-dd HH:mm'))", '')
foreach ($r in $res) { $lineas += ('{0,2}. [{1}] {2} = {3} (esperado {4}) {5}' -f $r.n, $r.estado, $r.control, $r.valor, $r.esperado, $r.detalle) }
[IO.File]::WriteAllLines((Join-Path $Trabajo 'validaciones.txt'), $lineas, (New-Object Text.UTF8Encoding $false))
$lineas | ForEach-Object { Write-Host $_ }
if (@($res | Where-Object { $_.estado -eq 'REVISAR' }).Count -gt 0) { exit 4 }
exit 0
