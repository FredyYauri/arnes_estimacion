<#
.SYNOPSIS
  Paso 9 del procedimiento: programa el Gantt diario (references/reglas-programacion.md) a partir de wbs.json.
  Nunca se asignan horas a mano: las tareas "fijadas" en ajustes.json se respetan tal cual y el resto
  se programa alrededor de ellas.
.OUTPUTS
  _trabajo/planning.json
.EXAMPLE
  .\3-Programar.ps1 -Trabajo 'Estimaciones\MigraMacro\_trabajo'
#>
param([Parameter(Mandatory = $true)][string]$Trabajo)
. (Join-Path $PSScriptRoot 'Comun.ps1')

$WBS = Read-JsonFile (Join-Path $Trabajo 'wbs.json')
$ENT = Read-JsonFile (Join-Path $Trabajo 'entrada.json')
$AJUSTES = Read-JsonFile (Join-Path $Trabajo 'ajustes.json')
if ($null -eq $WBS) { throw 'Falta wbs.json. Ejecute 2-Calcular-Estimacion.ps1.' }
$PROY = $WBS.proyecto
$MAXD = 1500
$avisos = New-Object System.Collections.Generic.List[string]

# ---------- Calendario ----------
$feriados = @($ENT.calendario | Where-Object { (Get-Val $_ 'Tipo') -in @('Feriado', 'No laborable') -and -not (Get-Val $_ 'Recurso (vacío = todos)') -and $_['Fecha'] } | ForEach-Object { [string]$_['Fecha'] } | Sort-Object -Unique)
$fechas = Get-DiasHabiles ([string]$PROY.fechaInicio) $MAXD $feriados ([bool]$PROY.sabado)
$diaDe = @{}; for ($i = 0; $i -lt $fechas.Count; $i++) { $diaDe[$fechas[$i]] = $i + 1 }
function DiaDesdeFecha([string]$iso) {
    if (-not $iso) { return 0 }
    for ($i = 0; $i -lt $fechas.Count; $i++) { if ($fechas[$i] -ge $iso) { return $i + 1 } }
    return $MAXD
}
$congelado = @{}; $ventanas = @{}
foreach ($c in $ENT.calendario) {
    $f = [string]$c['Fecha']; if (-not $f -or -not $diaDe.ContainsKey($f)) { continue }
    if ($c['Tipo'] -eq 'Congelamiento del cliente') { $congelado[$diaDe[$f]] = $true }
    if ($c['Tipo'] -eq 'Ventana de despliegue') { $ventanas[$diaDe[$f]] = $true }
}
foreach ($rs in $ENT.restricciones) {
    if ((Get-Val $rs 'Tipo') -eq 'Ventana de despliegue' -and (Get-Val $rs 'Fecha') -and $diaDe.ContainsKey([string]$rs['Fecha'])) { $ventanas[$diaDe[[string]$rs['Fecha']]] = $true }
}

# ---------- Capacidad por recurso y día ----------
$cap = @{}; $used = @{}
foreach ($r in $WBS.recursos) {
    $arr = New-Object int[] ($MAXD + 1); $u = New-Object int[] ($MAXD + 1)
    $c0 = [int]$r.cap
    foreach ($rs in $ENT.restricciones) {
        if ((Get-Val $rs 'Tipo') -eq 'Recurso compartido' -and ([string](Get-Val $rs 'Afecta a (IDs de Inventario)' '')).Contains($r.nombre)) {
            $pct = To-Num (Get-Val $rs 'Impacto / regla' 100) 100; if ($pct -le 1) { $pct = $pct * 100 }
            $c0 = [Math]::Max(1, [int][Math]::Floor($c0 * $pct / 100)); $avisos.Add("Recurso compartido: $($r.nombre) con $c0 h/día.")
        }
    }
    for ($d = 1; $d -le $MAXD; $d++) {
        $f = $fechas[$d - 1]
        $ok = (-not $r.desde -or $f -ge $r.desde) -and (-not $r.hasta -or $f -le $r.hasta)
        $arr[$d] = if ($ok) { $c0 } else { 0 }
    }
    foreach ($c in $ENT.calendario) {
        if ((Get-Val $c 'Recurso (vacío = todos)') -eq $r.nombre -and $c['Tipo'] -in @('Vacaciones', 'Feriado', 'No laborable') -and $diaDe.ContainsKey([string]$c['Fecha'])) { $arr[$diaDe[[string]$c['Fecha']]] = 0 }
    }
    $cap[$r.nombre] = $arr; $used[$r.nombre] = $u
}

# ---------- Estado de tareas ----------
$TAREAS = @($WBS.tareas)
$porClave = @{}; foreach ($t in $TAREAS) { $porClave[$t.clave] = $t; $t['rest'] = [int]$t.horas; $t['inicio'] = 0; $t['fin'] = 0; $t['dias'] = [ordered]@{}; $t['fijada'] = $false }
$compInfo = @{}; foreach ($c in $WBS.componentes) { $compInfo[$c.id] = $c }
foreach ($t in $TAREAS) {
    $t['minDia'] = 0
    if ((Get-Val $t 'minFecha')) { $t.minDia = DiaDesdeFecha ([string]$t.minFecha) }
    $t['esFase67'] = ($t.fase -like '06*' -or $t.fase -like '07*')
    if (-not $t.Contains('esPase')) { $t['esPase'] = $false }
}

function Add-Horas($t, [int]$d, [int]$h) {
    if ($h -le 0) { return }
    $k = [string]$d
    if ($t.dias.Contains($k)) { $t.dias[$k] = [int]$t.dias[$k] + $h } else { $t.dias[$k] = $h }
    $used[$t.recurso][$d] += $h
    $t.rest -= $h
    if ($t.inicio -eq 0 -or $d -lt $t.inicio) { $t.inicio = $d }
    if ($t.rest -le 0) { $t.fin = $d }
}

# Tareas fijadas por el usuario (ajustes.json → fijadas): se respetan sus horas por día.
$fij = Get-Val $AJUSTES 'fijadas' ([ordered]@{})
foreach ($k in $fij.Keys) {
    if (-not $porClave.ContainsKey($k)) { $avisos.Add("Fijada '$k' no corresponde a ninguna tarea; se ignora."); continue }
    $t = $porClave[$k]; $dias = Get-Val $fij[$k] 'dias' ([ordered]@{})
    if (Get-Val $fij[$k] 'recurso') { $t.recurso = [string]$fij[$k].recurso }
    $t.rest = 0; $tot = 0
    foreach ($dk in ($dias.Keys | Sort-Object { [int]$_ })) {
        $h = [int](To-Num $dias[$dk]); $d = [int]$dk
        $t.dias[[string]$d] = $h; $used[$t.recurso][$d] += $h; $tot += $h
        if ($t.inicio -eq 0 -or $d -lt $t.inicio) { $t.inicio = $d }
        if ($d -gt $t.fin) { $t.fin = $d }
    }
    $t.horas = $tot; $t.fijada = $true
    $t.observacion = (([string]$t.observacion) + ' Fijada por el usuario (horas por día).').Trim()
}

# ---------- Prioridad: ruta crítica restante ----------
$suc = @{}; foreach ($t in $TAREAS) { foreach ($p in $t.preds) { if (-not $suc.ContainsKey($p)) { $suc[$p] = New-Object System.Collections.Generic.List[string] }; $suc[$p].Add($t.clave) } }
$capDe = @{}; foreach ($r in $WBS.recursos) { $capDe[$r.nombre] = [Math]::Max(1, [int]$r.cap) }
# Orden topológico (Kahn) y ruta crítica restante en días, recorrida de atrás hacia adelante (sin recursión).
$grado = @{}; foreach ($t in $TAREAS) { $grado[$t.clave] = @($t.preds | Where-Object { $porClave.ContainsKey($_) }).Count }
$cola0 = New-Object System.Collections.Generic.Queue[string]; foreach ($t in $TAREAS) { if ($grado[$t.clave] -eq 0) { $cola0.Enqueue($t.clave) } }
$topo = New-Object System.Collections.Generic.List[string]
while ($cola0.Count -gt 0) {
    $k = $cola0.Dequeue(); $topo.Add($k)
    if ($suc.ContainsKey($k)) { foreach ($s in $suc[$k]) { $grado[$s]--; if ($grado[$s] -eq 0) { $cola0.Enqueue($s) } } }
}
if ($topo.Count -lt $TAREAS.Count) { throw 'El WBS tiene dependencias circulares; revise Orden obligatorio y la serie de desarrolladores.' }
$memo = @{}
for ($j = $topo.Count - 1; $j -ge 0; $j--) {
    $k = $topo[$j]; $t = $porClave[$k]
    $dur = if ($t.tipo -eq 'tasa_fija') { [int]$t.diasMin } else { [int][Math]::Ceiling($t.horas / [double]$capDe[$t.recurso]) }
    $m = 0; if ($suc.ContainsKey($k)) { foreach ($s in $suc[$k]) { if ($memo[$s] -gt $m) { $m = $memo[$s] } } }
    $memo[$k] = $dur + $m
}
foreach ($t in $TAREAS) {
    $t['critica'] = $memo[$t.clave]
    $ci = $compInfo[$t.componente]
    $t['olaNum'] = $(if ($t.ola -match '\d+') { [int]$Matches[0] } else { 0 })
    $t['prio'] = $(if ($ci) { [int]$ci.prioridad } else { 2 })
    $t['ordenSerie'] = $(if ($ci) { [int]$ci.orden } else { 0 })
}
$ordenPrio = @($TAREAS | Sort-Object @{e = { $_.critica }; Descending = $true }, @{e = { $_.olaNum } }, @{e = { $_.prio } }, @{e = { $_.ordenSerie } })

function Test-Lista($t, [int]$d) {
    if ($t.rest -le 0 -or $t.fijada) { return $false }
    if ($t.minDia -gt 0 -and $d -lt $t.minDia) { return $false }
    if ($t.esFase67 -and $congelado.ContainsKey($d)) { return $false }
    if ($t.esPase -and $ventanas.Count -gt 0 -and -not $ventanas.ContainsKey($d)) { return $false }
    foreach ($p in $t.preds) {
        $q = $porClave[$p]; if ($null -eq $q) { continue }
        if ($q.tipo -in @('distribuida', 'relleno')) { continue }
        if ($q.rest -gt 0 -and -not $q.fijada) { return $false }
        if ($q.fin -eq 0) { return $false }
        if ($q.recurso -eq $t.recurso) { if ($q.fin -gt $d) { return $false } } else { if ($q.fin -ge $d) { return $false } }
    }
    return $true
}

# ---------- List scheduling por día ----------
$pendientes = @($TAREAS | Where-Object { $_.tipo -in @('normal', 'tasa_fija') -and $_.rest -gt 0 -and -not $_.fijada })
$porRecurso = @{}; foreach ($r in $WBS.recursos) { $porRecurso[$r.nombre] = @($ordenPrio | Where-Object { $_.recurso -eq $r.nombre -and $_.tipo -in @('normal', 'tasa_fija') }) }
$d = 0; $sinAvance = 0
while (@($pendientes | Where-Object { $_.rest -gt 0 }).Count -gt 0 -and $d -lt $MAXD) {
    $d++; $avance = $false
    foreach ($r in $WBS.recursos) {
        $n = $r.nombre; $cola = $porRecurso[$n]
        if ($null -eq $cola -or $cola.Count -eq 0) { continue }
        $free = $cap[$n][$d] - $used[$n][$d]
        if ($cap[$n][$d] -le 0) { continue }
        # 1) Tasa fija activa o lista
        foreach ($t in $cola) {
            if ($t.tipo -ne 'tasa_fija' -or $t.rest -le 0 -or $t.fijada) { continue }
            $activa = $t.inicio -gt 0
            if (-not $activa -and -not (Test-Lista $t $d)) { continue }
            if ($t.esFase67 -and $congelado.ContainsKey($d)) { continue }
            $h = [Math]::Min([Math]::Min([int]$t.hDia, $t.rest), $free)
            if ($h -lt 1) { if ($activa) { $avisos.Add("Día ${d}: $($t.clave) sin capacidad para su tasa fija.") }; continue }
            Add-Horas $t $d $h; $free -= $h; $avance = $true
        }
        # 2) Normales ya iniciadas: mínimo 1 h cada una (continuidad), luego el resto en orden de inicio
        $iniciadas = @($cola | Where-Object { $_.tipo -eq 'normal' -and $_.inicio -gt 0 -and $_.rest -gt 0 -and -not $_.fijada } | Sort-Object { $_.inicio })
        foreach ($t in $iniciadas) {
            if ($t.esFase67 -and $congelado.ContainsKey($d)) { continue }
            if ($t.esPase -and $ventanas.Count -gt 0 -and -not $ventanas.ContainsKey($d)) { continue }
            if ($free -ge 1) { Add-Horas $t $d 1; $free -= 1; $avance = $true } else { $avisos.Add("Día ${d}: $($t.clave) sin capacidad (hueco de continuidad).") }
        }
        foreach ($t in $iniciadas) {
            if ($free -le 0 -or $t.rest -le 0) { continue }
            if ($t.dias.Contains([string]$d) -eq $false) { continue }
            $h = [Math]::Min($t.rest, $free); Add-Horas $t $d $h; $free -= $h
        }
        # 3) Normales listas, por prioridad; máximo 2 normales en curso por recurso
        $cambio = $true
        while ($free -gt 0 -and $cambio) {
            $cambio = $false
            $enCurso = @($cola | Where-Object { $_.tipo -eq 'normal' -and $_.inicio -gt 0 -and $_.rest -gt 0 }).Count
            foreach ($t in $cola) {
                if ($free -le 0) { break }
                if ($t.tipo -ne 'normal' -or $t.inicio -gt 0) { continue }
                if ($enCurso -ge 2) { break }
                if (-not (Test-Lista $t $d)) { continue }
                $h = [Math]::Min($t.rest, $free); Add-Horas $t $d $h; $free -= $h; $avance = $true; $cambio = $true
                if ($t.rest -gt 0) { $enCurso++ }
            }
        }
    }
    if ($avance) { $sinAvance = 0 } else { $sinAvance++ }
    if ($sinAvance -gt 260) { break }
}
$noProgramadas = @($pendientes | Where-Object { $_.rest -gt 0 } | ForEach-Object { $_.clave })
foreach ($k in $noProgramadas) { $avisos.Add("No se pudo programar $k (revise disponibilidad del recurso, dependencias o ventanas).") }

$ultimo = 0; foreach ($t in $TAREAS) { if ($t.fin -gt $ultimo) { $ultimo = $t.fin } }
if ($ultimo -eq 0) { $ultimo = 1 }

# ---------- Hitos (gates con criterio de salida) ----------
# G0 Inicio · G1 Análisis cerrado · G2 Diseño aprobado · G3 Listo para certificar y G4 Certificado (por bloque)
# · G5 Pase (por pase) · G6 Cierre (si hay hypercare). Ajustes del pedido: asignacion.json → gates
# { "G1": { "criterio": "...", "aprueba": "..." }, "G2": { "omitir": true } } (la clave sin el sufijo de bloque aplica a todos).
$ASIG = Read-JsonFile (Join-Path $Trabajo 'asignacion.json')
$ajGates = Get-Val $ASIG 'gates' ([ordered]@{})
$hitos = New-Object System.Collections.Generic.List[object]
function Fin-De($ts) { $m = @($ts | Where-Object { $_.fin -gt 0 } | ForEach-Object { $_.fin } | Measure-Object -Maximum).Maximum; if ($m) { return [int]$m }; return 0 }
function Add-Gate([string]$cod, [string]$nombre, [int]$dia, [string]$criterio, [string]$aprueba) {
    if ($dia -le 0) { return }
    $base = ($cod -split '\.')[0]
    foreach ($k in @($base, $cod)) {
        $a = Get-Val $ajGates $k $null; if ($null -eq $a) { continue }
        if (Is-Si ([string](Get-Val $a 'omitir' 'No')) -or (Get-Val $a 'omitir') -eq $true) { return }
        $criterio = [string](Get-Val $a 'criterio' $criterio); $aprueba = [string](Get-Val $a 'aprueba' $aprueba); $nombre = [string](Get-Val $a 'nombre' $nombre)
    }
    $script:hitos.Add([ordered]@{ codigo = $cod; nombre = $nombre; dia = $dia; criterio = $criterio; aprueba = $aprueba })
}
Add-Gate 'G0' 'Inicio' (Fin-De @($TAREAS | Where-Object { $_.fase -like '01*' })) 'Accesos, ambientes y repositorios listos; alcance e inventario confirmados' 'Jefe de Proyecto + cliente'
Add-Gate 'G1' 'Análisis cerrado' (Fin-De @($TAREAS | Where-Object { $_.fase -like '02*' })) 'Dependencias e incompatibilidades identificadas, reglas validadas con negocio, desvío de la re-estimación < 10 %' 'Líder Técnico'
Add-Gate 'G2' 'Diseño aprobado' (Fin-De @($TAREAS | Where-Object { $_.fase -like '03*' })) 'Arquitectura, modelo de datos y prototipos aprobados' 'Arquitecto + cliente'
$nb = 0; $varios = @($WBS.bloques).Count -gt 1
foreach ($b in $WBS.bloques) {
    $nb++; $suf = $(if ($varios) { ".$nb" } else { '' })
    $pu = Fin-De @($TAREAS | Where-Object { $b.componentes -contains $_.componente -and $_.codigo -eq 'PU-02' })
    $ce = @($TAREAS | Where-Object { $_.bloque -eq $b.nombre -and $_.fase -like '06*' })
    $ini = @($ce | Where-Object { $_.inicio -gt 0 -and $_.codigo -ne 'CE-01' } | ForEach-Object { $_.inicio } | Measure-Object -Minimum).Minimum
    $fin = Fin-De $ce
    $b['listo'] = [int]$pu; $b['inicioCert'] = [int]$ini; $b['finCert'] = [int]$fin
    Add-Gate "G3$suf" "Listo para certificar – $($b.nombre)" $pu 'Código revisado, pruebas unitarias en verde y sin defectos críticos abiertos' 'Líder Técnico'
    Add-Gate "G4$suf" "Certificado – $($b.nombre)" $fin 'Casos de QA ejecutados, 0 defectos bloqueantes y acta de certificación firmada' 'QA + negocio'
}
$np = 0; $variosP = @($WBS.pases).Count -gt 1
foreach ($pa in $WBS.pases) {
    $np++; $suf = $(if ($variosP) { ".$np" } else { '' })
    $de = Fin-De @($TAREAS | Where-Object { $_.bloque -eq $pa.nombre -and $_.codigo -eq 'DE-04' })
    $pa['dia'] = [int]$de
    Add-Gate "G5$suf" "Pase a producción – $($pa.nombre)" $de 'Despliegue validado en producción y plan de rollback probado' 'Jefe de Proyecto + operaciones'
    Add-Gate "G6$suf" "Cierre – $($pa.nombre)" (Fin-De @($TAREAS | Where-Object { $_.bloque -eq $pa.nombre -and $_.fase -like '08*' })) 'Sin incidentes abiertos y transferencia a soporte firmada' 'Cliente'
}
$hitos = @($hitos | Sort-Object { [int]$_.dia })
$diaHito = @{}; foreach ($h in $hitos) { $diaHito[[int]$h.dia] = $true }

# ---------- Gestión distribuida (TRV-GES) ----------
$ges = $porClave['TRV-GES']
if ($ges -and -not $ges.fijada) {
    $n = $ges.recurso; $horasGes = [int]$ges.horas
    $diasG = New-Object System.Collections.Generic.List[int]
    for ($x = 1; $x -le $ultimo; $x++) { if ($cap[$n][$x] - $used[$n][$x] -ge 1) { $diasG.Add($x) } }
    $asig = @{}; foreach ($x in $diasG) { $asig[$x] = 1 }
    $resto = $horasGes - $diasG.Count
    if ($resto -gt 0) {
        foreach ($x in $diasG) { if ($resto -le 0) { break }; $lunes = ((Get-IsoDate $fechas[$x - 1]).DayOfWeek -eq 'Monday'); if (($lunes -or $diaHito.ContainsKey($x)) -and ($cap[$n][$x] - $used[$n][$x] - $asig[$x]) -ge 1) { $asig[$x]++; $resto-- } }
        $vuelta = $true
        while ($resto -gt 0 -and $vuelta) { $vuelta = $false; foreach ($x in $diasG) { if ($resto -le 0) { break }; if (($cap[$n][$x] - $used[$n][$x] - $asig[$x]) -ge 1) { $asig[$x]++; $resto--; $vuelta = $true } } }
        if ($resto -gt 0) { $avisos.Add("TRV-GES: $resto h no caben en la capacidad de $n.") }
    }
    $totG = 0; foreach ($x in $asig.Keys) { $totG += $asig[$x] }
    $asumidas = [Math]::Max(0, $totG - $horasGes)
    $ges.rest = $totG - $asumidas
    if ($asumidas -gt 0) {
        # La diferencia hasta 1 h/día se separa en una fila Asumida, repartida en forma pareja.
        $aj = [ordered]@{}; foreach ($kk in $ges.Keys) { $aj[$kk] = $ges[$kk] }
        $aj.clave = 'TRV-GES-AJ'; $aj.codigo = 'TRV-GES'; $aj.tarea = 'Gestión: completar hasta 1 h por día hábil del proyecto'; $aj.origen = 'Asumida'
        $aj.driver = "1 h/día × $totG días − $horasGes h"; $aj.horasBase = $asumidas; $aj.horas = $asumidas; $aj.dias = [ordered]@{}; $aj.inicio = 0; $aj.fin = 0; $aj.rest = $asumidas
        $TAREAS += $aj; $porClave['TRV-GES-AJ'] = $aj
        $acc = 0; $i = 0
        foreach ($x in ($diasG | Sort-Object)) {
            $i++; $objetivo = [int][Math]::Floor($i * $asumidas / [double]$diasG.Count)
            $h = $asig[$x]
            if ($objetivo -gt $acc) { Add-Horas $aj $x 1; $acc++; $h-- }
            if ($h -gt 0) { Add-Horas $ges $x $h }
        }
        $ges.horas = $horasGes
    }
    else { foreach ($x in ($asig.Keys | Sort-Object)) { Add-Horas $ges $x $asig[$x] }; $ges.horas = $totG }
}
$aco = $porClave['TRV-ACO']
if ($aco) {
    $n = $aco.recurso
    $diasLT = @($TAREAS | Where-Object { $_.recurso -eq $n -and $_.inicio -gt 0 })
    $iniLT = ($diasLT | ForEach-Object { $_.inicio } | Measure-Object -Minimum).Minimum; $finLT = ($diasLT | ForEach-Object { $_.fin } | Measure-Object -Maximum).Maximum
    $aco.rest = 0
    for ($x = [int]$iniLT; $x -le [int]$finLT; $x++) { if ($used[$n][$x] -eq 0 -and $cap[$n][$x] -ge 1) { $aco.rest++; Add-Horas $aco $x 1 } }
    $aco.horas = 0; foreach ($v in $aco.dias.Values) { $aco.horas += [int]$v }
}

# ---------- Orden final e IDs ----------
$ordFase = @{}; $i = 0; foreach ($f in @('00', '01.1', '01.2', '02', '03', '04', '05', '06', '07', '08')) { $ordFase[$f] = $i; $i++ }
$ordBloque = @{}; $i = 0; foreach ($b in $WBS.bloques) { $ordBloque[$b.nombre] = $i; $i++ }; foreach ($pa in $WBS.pases) { if (-not $ordBloque.ContainsKey($pa.nombre)) { $ordBloque[$pa.nombre] = $i; $i++ } }
foreach ($t in $TAREAS) {
    $grupo = if ($t.ola -eq 'Transversal') { 0 } elseif ($t.bloque) { 2 } else { 1 }
    $f = ($t.fase -split ' ')[0]
    $t['_k'] = '{0}|{1:D3}|{2}|{3:D4}|{4:D2}|{5:D5}|{6}' -f $grupo, $(if ($grupo -eq 2) { [int]$ordBloque[$t.bloque] } else { [int]$t.olaNum }), $(if ($grupo -eq 1) { $t.aplicativo } else { '' }),
    $(if ($t.componente -like 'Base común*') { 0 } else { [int]$t.ordenSerie + 1 }), [int]$ordFase[$f], [int]$t.inicio, $t.clave
}
$TAREAS = @($TAREAS | Sort-Object { $_._k })
$i = 0; $idDe = @{}
foreach ($t in $TAREAS) { $i++; $t['id'] = 'T{0:D3}' -f $i; $idDe[$t.clave] = $t.id }

# ---------- Leyenda ----------
$paleta = @('#BDD7EE', '#C6E0B4', '#F8CBAD', '#FFE699', '#D9C3E9', '#B4DEDE', '#F4B6C2', '#D6DCE4', '#E2D3B5', '#DDD9C3', '#B7C9E2', '#E4DFEC')
$leyenda = New-Object System.Collections.Generic.List[object]; $vistos = @{}
foreach ($t in $TAREAS) {
    $a = $t.aplicativoOrigen
    if ($a -in @('Transversal', 'Base común', 'Varios') -or $vistos.ContainsKey($a)) { continue }
    $vistos[$a] = $true; $leyenda.Add([ordered]@{ aplicativoOrigen = $a; color = $paleta[($leyenda.Count) % 12] })
}
foreach ($g in @(@('Transversal', '#E7E6E6'), @('Base común', '#D9D9D9'), @('Varios', '#EDEDED'))) {
    if (@($TAREAS | Where-Object { $_.aplicativoOrigen -eq $g[0] }).Count -gt 0) { $leyenda.Add([ordered]@{ aplicativoOrigen = $g[0]; color = $g[1] }) }
}

# ---------- Salida ----------
$tareasOut = New-Object System.Collections.Generic.List[object]
foreach ($t in $TAREAS) {
    $tareasOut.Add([ordered]@{
            id = $t.id; clave = $t.clave; codigo = $t.codigo; ola = $t.ola; fase = $t.fase; aplicativoOrigen = $t.aplicativoOrigen; aplicativo = $t.aplicativo
            plataforma = $t.plataforma; componente = $t.componente; componenteNombre = $t.componenteNombre; complejidad = $t.complejidad
            tarea = $t.tarea; recurso = $t.recurso; horas = [int]$t.horas; horasProgramadas = [int]$(($t.dias.Values | Measure-Object -Sum).Sum)
            horasBase = $t.horasBase; factorIA = $t.factorIA; origen = $t.origen; driver = $t.driver; observacion = $t.observacion
            tipo = $t.tipo; bloque = $t.bloque; fijada = $t.fijada; esPase = $t.esPase; diasMin = $t.diasMin
            preds = @($t.preds | Where-Object { $idDe.ContainsKey($_) } | ForEach-Object { $idDe[$_] }); predsClaves = @($t.preds)
            inicio = [int]$t.inicio; fin = [int]$t.fin; dias = $t.dias
        })
}
$recOut = New-Object System.Collections.Generic.List[object]
foreach ($r in $WBS.recursos) {
    $capDia = New-Object System.Collections.Generic.List[int]; for ($x = 1; $x -le $ultimo; $x++) { $capDia.Add($cap[$r.nombre][$x]) }
    $o = [ordered]@{}; foreach ($k in $r.Keys) { $o[$k] = $r[$k] }; $o['capDia'] = $capDia
    $recOut.Add($o)
}
$planning = [ordered]@{
    generado = (Get-Date).ToString('yyyy-MM-dd HH:mm'); fechaInicio = $PROY.fechaInicio; feriados = $feriados
    fechas = @($fechas | Select-Object -First $ultimo); ultimoDia = $ultimo; jornada = $PROY.jornada
    congelamiento = @($congelado.Keys); ventanas = @($ventanas.Keys)
    hitos = $hitos; leyenda = $leyenda; recursos = $recOut; tareas = $tareasOut
    bloques = $WBS.bloques; pases = $WBS.pases; noProgramadas = $noProgramadas; avisos = @($avisos | Select-Object -Unique)
}
Write-JsonFile (Join-Path $Trabajo 'planning.json') $planning
$tot = 0; foreach ($t in $tareasOut) { $tot += $t.horasProgramadas }
Write-Host "Planning: $($tareasOut.Count) tareas, $ultimo días hábiles (fin $($fechas[$ultimo - 1])), $tot h."
if ($noProgramadas.Count -gt 0) { Write-Host "ATENCIÓN: $($noProgramadas.Count) tareas sin programar."; exit 3 }
