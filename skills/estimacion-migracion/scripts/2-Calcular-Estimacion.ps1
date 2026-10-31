<#
.SYNOPSIS
  Pasos 4-8 del procedimiento: clasifica complejidad, arma el WBS y calcula horas con el modelo
  paramétrico (references/modelo-horas.md y catalogo-fases-tareas.md). Todo cálculo es determinístico:
  la misma entrada + asignación + ajustes produce siempre el mismo wbs.json.
.DESCRIPTION
  Entradas en _trabajo/:
    entrada.json     (1-Leer-Entrada.ps1)
    asignacion.json  (lo arma el agente: decisiones de criterio — ola, desarrollador y orden por componente,
                      responsables por rol, recursos Propuestos autorizados). Ver scripts/README.md.
    ajustes.json     (opcional: cambios manuales del usuario que deben sobrevivir a cada regeneración).
  Salida: _trabajo/wbs.json
.EXAMPLE
  .\2-Calcular-Estimacion.ps1 -Trabajo 'Estimaciones\MigraMacro\_trabajo'
#>
param([Parameter(Mandatory = $true)][string]$Trabajo)
. (Join-Path $PSScriptRoot 'Comun.ps1')

$ENT = Read-JsonFile (Join-Path $Trabajo 'entrada.json')
if ($null -eq $ENT) { throw "No existe $Trabajo\entrada.json. Ejecute 1-Leer-Entrada.ps1." }
$ASIG = Read-JsonFile (Join-Path $Trabajo 'asignacion.json')
if ($null -eq $ASIG) { throw "No existe $Trabajo\asignacion.json (lo prepara el agente; ver scripts/README.md)." }
$AJUSTES = Read-JsonFile (Join-Path $Trabajo 'ajustes.json')
$PROY = $ENT.proyecto
$errores = New-Object System.Collections.Generic.List[string]

# ---------- Parámetros ----------
$defectos = [ordered]@{
    'Productividad LOC/h (sin IA)' = 45; 'Factor complejidad Baja' = 0.8; 'Factor complejidad Media' = 1.0; 'Factor complejidad Alta' = 1.3
    'Horas por pantalla Baja' = 4; 'Horas por pantalla Media' = 8; 'Horas por pantalla Alta' = 16
    'Horas por integración Baja' = 6; 'Horas por integración Media' = 12; 'Horas por integración Alta' = 24
    'Horas por reporte Baja' = 2; 'Horas por reporte Media' = 4; 'Horas por reporte Alta' = 8
    'Horas por tabla Baja' = 1; 'Horas por tabla Media' = 2; 'Horas por tabla Alta' = 4
    'Factor IA Análisis' = 0.5; 'Factor IA Diseño' = 0.75; 'Factor IA Construcción' = 0.6; 'Factor IA Revisión de código' = 1.0
    'Factor IA Preparar casos PU' = 0.9; 'Factor IA Ejecutar PU' = 0.5; 'Factor IA Manuales' = 0.4; 'Factor IA Otros' = 1.0
    '% Gestión del proyecto' = 8; '% Errores y secretos' = 10; '% Revisión de código' = 5; '% Preparar casos PU' = 5; '% Ejecutar PU' = 15
    'Contingencia conocido' = 3; 'Contingencia integración' = 8; 'Contingencia requerimientos poco claros' = 15
    'Contingencia dependencia externa' = 15; 'Contingencia componente IA' = 15
    'Soporte dev en certificación (h/día)' = 4; 'Ejecución QA (h/día)' = 6; 'Coordinación LT en certificación (h/día)' = 2
    'Hypercare dev (h/día por aplicativo)' = 2; 'Hypercare LT (h/día)' = 2; 'Módulos por lote' = 10
}
function Par([string]$n) { return To-Num (Get-Val $ENT.parametros $n $defectos[$n]) $defectos[$n] }
function ParTxt([string]$n, [string]$d) { return [string](Get-Val $ENT.parametros $n $d) }

$usaIA = Is-Si (Get-Val $PROY 'Uso de asistente IA en el desarrollo' 'Sí')
function FIA([string]$tipo) { if (-not $usaIA) { return 1.0 }; return Par "Factor IA $tipo" }
$jornada = [int](To-Num (Get-Val $PROY 'Jornada (h/día)' 8) 8)
$tipoProyecto = [string](Get-Val $PROY 'Tipo de trabajo' 'Migración')
$alcance = [string](Get-Val $PROY 'La estimación llega hasta' 'Producción')
$alcanceProd = $alcance -in @('Producción', 'Hypercare')
$alcanceHyper = $alcance -eq 'Hypercare'
# Olas: si el pedido indica n olas, el agente lo pone en asignacion.numeroOlas (manda sobre la entrada).
# Si no se pidió nada, es una sola entrega.
$nOlas = Get-Val $ASIG 'numeroOlas'
if ($null -eq $nOlas) {
    $nOlas = if ([string](Get-Val $PROY 'Modelo de entrega' 'Entrega única') -eq 'Oleadas') { Get-Val $PROY 'Número de olas' 1 } else { 1 }
}
$nOlas = [int](To-Num $nOlas 1); if ($nOlas -lt 1) { $nOlas = 1 }
$modelo = if ($nOlas -gt 1) { 'Oleadas' } else { 'Entrega única' }
$agrupacion = [string](Get-Val $PROY 'Agrupación de certificación y despliegue' 'Por ola')
$certA = [string](Get-Val $PROY 'Certificación a cargo de' 'QA proveedor')
$qaProveedor = $certA -in @('QA proveedor', 'Mixta')
$sensibles = Is-Si (Get-Val $PROY 'Datos sensibles o exposición a internet' 'No')
$claridad = [string](Get-Val $PROY 'Claridad de requerimientos' 'Media')
$diasCertMin = [int](To-Num (Get-Val $PROY 'Duración mínima de certificación (días hábiles)' 5) 5)
$diasHyper = [int](To-Num (Get-Val $PROY 'Duración de hypercare (días hábiles)' 5) 5)
$redondeo4 = Is-Si (ParTxt 'Redondeo a bloques de 4 h' 'No')
# Por defecto cada recurso se planifica a la jornada completa; "Capacidad productiva del rol" es opcional.
$cargaJornada = (ParTxt 'Carga diaria máxima' 'Jornada') -eq 'Jornada'
$retiroLegacy = Is-Si (ParTxt 'Incluir retiro del legacy' 'Sí')

$FASES = [ordered]@{
    '00' = '00 Gestión'; '01.1' = '01.1 Preparación – Entendimiento inicial'; '01.2' = '01.2 Preparación – Relevamiento y preparación técnica'
    '02' = '02 Análisis'; '03' = '03 Diseño'; '04' = '04 Construcción'; '05' = '05 Pruebas unitarias'
    '06' = '06 Certificación'; '07' = '07 Despliegue'; '08' = '08 Seguimiento en producción'
}

# ---------- Recursos ----------
$recursos = New-Object System.Collections.Generic.List[object]
function Add-Recurso($nombre, $rol, $capH, $ded, $desde, $hasta, $plat, $tarifa, [bool]$propuesto) {
    $capRol = To-Num (Get-Val $ENT.capacidadRol $rol $jornada) $jornada
    $base = if ($cargaJornada) { $jornada } elseif ($null -ne $capH -and (To-Num $capH) -gt 0) { To-Num $capH } else { $capRol }
    $d = To-Num $ded 100; if ($d -le 1) { $d = $d * 100 }
    $cap = [int][Math]::Floor([Math]::Min($base * $d / 100, $jornada))
    if ($cap -lt 1) { $cap = 1 }
    $recursos.Add([ordered]@{ nombre = [string]$nombre; rol = [string]$rol; cap = $cap; desde = $desde; hasta = $hasta
            plataformas = [string]$plat; tarifa = $tarifa; propuesto = $propuesto })
}
foreach ($r in $ENT.recursos) {
    Add-Recurso (Get-Val $r 'Recurso (nombre o alias)') (Get-Val $r 'Rol') (Get-Val $r 'Capacidad h/día (opcional)') (Get-Val $r 'Dedicación %' 100) `
    (Get-Val $r 'Disponible desde') (Get-Val $r 'Disponible hasta') (Get-Val $r 'Plataformas / skills' '') (Get-Val $r 'Tarifa por hora (opcional)') $false
}
foreach ($r in @(Get-Val $ASIG 'propuestos' @())) {
    Add-Recurso $r.nombre $r.rol (Get-Val $r 'cap') (Get-Val $r 'dedicacion' 100) (Get-Val $r 'desde') (Get-Val $r 'hasta') (Get-Val $r 'plataformas' '') (Get-Val $r 'tarifa') $true
}
$porNombre = @{}; foreach ($r in $recursos) { $porNombre[$r.nombre] = $r }

function Find-Rec([string]$patron) {
    foreach ($r in $recursos) { if ($r.rol -match $patron) { return $r.nombre } }
    return $null
}
$rolesA = Get-Val $ASIG 'roles' ([ordered]@{})
function Get-Responsable([string]$clave) {
    $v = Get-Val $rolesA $clave
    if ($v) { return [string]$v }
    switch ($clave) {
        'Líder Técnico' { return Find-Rec '^Líder Técnico' }
        'Jefe de Proyecto' { $x = Find-Rec '^Jefe de Proyecto'; if ($x) { return $x }; return Get-Responsable 'Líder Técnico' }
        'Arquitecto de Aplicaciones' { $x = Find-Rec '^Arquitecto de Aplicaciones'; if ($x) { return $x }; return Get-Responsable 'Líder Técnico' }
        'Especialista de Plataformas' { $x = Find-Rec '^Especialista de Plataformas'; if ($x) { return $x }; $x = Get-Responsable 'Developer Sr'; if ($x) { return $x }; return Get-Responsable 'Líder Técnico' }
        'Developer Sr' { $x = Find-Rec '^(Developer|Data Engineer).* Sr'; if ($x) { return $x }; return Get-Responsable 'Developer' }
        'Developer' { return Find-Rec '^(Developer|Data Engineer)' }
        'Analista Calidad Sr' { $x = Find-Rec '^Analista Calidad Sr'; if ($x) { return $x }; return Find-Rec '^(Analista Calidad|Tester)' }
        'Analista Calidad Ssr' { $x = Find-Rec '^(Analista Calidad Ssr|Tester)'; if ($x) { return $x }; return Find-Rec '^Analista Calidad' }
        'Seguridad' { return Find-Rec 'hacking' }
        'UX' { $x = Find-Rec '^Especialista (UX|UI)'; if ($x) { return $x }; return Get-Responsable 'Developer' }
        'Datos' { $x = Find-Rec '^(Arquitecto de Datos|Data Engineer|Modelador)'; if ($x) { return $x }; return Get-Responsable 'Líder Técnico' }
    }
    return $null
}
$LT = Get-Responsable 'Líder Técnico'
if (-not $LT) { $errores.Add('No hay Líder Técnico en Recursos ni en asignacion.roles.') }

# ---------- Componentes ----------
function Nivel3([string]$v, [string[]]$escala) { $i = [array]::IndexOf($escala, $v); if ($i -lt 0) { return 1 }; return $i + 1 }
$componentes = New-Object System.Collections.Generic.List[object]
$asigComp = Get-Val $ASIG 'componentes' ([ordered]@{})
# Filas agregadas: una fila con "Módulos (n)" > 1 (aplicativo grande sin detalle por módulo) se parte en lotes
# de módulos para que varios desarrolladores trabajen en paralelo. La complejidad se mide por módulo promedio y
# las cantidades (LOC, pantallas, etc.) se reparten entre los lotes, así las horas totales salen de los totales.
# Lotes = asignacion.componentes.<ID>.lotes; si no, uno por desarrollador de "desarrolladores"; si no, ceil(n / Módulos por lote).
function Split-Int([double]$total, [int]$k, [int]$j) { return [Math]::Floor($total * $j / $k) - [Math]::Floor($total * ($j - 1) / $k) }
$camposReparto = @('LOC', 'Pantallas / formularios', 'Reportes / salidas', 'Integraciones (n)', 'Tablas / hojas / fuentes de datos (n)', 'Horas override (opcional)')
$filasInv = New-Object System.Collections.Generic.List[object]
$lotesDe = @{}   # ID de Inventario -> IDs de componente (el mismo ID o sus lotes)
foreach ($r0 in $ENT.inventario) {
    $id0 = [string]$r0['ID']; $nMod = [int](To-Num (Get-Val $r0 'Módulos (n)' 1) 1); if ($nMod -lt 1) { $nMod = 1 }
    $ac0 = Get-Val $asigComp $id0 ([ordered]@{}); $devs0 = @(Get-Val $ac0 'desarrolladores' @())
    $nL = 1
    if ($nMod -gt 1 -and (Get-Val $r0 'Estrategia' 'Migrar') -notin @('Retirar', 'Mantener')) {
        $nL = [int](To-Num (Get-Val $ac0 'lotes' 0) 0)
        if ($nL -lt 1) { $nL = $(if ($devs0.Count -gt 1) { $devs0.Count } else { [int][Math]::Ceiling($nMod / (Par 'Módulos por lote')) }) }
        $nL = [Math]::Max(1, [Math]::Min($nL, $nMod))
    }
    $lotesDe[$id0] = New-Object System.Collections.Generic.List[string]
    for ($j = 1; $j -le $nL; $j++) {
        $f = [ordered]@{}; foreach ($kk in $r0.Keys) { $f[$kk] = $r0[$kk] }
        $f['_idBase'] = $id0; $f['_base'] = $r0; $f['_modulos'] = $nMod; $f['_lote'] = $j; $f['_nLotes'] = $nL
        if ($nL -gt 1) {
            $f['ID'] = "$id0-L$j"
            $m1 = 0; for ($q = 1; $q -lt $j; $q++) { $m1 += Split-Int $nMod $nL $q }
            $mJ = Split-Int $nMod $nL $j
            $f['Nombre'] = "$(Get-Val $r0 'Nombre' $id0) – lote $j/$nL (módulos $($m1 + 1)-$($m1 + $mJ) de $nMod)"
            foreach ($cp in $camposReparto) { $v = Get-Val $r0 $cp; if ($null -ne $v -and (To-Num $v 0) -gt 0) { $f[$cp] = Split-Int (To-Num $v) $nL $j } }
            if ($j -gt 1) { $f['Migración de datos'] = 'Ninguna'; $f['Requiere permisos y auditoría'] = 'No' }   # una sola vez por aplicativo (lote 1)
            $f['_dev'] = $(if ($devs0.Count -gt 0) { [string]$devs0[($j - 1) % $devs0.Count] } else { $null })
        }
        $filasInv.Add($f); $lotesDe[$id0].Add([string]$f['ID'])
    }
}

$ordenAuto = 0
foreach ($r in $filasInv) {
    $ordenAuto++
    $id = [string]$r['ID']; $idBase = [string]$r['_idBase']; $rb = $r['_base']; $nMod = [int]$r['_modulos']
    $est = [string](Get-Val $r 'Estrategia' 'Migrar')
    $loc = To-Num (Get-Val $r 'LOC' 0); $pant = To-Num (Get-Val $r 'Pantallas / formularios' 0); $rep = To-Num (Get-Val $r 'Reportes / salidas' 0)
    $int = To-Num (Get-Val $r 'Integraciones (n)' 0); $tab = To-Num (Get-Val $r 'Tablas / hojas / fuentes de datos (n)' 0)
    # Para la complejidad: tamaño del módulo promedio de la fila original (con 1 módulo = la fila tal cual)
    $locC = (To-Num (Get-Val $rb 'LOC' 0)) / $nMod; $pantC = (To-Num (Get-Val $rb 'Pantallas / formularios' 0)) / $nMod
    $intC = (To-Num (Get-Val $rb 'Integraciones (n)' 0)) / $nMod
    $migrC = [string](Get-Val $rb 'Migración de datos' 'Ninguna')
    $migr = [string](Get-Val $r 'Migración de datos' 'Ninguna'); $reglas = [string](Get-Val $r 'Reglas de negocio' 'Bajo')
    $riesgo = [string](Get-Val $r 'Riesgo técnico' 'Bajo'); $dep = [string](Get-Val $r 'Dependencia externa' 'Ninguna')
    $reut = [string](Get-Val $r 'Reutilización' 'Media')

    # Puntaje de complejidad (complejidad-priorizacion.md)
    $crit = [ordered]@{
        'Reglas de negocio' = Nivel3 $reglas @('Bajo', 'Medio', 'Alto')
        'Integraciones'     = $(if ($intC -le 0) { 1 } elseif ($intC -le 1) { 2 } else { 3 })
        'Datos'             = $(if ($migrC -eq 'Compleja') { 3 } elseif ($migrC -eq 'Simple') { 2 } else { 1 })
        'Pantallas'         = $(if ($pantC -le 0) { 1 } elseif ($pantC -le 3) { 2 } else { 3 })
        'Riesgo técnico'    = Nivel3 $riesgo @('Bajo', 'Medio', 'Alto')
        'Dependencia externa' = Nivel3 $dep @('Ninguna', 'Parcial', 'Crítica')
        'Reutilización'     = Nivel3 $reut @('Alta', 'Media', 'Baja')
        'Tamaño (LOC)'      = $(if ($locC -lt 1000) { 1 } elseif ($locC -le 3000) { 2 } else { 3 })
    }
    $suma = 0; $dom = ''; $max = 0
    foreach ($k in $crit.Keys) { $suma += $crit[$k]; if ($crit[$k] -gt $max) { $max = $crit[$k]; $dom = $k } }
    $puntaje = [Math]::Round($suma / 8.0, 2)
    $cx = if ($puntaje -le 1.6) { 'Baja' } elseif ($puntaje -le 2.3) { 'Media' } else { 'Alta' }
    $motivoCx = "Puntaje $puntaje"; if ($nMod -gt 1) { $motivoCx += " (por módulo promedio de $nMod)" }
    if ($riesgo -eq 'Alto' -and $dep -eq 'Crítica') { $cx = 'Alta'; $motivoCx += '; escalada: riesgo Alto y dependencia Crítica' }
    if ($locC -gt 5000 -or $intC -ge 4) {
        if ($cx -eq 'Baja') { $cx = 'Media'; $motivoCx += '; mínimo Media por tamaño/integraciones' }
        if ($reglas -eq 'Alto') { $cx = 'Alta'; $motivoCx += '; Alta por reglas Alto' }
    }
    $forz = Get-Val $r 'Complejidad forzada (opcional)'
    if ($forz) { $cx = [string]$forz; $motivoCx += "; forzada en la entrada: $forz" }

    $ac = Get-Val $asigComp $id $null; if ($null -eq $ac) { $ac = Get-Val $asigComp $idBase ([ordered]@{}) }   # un lote hereda la asignación de su fila
    $ola = Get-Val $ac 'ola'
    if (-not $ola) { $os = Get-Val $r 'Ola sugerida'; if ($os) { $ola = $(if ([string]$os -match '^\d+$') { "Ola $os" } else { [string]$os }) } }
    if ($modelo -eq 'Entrega única') { $ola = 'Ola 1' }
    elseif ($ola) {
        $nO = $(if ([string]$ola -match '\d+') { [int]$Matches[0] } else { 0 })
        if ($nO -lt 1 -or $nO -gt $nOlas) { $errores.Add("Componente ${id}: '$ola' fuera de las $nOlas olas pedidas.") } else { $ola = "Ola $nO" }
    }
    $tieneTareas = $est -ne 'Mantener'
    $dev = Get-Val $ac 'desarrollador'; if ($r['_dev'] -and -not (Get-Val $asigComp $id $null)) { $dev = $r['_dev'] }
    if (-not $dev) { $dl = @(Get-Val $ac 'desarrolladores' @()); if ($dl.Count -gt 0) { $dev = $dl[0] } }
    if ($tieneTareas -and -not $ola) { $errores.Add("Componente $id sin ola (asignacion.componentes.$id.ola).") }
    if ($tieneTareas -and -not $dev) { $errores.Add("Componente $id sin desarrollador (asignacion.componentes.$id.desarrollador).") }
    if ($dev -and -not $porNombre.ContainsKey([string]$dev)) { $errores.Add("Componente ${id}: el desarrollador '$dev' no está en Recursos.") }

    $componentes.Add([ordered]@{
            id = $id; idInventario = $idBase; lote = [int]$r['_lote']; lotes = [int]$r['_nLotes']; modulos = $nMod; nombre = [string](Get-Val $r 'Nombre' $id); aplicativoOrigen = [string](Get-Val $r 'Aplicativo origen' '')
            tecnologiaOrigen = [string](Get-Val $r 'Tecnología origen' ''); aplicativo = [string](Get-Val $r 'Aplicativo destino' '')
            destino = [string](Get-Val $r 'Destino existente o nuevo' 'Existente'); plataforma = [string](Get-Val $r 'Plataforma destino' '')
            estrategia = $est; tipo = [string](Get-Val $r 'Tipo de trabajo' $tipoProyecto); prioridad = [int](To-Num (Get-Val $r 'Prioridad de negocio (1 alta - 3 baja)' 2) 2)
            loc = $loc; pantallas = $pant; reportes = $rep; integraciones = $int; tablas = $tab; migracion = $migr; reglas = $reglas
            riesgo = $riesgo; dependencia = $dep; reutilizacion = $reut; permisos = (Is-Si (Get-Val $r 'Requiere permisos y auditoría' 'No'))
            faltante = [string](Get-Val $r 'Información faltante o por validar' ''); override = (Get-Val $r 'Horas override (opcional)')
            puntaje = $puntaje; criterios = $crit; criterioDominante = $dom; complejidad = $cx; motivoComplejidad = $motivoCx
            ola = [string]$ola; desarrollador = [string]$dev; revisor = [string](Get-Val $ac 'revisor' $LT)
            orden = [int](To-Num (Get-Val $ac 'orden' $ordenAuto) $ordenAuto); motivo = [string](Get-Val $ac 'motivo' '')
            componenteIA = [bool](Get-Val $ac 'componenteIA' $false); tieneTareas = $tieneTareas
            contingenciaPct = 0; motivoContingencia = ''
        })
}
if ($errores.Count -gt 0) {
    Write-JsonFile (Join-Path $Trabajo 'wbs_errores.json') $errores
    $errores | ForEach-Object { Write-Host " - $_" }
    throw "Hay $($errores.Count) errores de asignación. Corrija asignacion.json o la entrada."
}

$compT = @($componentes | Where-Object { $_.tieneTareas })
$compTrab = @($compT | Where-Object { $_.estrategia -ne 'Retirar' })
$nComp = $compTrab.Count
$apps = @($compTrab | ForEach-Object { $_.aplicativo } | Sort-Object -Unique)
$nApp = $apps.Count
$nPlat = @($compTrab | ForEach-Object { $_.plataforma } | Sort-Object -Unique).Count
$baseComun = @{}
foreach ($app in $apps) {
    $cs = @($compTrab | Where-Object { $_.aplicativo -eq $app })
    if ($cs.Count -ge 2 -or @($cs | Where-Object { $_.destino -eq 'Nuevo' }).Count -gt 0) { $baseComun[$app] = (@($cs | Where-Object { $_.destino -eq 'Nuevo' }).Count -gt 0) }
}

# ---------- Tareas ----------
$tareas = New-Object System.Collections.Generic.List[object]
$porClave = @{}
function New-Tarea([string]$clave, [string]$codigo, [string]$fase, [string]$nombre, [string]$recurso, [double]$base, [double]$factor,
    [string]$origen, [string]$driver, [string[]]$preds, $ctx, [string]$tipo = 'normal', [int]$minimo = 1) {
    $h = Get-HorasFinales $base $factor $minimo
    $obs = ''
    if ($redondeo4 -and $tipo -eq 'normal') { $h4 = [int]([Math]::Ceiling($h / 4.0) * 4); if ($h4 -ne $h) { $obs = "Redondeo a bloques de 4 h: +$($h4 - $h) h asumidas"; $h = $h4 } }
    $t = [ordered]@{
        clave = $clave; codigo = $codigo; fase = $FASES[$fase]; tarea = $nombre; recurso = $recurso
        horasBase = [Math]::Round($base, 2); factorIA = $factor; horas = $h; origen = $origen; driver = $driver
        preds = @($preds | Where-Object { $_ }); tipo = $tipo; hDia = 0; diasMin = 0; minDia = 0
        ola = $ctx.ola; aplicativoOrigen = $ctx.aplicativoOrigen; aplicativo = $ctx.aplicativo; plataforma = $ctx.plataforma
        componente = $ctx.componente; componenteNombre = $ctx.componenteNombre; complejidad = $ctx.complejidad; bloque = $ctx.bloque
        observacion = $obs
    }
    if (-not $recurso) { $errores.Add("Sin responsable para $clave ($nombre).") }
    $tareas.Add($t); $porClave[$clave] = $t
    return $t
}
$ctxTRV = @{ ola = 'Transversal'; aplicativoOrigen = 'Transversal'; aplicativo = 'Todos'; plataforma = 'Transversal'; componente = 'TRV'; componenteNombre = 'Transversal'; complejidad = ''; bloque = '' }

# Transversales (modelo-horas.md, tabla Transversales)
$JP = Get-Responsable 'Jefe de Proyecto'
$null = New-Tarea 'TRV-ENT' 'TRV-ENT' '01.1' 'Kick-off y entendimiento funcional/técnico del alcance, dependencias, integraciones y prioridades' $LT ([Math]::Min(40, 16 + 2 * $nComp)) 1 'Asumida' "16 + 2 × $nComp comp. (máx. 40)" @() $ctxTRV
$null = New-Tarea 'TRV-REL' 'TRV-REL' '01.2' 'Relevamiento técnico: dependencias, accesos, permisos, restricciones y confirmación de la estrategia de oleadas' $LT ([Math]::Min(40, 16 + 4 * $nApp)) 1 'Asumida' "16 + 4 × $nApp aplicativos (máx. 40)" @('TRV-ENT') $ctxTRV
$null = New-Tarea 'TRV-AMB' 'TRV-AMB' '01.2' 'Ambientes, accesos, repositorios y pipelines por aplicativo destino' (Get-Responsable 'Especialista de Plataformas') (8 * $nApp) 1 'Asumida' "8 × $nApp aplicativos" @('TRV-ENT') $ctxTRV
if ($usaIA) { $null = New-Tarea 'TRV-IA' 'TRV-IA' '01.2' 'Configurar el asistente IA por repositorio: instrucciones, prompts de migración y reglas de sanitización' (Get-Responsable 'Developer Sr') (4 * $nApp) 1 'Asumida' "4 × $nApp aplicativos" @('TRV-ENT') $ctxTRV }
$null = New-Tarea 'TRV-01' 'TRV-01' '02' 'Relevar usuarios, roles y permisos actuales de los componentes' (Get-Responsable 'Developer') (4 + 0.5 * $nComp) 1 'Paramétrica' "4 + 0,5 × $nComp" @('TRV-ENT') $ctxTRV
$null = New-Tarea 'TRV-02' 'TRV-02' '02' 'Relevar roles y perfiles de los aplicativos destino' (Get-Responsable 'Developer') (4 * $nApp) 1 'Paramétrica' "4 × $nApp" @('TRV-ENT') $ctxTRV
$null = New-Tarea 'TRV-06' 'TRV-06' '03' 'Arquitectura de referencia: patrones de integración, errores, configuración, secretos y observabilidad' (Get-Responsable 'Arquitecto de Aplicaciones') ($(if ($nPlat -ge 3) { 24 } else { 16 })) 1 'Paramétrica' "16 (+8 si ≥ 3 plataformas; hay $nPlat)" @('TRV-REL') $ctxTRV
$null = New-Tarea 'TRV-05' 'TRV-05' '03' 'Eventos de auditoría y trazabilidad' $LT 8 1 'Paramétrica' '8' @('TRV-REL') $ctxTRV
$null = New-Tarea 'TRV-07' 'TRV-07' '03' 'Modelo de tablas comunes de soporte (auditoría, parámetros, bitácora)' $LT 8 1 'Paramétrica' '8' @('TRV-REL') $ctxTRV
$null = New-Tarea 'TRV-04' 'TRV-04' '03' 'Matriz de roles por pantalla y operación' $LT (8 + 0.5 * $nComp) 1 'Paramétrica' "8 + 0,5 × $nComp" @('TRV-01', 'TRV-02') $ctxTRV
$nMigr = @($compTrab | Where-Object { $_.migracion -ne 'Ninguna' }).Count
if ($nMigr -gt 0) { $null = New-Tarea 'TRV-08' 'TRV-08' '03' 'Estrategia de carga inicial y tratamiento del histórico' (Get-Responsable 'Datos') (8 + 2 * $nMigr) 1 'Paramétrica' "8 + 2 × $nMigr comp. con migración" @('TRV-REL') $ctxTRV }
if ($qaProveedor) { $null = New-Tarea 'TRV-QA' 'TRV-QA' '03' 'Estrategia y plan de pruebas de certificación' (Get-Responsable 'Analista Calidad Sr') (8 + $nComp) 1 'Paramétrica' "8 + 1 × $nComp" @('TRV-ENT') $ctxTRV }
if ($sensibles) { $null = New-Tarea 'TRV-SEC' 'TRV-SEC' '03' 'Threat modeling y lineamientos de desarrollo seguro' (Get-Responsable 'Seguridad') 16 1 'Paramétrica' '16' @('TRV-REL') $ctxTRV }
$totPant = 0; foreach ($c in $compTrab) { $totPant += $c.pantallas }
if (@($compTrab | Where-Object { $_.destino -eq 'Nuevo' }).Count -gt 0 -or $totPant -ge 10) { $null = New-Tarea 'TRV-UX' 'TRV-UX' '03' 'Guía UX/UI y sistema de diseño' (Get-Responsable 'UX') 24 1 'Paramétrica' "24 (destino nuevo o $totPant pantallas)" @('TRV-REL') $ctxTRV }
$prereq04 = @('TRV-01', 'TRV-02', 'TRV-04', 'TRV-05', 'TRV-06', 'TRV-07')

# Base común por aplicativo destino
foreach ($app in $baseComun.Keys) {
    $cs = @($compTrab | Where-Object { $_.aplicativo -eq $app } | Sort-Object { [int]($_.ola -replace '\D', '') }, orden, lote)
    $c0 = $cs[0]
    $ctx = @{ ola = $c0.ola; aplicativoOrigen = 'Base común'; aplicativo = $app; plataforma = $c0.plataforma; componente = "Base común $app"; componenteNombre = "Base común $app"; complejidad = ''; bloque = '' }
    $dev = $c0.desarrollador; $k = "BC|$app|"
    $n03 = if ($baseComun[$app]) { 80 } else { 40 }
    $null = New-Tarea "${k}BC-01" 'BC-01' '02' "Analizar impacto sobre el aplicativo destino ($app)" $dev 8 (FIA 'Análisis') 'Paramétrica' '8' @('TRV-ENT') $ctx
    $null = New-Tarea "${k}BC-02" 'BC-02' '03' "Definir estructura común y habilitación de funciones por perfil y por ola ($app)" $dev 8 (FIA 'Diseño') 'Paramétrica' '8' @("${k}BC-01") $ctx
    $null = New-Tarea "${k}BC-03" 'BC-03' '04' "Construir estructura común: navegación, layout, configuración, logging ($app)" $dev $n03 (FIA 'Construcción') 'Paramétrica' "$n03" (@("${k}BC-02") + $prereq04) $ctx
    $null = New-Tarea "${k}BC-04" 'BC-04' '04' "Implementar modelo base de permisos y auditoría ($app)" $dev 24 (FIA 'Construcción') 'Paramétrica' '24' @("${k}BC-03") $ctx
    $null = New-Tarea "${k}BC-05" 'BC-05' '04' "Revisión de código de la base común ($app)" $LT 3 (FIA 'Revisión de código') 'Paramétrica' '3' @("${k}BC-04") $ctx
    $null = New-Tarea "${k}BC-06" 'BC-06' '05' "Pruebas unitarias de la base común ($app)" $dev 10 (FIA 'Ejecutar PU') 'Paramétrica' '10' @("${k}BC-05") $ctx
}

# Tareas por componente
$ix = @{ 'Baja' = 0; 'Media' = 1; 'Alta' = 2 }
$restr = @($ENT.restricciones)
$ultimoDeSerie = @{}   # desarrollador -> componente anterior en su serie
$compOrdenados = @($compT | Sort-Object { [int]($_.ola -replace '\D', '') }, orden, lote)
foreach ($c in $compOrdenados) {
    $id = $c.id; $k = "$id|"; $i = $ix[$c.complejidad]; if ($null -eq $i) { $i = 1 }
    $ctx = @{ ola = $c.ola; aplicativoOrigen = $c.aplicativoOrigen; aplicativo = $c.aplicativo; plataforma = $c.plataforma; componente = $id; componenteNombre = $c.nombre; complejidad = $c.complejidad; bloque = '' }
    $dev = $c.desarrollador; $alta = $c.complejidad -eq 'Alta'
    if ($c.estrategia -eq 'Retirar') {
        $null = New-Tarea "${k}AN-02" 'AN-02' '02' "Confirmar con negocio el retiro ($id)" $dev 2 (FIA 'Otros') 'Paramétrica' 'Retiro' @('TRV-ENT') $ctx
        $c['tareas'] = @("${k}AN-02"); continue
    }
    $chain = New-Object System.Collections.Generic.List[string]
    $prev = $ultimoDeSerie[$dev]
    $fcx = Par "Factor complejidad $($c.complejidad)"
    $tipo = $c.tipo
    # Análisis
    $an01 = (@(4, 8, 16)[$i]) + [Math]::Max(0, [Math]::Floor(($c.loc - 3000) / 1000))
    $an01Drv = "$($c.complejidad) + 1 h c/1.000 LOC > 3.000"
    if ($tipo -eq 'Mejora evolutiva') { $an01 = $an01 / 2; $an01Drv += ' ÷ 2 (evolutiva)' }
    $n01 = if ($tipo -eq 'Upgrade') { "Inventario de dependencias incompatibles ($id)" } else { "Análisis del componente origen: extraer/sanitizar código, revisar lógica, objetos ocultos, integraciones y datos ($id)" }
    $predAN = @('TRV-ENT'); if ($prev) { $predAN += $prev.ultimoCO }
    $null = New-Tarea "${k}AN-01" 'AN-01' '02' $n01 $dev $an01 (FIA 'Análisis') 'Paramétrica' $an01Drv $predAN $ctx; $chain.Add("${k}AN-01")
    $null = New-Tarea "${k}AN-02" 'AN-02' '02' "Validar con negocio las reglas existentes ($id)" $(if ($alta) { $LT } else { $dev }) (@(2, 3, 6)[$i]) (FIA 'Otros') 'Paramétrica' $c.complejidad @($chain[-1]) $ctx; $chain.Add("${k}AN-02")
    $tieneBC = $baseComun.ContainsKey($c.aplicativo)
    if ($c.destino -eq 'Existente' -and -not $tieneBC) { $null = New-Tarea "${k}AN-03" 'AN-03' '02' "Analizar el aplicativo destino y el impacto ($id)" $dev (@(2, 4, 8)[$i]) (FIA 'Análisis') 'Paramétrica' $c.complejidad @($chain[-1]) $ctx; $chain.Add("${k}AN-03") }
    if ($c.migracion -ne 'Ninguna') { $null = New-Tarea "${k}AN-04" 'AN-04' '02' "Análisis de datos: fuentes, calidad, volumen e histórico ($id)" $dev ($(if ($c.migracion -eq 'Compleja') { 12 } else { 4 })) (FIA 'Análisis') 'Paramétrica' $c.migracion @($chain[-1]) $ctx; $chain.Add("${k}AN-04") }
    # Diseño
    $null = New-Tarea "${k}DI-01" 'DI-01' '03' "Diseñar la solución técnica ($id)" $(if ($alta) { $LT } else { $dev }) (@(3, 8, 16)[$i]) (FIA 'Diseño') 'Paramétrica' $c.complejidad @($chain[-1]) $ctx; $chain.Add("${k}DI-01")
    if ($c.pantallas -gt 0) { $null = New-Tarea "${k}DI-02" 'DI-02' '03' "Diseñar prototipos de pantallas en el destino ($id)" $dev ($c.pantallas * @(1, 2, 3)[$i]) (FIA 'Diseño') 'Paramétrica' "$($c.pantallas) pant. × $(@(1,2,3)[$i])" @($chain[-1]) $ctx; $chain.Add("${k}DI-02") }
    if ($c.tablas -gt 0) { $b = $c.tablas * @(0.5, 1, 2)[$i] + $(if ($c.migracion -eq 'Compleja') { 4 } else { 0 }); $null = New-Tarea "${k}DI-03" 'DI-03' '03' "Diseñar modelo de datos, mapeo y carga ($id)" $dev $b (FIA 'Diseño') 'Paramétrica' "$($c.tablas) tablas" @($chain[-1]) $ctx; $chain.Add("${k}DI-03") }
    # Construcción
    $pre04 = @($chain[-1]) + $prereq04
    if ($tieneBC) { $pre04 += "BC|$($c.aplicativo)|BC-04" }
    if ($prev) { $pre04 += $prev.ultimoPU }
    $minDia = ''
    foreach ($rs in $restr) {
        $afecta = [string](Get-Val $rs 'Afecta a (IDs de Inventario)' '')
        if ((Get-Val $rs 'Tipo') -eq 'Dependencia externa' -and (($afecta -split '[,;\s]+') -contains $id -or ($afecta -split '[,;\s]+') -contains $c.idInventario) -and (Get-Val $rs 'Fecha')) { $minDia = [string]$rs['Fecha'] }
    }
    $sumCO26 = 0.0; $sumCO19 = 0.0; $first04 = $true; $iaC = FIA 'Construcción'
    function Add-CO([string]$cod, [string]$nom, [double]$base, [string]$drv) {
        if ($base -le 0) { return }
        $preds = if ($script:first04) { $script:pre04 } else { @($script:chain[-1]) }
        $t = New-Tarea "$($script:k)$cod" $cod '04' "$nom ($($script:id))" $script:dev $base $script:iaC 'Paramétrica' $drv $preds $script:ctx
        if ($script:first04 -and $script:minDia) { $t.minFecha = $script:minDia }
        $script:first04 = $false; $script:chain.Add("$($script:k)$cod")
        $script:sumCO19 += $base
        if ($cod -in @('CO-02', 'CO-03', 'CO-04', 'CO-05', 'CO-06')) { $script:sumCO26 += $base }
    }
    if (-not $tieneBC) { Add-CO 'CO-01' 'Preparar el aplicativo destino para recibir la funcionalidad' (@(8, 16, 24)[$i]) $c.complejidad }
    $prod = Par 'Productividad LOC/h (sin IA)'
    $nCO2 = if ($tipo -eq 'Upgrade') { 'Adecuación de código (LOC afectadas)' } else { 'Implementar la lógica de negocio, cálculos y procesos' }
    Add-CO 'CO-02' $nCO2 ($c.loc / $prod * $fcx) "$($c.loc) LOC / $prod × $fcx"
    Add-CO 'CO-03' 'Implementar pantallas y navegación' ($c.pantallas * (Par "Horas por pantalla $($c.complejidad)")) "$($c.pantallas) pant."
    Add-CO 'CO-04' 'Convertir hojas de datos y parametría en tablas o configuración' ($c.tablas * (Par "Horas por tabla $($c.complejidad)")) "$($c.tablas) tablas"
    Add-CO 'CO-05' 'Implementar integraciones y sustituir automatizaciones de Office' ($c.integraciones * (Par "Horas por integración $($c.complejidad)")) "$($c.integraciones) integ."
    Add-CO 'CO-06' 'Implementar reportes y salidas' ($c.reportes * (Par "Horas por reporte $($c.complejidad)")) "$($c.reportes) reportes"
    Add-CO 'CO-07' 'Manejo explícito de errores y externalización de rutas y secretos' ([Math]::Max(1, (Par '% Errores y secretos') / 100 * $sumCO26)) "% errores × Σ CO-02..06"
    if ($c.permisos) { Add-CO 'CO-08' 'Implementar permisos y auditoría' (@(4, 8, 12)[$i]) $c.complejidad }
    if ($c.migracion -ne 'Ninguna') { Add-CO 'CO-09' 'Scripts de migración/carga de datos e histórico' ($(if ($c.migracion -eq 'Compleja') { 24 } else { 8 })) $c.migracion }
    $ultimoCOdev = $chain[-1]
    # Contingencia: se calcula al final (necesita Σ 02-05); se reserva el lugar en la cadena.
    $kCT = "${k}CO-CT"
    $revisor = if ($c.revisor -and $c.revisor -ne $dev) { $c.revisor } elseif ($dev -ne $LT) { $LT } else { Get-Responsable 'Arquitecto de Aplicaciones' }
    $null = New-Tarea "${k}CO-10" 'CO-10' '04' "Revisar código y ejecutar análisis estático ($id)" $revisor ([Math]::Max(1, (Par '% Revisión de código') / 100 * $sumCO19)) (FIA 'Revisión de código') 'Paramétrica' '% revisión × Σ CO' @($kCT) $ctx
    $null = New-Tarea "${k}PU-01" 'PU-01' '05' "Preparar casos y datos de prueba unitaria contra resultados del legacy ($id)" $dev ([Math]::Max(2, (Par '% Preparar casos PU') / 100 * $sumCO19)) (FIA 'Preparar casos PU') 'Paramétrica' '% prep. PU × Σ CO' @($kCT) $ctx 'normal' 2
    $bPU2 = [Math]::Max(2, (Par '% Ejecutar PU') / 100 * $sumCO19); if ($tipo -eq 'Upgrade') { $bPU2 = $bPU2 * 1.5 }
    $prePU2 = @("${k}CO-10", "${k}PU-01"); if ($tieneBC) { $prePU2 += "BC|$($c.aplicativo)|BC-06" }
    $null = New-Tarea "${k}PU-02" 'PU-02' '05' "Implementar y ejecutar pruebas unitarias y corregir ($id)" $dev $bPU2 (FIA 'Ejecutar PU') 'Paramétrica' $(if ($tipo -eq 'Upgrade') { '% ejec. PU × Σ CO × 1,5 (Upgrade)' } else { '% ejec. PU × Σ CO' }) $prePU2 $ctx 'normal' 2
    if ($alcanceProd) { $null = New-Tarea "${k}DE-01" 'DE-01' '07' "Manual de usuario y runbook del componente ($id)" $dev (@(2, 4, 8)[$i]) (FIA 'Manuales') 'Paramétrica' $c.complejidad @("${k}PU-02") $ctx }

    # Contingencia del componente
    $pcts = New-Object System.Collections.Generic.List[object]
    if ($c.reutilizacion -eq 'Alta' -and $c.riesgo -eq 'Bajo') { $pcts.Add(@((Par 'Contingencia conocido'), 'trabajo conocido')) }
    if ($c.destino -eq 'Existente' -or $c.integraciones -gt 0) { $pcts.Add(@((Par 'Contingencia integración'), 'integración con sistema existente')) }
    if ($c.faltante -or $c.reglas -eq 'Alto' -or $claridad -eq 'Baja') { $pcts.Add(@((Par 'Contingencia requerimientos poco claros'), 'requerimientos poco claros')) }
    if ($c.dependencia -eq 'Parcial') { $pcts.Add(@(10, 'dependencia externa parcial')) }
    if ($c.dependencia -eq 'Crítica') { $pcts.Add(@((Par 'Contingencia dependencia externa'), 'dependencia externa crítica')) }
    if ($c.componenteIA) { $pcts.Add(@((Par 'Contingencia componente IA'), 'componente de IA nuevo')) }
    if ($pcts.Count -eq 0) { $pcts.Add(@((Par 'Contingencia conocido'), 'trabajo conocido')) }
    $best = $pcts | Sort-Object { $_[0] } -Descending | Select-Object -First 1
    $pct = [Math]::Min(20, [double]$best[0])
    $c.contingenciaPct = $pct; $c.motivoContingencia = [string]$best[1]
    $c['kCT'] = $kCT; $c['ultimoCOdev'] = $ultimoCOdev; $c['ctx'] = $ctx
    $ultimoDeSerie[$dev] = @{ ultimoCO = $ultimoCOdev; ultimoPU = "${k}PU-02" }
    $c['tareas'] = @($tareas | Where-Object { $_.componente -eq $id } | ForEach-Object { $_.clave })
}

# ---------- Ajustes manuales de horas / recurso (ajustes.json) ----------
$ajT = Get-Val $AJUSTES 'tareas' ([ordered]@{})
foreach ($kAj in $ajT.Keys) {
    if (-not $porClave.ContainsKey($kAj)) { Write-Host " - Aviso: el ajuste '$kAj' no corresponde a ninguna tarea actual; se ignora."; continue }
    $t = $porClave[$kAj]; $a = $ajT[$kAj]
    if ($null -ne (Get-Val $a 'horas')) { $t.horas = [int](To-Num $a.horas); $t.origen = 'Entrada'; $t.observacion = 'Ajuste manual: ' + [string](Get-Val $a 'motivo' 'horas definidas por el usuario') }
    if (Get-Val $a 'recurso') { $t.recurso = [string]$a.recurso; $t.observacion = (($t.observacion + ' Recurso definido por el usuario.').Trim()) }
}

# Horas override del Inventario y filas de contingencia (necesitan las horas finales 02-05)
foreach ($c in @($compTrab)) {
    $ts = @($tareas | Where-Object { $_.componente -eq $c.id -and $_.fase -match '^0[2-5] ' })
    if ($null -ne $c.override -and (To-Num $c.override) -gt 0) {
        $obj = [int](To-Num $c.override); $tot = 0; foreach ($t in $ts) { $tot += $t.horas }
        $acum = 0
        for ($j = 0; $j -lt $ts.Count; $j++) {
            $nuevo = if ($j -eq $ts.Count - 1) { $obj - $acum } else { [int][Math]::Max(1, [Math]::Round($ts[$j].horas * $obj / [double]$tot)) }
            if ($nuevo -lt 1) { $nuevo = 1 }
            $acum += $nuevo; $ts[$j].horas = $nuevo; $ts[$j].origen = 'Entrada'; $ts[$j].observacion = "Horas override del Inventario ($obj h) repartidas proporcionalmente"
        }
    }
    $s = 0; foreach ($t in $ts) { $s += $t.horas }
    $hCT = Ceil-Int ($c.contingenciaPct / 100.0 * $s)
    if ($hCT -gt 0) {
        $null = New-Tarea $c.kCT 'CO-CT' '04' "Contingencia ($($c.contingenciaPct) %) – $($c.motivoContingencia) ($($c.id))" $c.desarrollador $hCT 1 'Contingencia' "$($c.contingenciaPct) % × $s h (02-05)" @($c.ultimoCOdev) $c.ctx
        $porClave[$c.kCT].horas = $hCT; $porClave[$c.kCT].horasBase = $hCT
    }
    else {
        # Sin contingencia: las sucesoras dependen del último CO del desarrollador.
        foreach ($t in $tareas) { $t.preds = @($t.preds | ForEach-Object { if ($_ -eq $c.kCT) { $c.ultimoCOdev } else { $_ } }) }
    }
}

# ---------- Bloques de certificación / despliegue / seguimiento ----------
$bloques = New-Object System.Collections.Generic.List[object]
$olas = @($compT | ForEach-Object { $_.ola } | Sort-Object { [int]($_ -replace '\D', '') } -Unique)
if ($olas.Count -lt $nOlas) { Write-Warning "Se pidieron $nOlas olas pero solo $($olas.Count) tienen componentes." }
foreach ($ola in $olas) {
    $cs = @($compT | Where-Object { $_.ola -eq $ola })
    if ($agrupacion -eq 'Por aplicativo destino') {
        foreach ($app in @($cs | ForEach-Object { $_.aplicativo } | Sort-Object -Unique)) { $bloques.Add([ordered]@{ nombre = "$ola · $app"; ola = $ola; componentes = @($cs | Where-Object { $_.aplicativo -eq $app } | ForEach-Object { $_.id }) }) }
    }
    else { $bloques.Add([ordered]@{ nombre = $ola; ola = $ola; componentes = @($cs | ForEach-Object { $_.id }) }) }
}
$hEjecQA = Par 'Ejecución QA (h/día)'; $hSop = Par 'Soporte dev en certificación (h/día)'; $hCoord = Par 'Coordinación LT en certificación (h/día)'
$QAssr = Get-Responsable 'Analista Calidad Ssr'
$certClaves = New-Object System.Collections.Generic.List[string]
foreach ($b in $bloques) {
    $cs = @($componentes | Where-Object { $b.componentes -contains $_.id })
    $csT = @($cs | Where-Object { $_.estrategia -ne 'Retirar' })
    $h04 = 0.0; foreach ($t in $tareas) { if ($b.componentes -contains $t.componente -and $t.fase -like '04*' -and $t.codigo -ne 'CO-CT') { $h04 += $t.horasBase } }
    $devs = @($csT | ForEach-Object { $_.desarrollador } | Sort-Object -Unique)
    $appsB = @($csT | ForEach-Object { $_.aplicativo } | Sort-Object -Unique)
    $b['devs'] = $devs; $b['aplicativos'] = $appsB; $b['h04'] = [Math]::Round($h04, 1)
    $b['diasCert'] = [Math]::Max($diasCertMin, (Ceil-Int (0.2 * $h04 / $hEjecQA)))
    $ctx = @{ ola = $b.ola; aplicativoOrigen = 'Varios'; aplicativo = ($appsB -join ', '); plataforma = (@($csT | ForEach-Object { $_.plataforma } | Sort-Object -Unique) -join ', '); componente = "Bloque $($b.nombre)"; componenteNombre = "Bloque $($b.nombre)"; complejidad = ''; bloque = $b.nombre }
    $pu = @($csT | ForEach-Object { "$($_.id)|PU-02" }); $co10 = @($csT | ForEach-Object { "$($_.id)|CO-10" })
    if ($csT.Count -eq 0) { $b['certClaves'] = @(); continue }
    $kb = "$($b.nombre)|"; $preCert = @($pu)
    $b['certClaves'] = New-Object System.Collections.Generic.List[string]
    if ($qaProveedor) {
        $preCE01 = @($co10); if ($porClave.ContainsKey('TRV-QA')) { $preCE01 += 'TRV-QA' }
        $null = New-Tarea "${kb}CE-01" 'CE-01' '06' "Diseño de casos de certificación ($($b.nombre))" $QAssr ([Math]::Max(4, 0.1 * $h04)) (FIA 'Otros') 'Paramétrica' '10 % × H04 (mín. 4)' $preCE01 $ctx
        $preCert += "${kb}CE-01"
        $t = New-Tarea "${kb}CE-02" 'CE-02' '06' "Ejecución de certificación, registro de defectos y evidencias ($($b.nombre))" $QAssr ($b.diasCert * $hEjecQA) 1 'Paramétrica' "$($b.diasCert) días × $hEjecQA h" $preCert $ctx 'tasa_fija'
        $t.hDia = $hEjecQA; $t.diasMin = $b.diasCert; $b.certClaves.Add($t.clave)
    }
    foreach ($d in $devs) {
        $t = New-Tarea "${kb}CE-03|$d" 'CE-03' '06' "Soporte en certificación y corrección de defectos ($($b.nombre))" $d ($b.diasCert * $hSop) 1 'Paramétrica' "$($b.diasCert) días × $hSop h" $preCert $ctx 'tasa_fija'
        $t.hDia = $hSop; $t.diasMin = $b.diasCert; $b.certClaves.Add($t.clave)
    }
    $t = New-Tarea "${kb}CE-04" 'CE-04' '06' "Coordinación de certificación, congelamiento y gestión de defectos críticos ($($b.nombre))" $LT ($b.diasCert * $hCoord) 1 'Paramétrica' "$($b.diasCert) días × $hCoord h" $preCert $ctx 'tasa_fija'
    $t.hDia = $hCoord; $t.diasMin = $b.diasCert; $b.certClaves.Add($t.clave)
    if ($sensibles) { $t = New-Tarea "${kb}CE-05" 'CE-05' '06' "Pruebas de seguridad y validación de remediaciones ($($b.nombre))" (Get-Responsable 'Seguridad') (8 + 2 * $csT.Count) 1 'Paramétrica' "8 + 2 × $($csT.Count) comp." $preCert $ctx; $b.certClaves.Add($t.clave) }
    foreach ($x in $b.certClaves) { $certClaves.Add($x) }
}

# Despliegue y seguimiento: por bloque, o uno solo al final si es Pase único.
$pases = New-Object System.Collections.Generic.List[object]
if ($agrupacion -eq 'Pase único') { $pases.Add([ordered]@{ nombre = 'Pase único'; bloques = @($bloques | ForEach-Object { $_.nombre }) }) }
else { foreach ($b in $bloques) { $pases.Add([ordered]@{ nombre = $b.nombre; bloques = @($b.nombre) }) } }
$plat = Get-Responsable 'Especialista de Plataformas'
foreach ($pa in $pases) {
    $bs = @($bloques | Where-Object { $pa.bloques -contains $_.nombre })
    $devs = @($bs | ForEach-Object { $_.devs } | Sort-Object -Unique); $appsP = @($bs | ForEach-Object { $_.aplicativos } | Sort-Object -Unique)
    $ids = @($bs | ForEach-Object { $_.componentes })
    $pre = @($bs | ForEach-Object { $_.certClaves })
    if ($pre.Count -eq 0) { continue }
    $ctx = @{ ola = $(if ($pa.nombre -eq 'Pase único') { 'Pase único' } else { $bs[0].ola }); aplicativoOrigen = 'Varios'; aplicativo = ($appsP -join ', '); plataforma = ''; componente = "Bloque $($pa.nombre)"; componenteNombre = "Bloque $($pa.nombre)"; complejidad = ''; bloque = $pa.nombre }
    $kp = "$($pa.nombre)|"; $finPase = $pre
    if ($alcanceProd) {
        $null = New-Tarea "${kp}DE-02" 'DE-02' '07' "Plan de pase y comunicación ($($pa.nombre))" $JP 8 1 'Paramétrica' '8 por pase' $pre $ctx
        $null = New-Tarea "${kp}DE-03" 'DE-03' '07' "Preparar y ensayar paquetes, scripts DDL y carga inicial ($($pa.nombre))" $plat (8 * $appsP.Count) 1 'Paramétrica' "8 × $($appsP.Count) aplicativos" @("${kp}DE-02") $ctx
        $t = New-Tarea "${kp}DE-04" 'DE-04' '07' "Ejecutar el pase y validar post-despliegue ($($pa.nombre))" $plat (4 * $appsP.Count) 1 'Paramétrica' "4 × $($appsP.Count) aplicativos" @("${kp}DE-03") $ctx
        $t['esPase'] = $true; $finPase = @("${kp}DE-04")
        foreach ($d in $devs) { $t = New-Tarea "${kp}DE-04|$d" 'DE-04' '07' "Validación post-despliegue del desarrollador ($($pa.nombre))" $d 2 1 'Paramétrica' '2 por desarrollador' @("${kp}DE-03") $ctx; $t['esPase'] = $true; $finPase += $t.clave }
    }
    if ($alcanceHyper) {
        $preSE = $finPase; $se = @()
        foreach ($app in $appsP) {
            $dApp = @($compTrab | Where-Object { $_.aplicativo -eq $app -and $ids -contains $_.id } | ForEach-Object { $_.desarrollador })[0]
            $hd = Par 'Hypercare dev (h/día por aplicativo)'
            $t = New-Tarea "${kp}SE-02|$app" 'SE-02' '08' "Hypercare: estabilización, monitoreo y soporte inicial ($app)" $dApp ($diasHyper * $hd) 1 'Paramétrica' "$diasHyper días × $hd h" $preSE $ctx 'tasa_fija'
            $t.hDia = $hd; $t.diasMin = $diasHyper; $se += $t.clave
        }
        $hl = Par 'Hypercare LT (h/día)'
        $t = New-Tarea "${kp}SE-02|LT" 'SE-02' '08' "Hypercare: coordinación técnica ($($pa.nombre))" $LT ($diasHyper * $hl) 1 'Paramétrica' "$diasHyper días × $hl h" $preSE $ctx 'tasa_fija'
        $t.hDia = $hl; $t.diasMin = $diasHyper; $se += $t.clave
        $null = New-Tarea "${kp}SE-03" 'SE-03' '08' "Transferencia a soporte y mantenimiento ($($pa.nombre))" $LT (8 + 2 * $appsP.Count) 1 'Paramétrica' "8 + 2 × $($appsP.Count)" $se $ctx
        if ($retiroLegacy) {
            foreach ($c in @($compT | Where-Object { $ids -contains $_.id })) {
                $i = $ix[$c.complejidad]; if ($null -eq $i) { $i = 1 }
                $null = New-Tarea "$($c.id)|SE-01" 'SE-01' '08' "Retiro del componente legacy: desactivar, archivar y comunicar ($($c.id))" $c.desarrollador (@(1, 2, 4)[$i]) (FIA 'Otros') 'Paramétrica' $c.complejidad $se $c.ctx
            }
        }
    }
}

# Orden obligatorio (A antes que B): el análisis de B espera el PU-02 de A.
foreach ($rs in $restr) {
    if ((Get-Val $rs 'Tipo') -ne 'Orden obligatorio') { continue }
    $idsR = @(([string](Get-Val $rs 'Afecta a (IDs de Inventario)' '')) -split '[,;\s>]+' | Where-Object { $_ })
    for ($j = 1; $j -lt $idsR.Count; $j++) {
        $idsA = $(if ($lotesDe.ContainsKey($idsR[$j - 1])) { $lotesDe[$idsR[$j - 1]] } else { @($idsR[$j - 1]) })
        $idsB = $(if ($lotesDe.ContainsKey($idsR[$j])) { $lotesDe[$idsR[$j]] } else { @($idsR[$j]) })
        foreach ($b in $idsB) {
            $tB = $porClave["$b|AN-01"]; if (-not $tB) { continue }
            foreach ($a in $idsA) { if ($porClave.ContainsKey("$a|PU-02")) { $tB.preds = @($tB.preds) + "$a|PU-02" } }
        }
    }
}

# Gestión (TRV-GES): % sobre las demás horas; el mínimo de 1 h/día lo completa 3-Programar.ps1.
$sumOtras = 0; foreach ($t in $tareas) { $sumOtras += $t.horas }
$null = New-Tarea 'TRV-GES' 'TRV-GES' '00' 'Gestión del proyecto, backlog, seguimiento, riesgos, reportes y coordinación' $JP ((Par '% Gestión del proyecto') / 100 * $sumOtras) 1 'Paramétrica' "% gestión × $sumOtras h" @() $ctxTRV 'distribuida'
if (Is-Si (ParTxt 'Acompañamiento LT en días libres' 'No')) { $t = New-Tarea 'TRV-ACO' 'TRV-ACO' '00' 'Acompañamiento técnico en días sin otra tarea' $LT 0 1 'Asumida' '1 h por día libre del LT' @() $ctxTRV 'relleno' 0; $t.horas = 0 }

if ($errores.Count -gt 0) { $errores | ForEach-Object { Write-Host " - $_" }; throw "Hay $($errores.Count) errores en el WBS." }

# Horas base y finales por componente
foreach ($c in $componentes) {
    $hb = 0.0; $hf = 0; foreach ($t in $tareas) { if ($t.componente -eq $c.id) { $hb += $t.horasBase; $hf += $t.horas } }
    $c['horasBase'] = [Math]::Round($hb, 1); $c['horasFinales'] = $hf
    foreach ($x in @('kCT', 'ultimoCOdev', 'ctx')) { if ($c.Contains($x)) { $c.Remove($x) } }
}

$wbs = [ordered]@{
    generado    = (Get-Date).ToString('yyyy-MM-dd HH:mm')
    proyecto    = [ordered]@{
        nombre = [string](Get-Val $PROY 'Nombre del proyecto' ''); cliente = [string](Get-Val $PROY 'Cliente / área solicitante' ''); tipo = $tipoProyecto
        alcance = $alcance; modelo = $modelo; numeroOlas = $nOlas; agrupacion = $agrupacion; certificacion = $certA; sensibles = $sensibles; claridad = $claridad
        fechaInicio = (Get-Val $PROY 'Fecha tentativa de inicio'); fechaObjetivo = (Get-Val $PROY 'Fecha objetivo (comprometida o deseada)')
        jornada = $jornada; sabado = ((Get-Val $PROY 'Días laborables' '') -eq 'Lunes a sábado'); usaIA = $usaIA; version = [string](Get-Val $PROY 'Versión de la estimación' 'v1')
        diasCertMin = $diasCertMin; diasHyper = $diasHyper
        tarifas = (Is-Si (Get-Val $PROY 'Tarifas por rol disponibles (costeo)' 'No')); moneda = [string](Get-Val $PROY 'Moneda' '')
        fueraAlcance = [string](Get-Val $PROY 'Fuera de alcance' ''); alcanceResumen = [string](Get-Val $PROY 'Alcance (resumen)' ''); objetivo = [string](Get-Val $PROY 'Objetivo' '')
        ambientes = [string](Get-Val $PROY 'Ambientes disponibles' '')
    }
    parametros  = $ENT.parametros
    lider       = $LT
    recursos    = $recursos
    componentes = $componentes
    bloques     = $bloques
    pases       = $pases
    tareas      = $tareas
}
Write-JsonFile (Join-Path $Trabajo 'wbs.json') $wbs
$tot = 0; foreach ($t in $tareas) { $tot += $t.horas }
Write-Host "WBS: $($tareas.Count) tareas, $nComp componentes con trabajo, $($bloques.Count) bloques, $tot h (antes del mínimo de gestión)."
