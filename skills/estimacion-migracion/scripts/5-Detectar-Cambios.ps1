<#
.SYNOPSIS
  Compara el Excel de estimación actual con el snapshot de la última generación y lista lo que el usuario
  cambió a mano. Con -Aplicar, guarda esos cambios en ajustes.json para que sobrevivan a la regeneración.
.DESCRIPTION
  Se conserva (con -Aplicar):
    - Planning: horas por día y recurso de una tarea  -> ajustes.fijadas (la tarea queda fijada tal cual)
    - Detalle de horas: Horas finales de una tarea      -> ajustes.tareas[clave].horas (se reprograma)
    - Textos editables (motivos, criterios, racional, justificaciones, observaciones) -> ajustes.textos
    - Filas nuevas o modificadas de Supuestos y Riesgos -> ajustes.filasManuales
    - Hojas propias del usuario: 6-Generar-Excel.ps1 las copia siempre, no requieren ajuste.
  No se conserva: filas nuevas agregadas en Planning, formatos y cambios en celdas calculadas. Se informan.
.OUTPUTS
  _trabajo/cambios_manuales.json. Código de salida: 0 = sin cambios, 1 = hay cambios, 2 = no hay snapshot.
.EXAMPLE
  .\5-Detectar-Cambios.ps1 -Trabajo 'Estimaciones\MigraMacro\_trabajo' -Archivo 'Estimaciones\MigraMacro\estimacion_MigraMacro.xlsx'
  .\5-Detectar-Cambios.ps1 -Trabajo ... -Archivo ... -Aplicar
#>
param(
    [Parameter(Mandatory = $true)][string]$Trabajo,
    [Parameter(Mandatory = $true)][string]$Archivo,
    [switch]$Aplicar
)
. (Join-Path $PSScriptRoot 'Comun.ps1')

$Archivo = [IO.Path]::GetFullPath($Archivo)
$salidaJson = Join-Path $Trabajo 'cambios_manuales.json'
$SNAP = Read-JsonFile (Join-Path $Trabajo 'snapshot.json')
$hash = Get-FileHashSafe $Archivo
if ($null -eq $hash) { throw "No existe $Archivo" }
if ($null -eq $SNAP) {
    Write-JsonFile $salidaJson ([ordered]@{ archivo = $Archivo; hash = $hash; sinSnapshot = $true })
    Write-Host 'No hay snapshot.json: el archivo no fue generado por 6-Generar-Excel.ps1 (o se borró _trabajo).'
    exit 2
}
if ($hash -eq $SNAP.hash) {
    Write-JsonFile $salidaJson ([ordered]@{ archivo = $Archivo; hash = $hash; sinCambios = $true })
    Write-Host 'Sin cambios manuales: el archivo es idéntico al generado.'
    exit 0
}

$cambios = [ordered]@{
    archivo = $Archivo; hash = $hash; revisado = (Get-Date).ToString('yyyy-MM-dd HH:mm')
    planning = New-Object System.Collections.Generic.List[object]     # tareas con horas por día o recurso cambiados
    horas = New-Object System.Collections.Generic.List[object]        # Horas finales cambiadas en Detalle de horas
    textos = New-Object System.Collections.Generic.List[object]
    filas = New-Object System.Collections.Generic.List[object]
    noConservables = New-Object System.Collections.Generic.List[string]
    hojasPropias = New-Object System.Collections.Generic.List[string]
}
function Same($a, $b) { return ([string]$a).Trim() -eq ([string]$b).Trim() }
function DiasTexto($dias) { return (($dias.Keys | Sort-Object { [int]$_ } | ForEach-Object { "d$($_)=$($dias[$_])" }) -join ' ') }

$xl = Open-Excel
try {
    $wb = Invoke-Com { $xl.Workbooks.Open($Archivo, 0, $true) }
    $nombres = @(); foreach ($w in $wb.Worksheets) { $nombres += $w.Name }
    foreach ($h in $SNAP.hojasEstandar) { if ($nombres -notcontains $h) { $cambios.noConservables.Add("Se eliminó o renombró la hoja '$h'.") } }
    foreach ($n in $nombres) { if ($SNAP.hojasEstandar -notcontains $n) { $cambios.hojasPropias.Add($n) } }

    # Planning
    $hojaP = [string]$SNAP.hojasEstandar[3]
    if ($nombres -contains $hojaP) {
        $act = Read-PlanningTareas $wb.Worksheets.Item($hojaP)
        foreach ($id in $SNAP.planning.Keys) {
            $antes = $SNAP.planning[$id]; $clave = [string]$SNAP.claves[$id]
            if (-not $act.tareas.Contains($id)) { $cambios.noConservables.Add("Planning: se eliminó la fila $id ($clave)."); continue }
            $ahora = $act.tareas[$id]
            $difDias = (DiasTexto $antes.dias) -ne (DiasTexto $ahora.dias)
            $difRec = -not (Same $antes.recurso $ahora.recurso)
            if ($difDias -or $difRec) {
                $cambios.planning.Add([ordered]@{ id = $id; clave = $clave; tarea = $ahora.tarea; diasCambiados = $difDias; recursoAntes = $antes.recurso; recursoAhora = $ahora.recurso
                        diasAntes = (DiasTexto $antes.dias); diasAhora = (DiasTexto $ahora.dias); dias = $ahora.dias })
            }
        }
        foreach ($id in $act.tareas.Keys) { if (-not $SNAP.planning.Contains($id)) { $cambios.noConservables.Add("Planning: fila nueva con ID $id (no existe en el WBS; agréguela como componente o tarea en la entrada).") } }
        foreach ($f in $act.filasSinId) { $cambios.noConservables.Add("Planning: fila $f sin ID de tarea (no se conserva; si es trabajo real, agréguelo a la entrada).") }
    }

    # Tablas con textos editables (y Horas finales en Detalle de horas)
    $te = Get-TablasEditables
    foreach ($h in $te.Keys) {
        if ($nombres -notcontains $h -or -not $SNAP.tablas.Contains($h)) { continue }
        $act = Read-TablaHasta $wb.Worksheets.Item($h) $te[$h].filaEnc $te[$h].clave
        $porClave = @{}; foreach ($f in $act) { $porClave[[string]$f[$te[$h].clave]] = $f }
        foreach ($k in $SNAP.tablas[$h].Keys) {
            if (-not $porClave.ContainsKey($k)) { continue }
            $antes = $SNAP.tablas[$h][$k]; $ahora = $porClave[$k]
            foreach ($col in $te[$h].editables) {
                if ($antes.Contains($col) -and -not (Same $antes[$col] $ahora[$col])) { $cambios.textos.Add([ordered]@{ hoja = $h; clave = $k; columna = $col; antes = $antes[$col]; ahora = $ahora[$col] }) }
            }
            if ($h -eq 'Detalle de horas' -and $antes.Contains('Horas finales') -and -not (Same $antes['Horas finales'] $ahora['Horas finales'])) {
                $cambios.horas.Add([ordered]@{ clave = $k; antes = $antes['Horas finales']; ahora = $ahora['Horas finales']; tarea = $ahora['Tarea'] })
            }
        }
    }

    # Filas completas de Supuestos y Riesgos
    foreach ($h in (Get-HojasFilasManuales)) {
        if ($nombres -notcontains $h) { continue }
        $antesH = Get-Val $SNAP.filas $h ([ordered]@{})
        $actuales = @{}
        foreach ($f in (Read-TablaHasta $wb.Worksheets.Item($h) 3 'ID')) {
            $id = [string]$f['ID']; $o = [ordered]@{}; foreach ($k in $f.Keys) { if ($k -ne '_fila') { $o[$k] = $f[$k] } }
            $actuales[$id] = $true
            if (-not $antesH.Contains($id)) { $cambios.filas.Add([ordered]@{ hoja = $h; id = $id; tipo = 'nueva'; fila = $o }); continue }
            $dif = $false; foreach ($k in $o.Keys) { if (-not (Same $o[$k] (Get-Val $antesH[$id] $k ''))) { $dif = $true } }
            if ($dif) { $cambios.filas.Add([ordered]@{ hoja = $h; id = $id; tipo = 'modificada'; fila = $o }) }
        }
        foreach ($id in $antesH.Keys) { if (-not $actuales.ContainsKey($id)) { $cambios.noConservables.Add("${h}: se eliminó la fila $id (se volverá a generar; si no aplica, márquela como cerrada en vez de borrarla).") } }
    }
    $wb.Close($false)
}
finally { Close-Excel $xl }

$n = $cambios.planning.Count + $cambios.horas.Count + $cambios.textos.Count + $cambios.filas.Count
Write-JsonFile $salidaJson $cambios
Write-Host "Cambios manuales detectados en $(Split-Path -Leaf $Archivo):"
foreach ($c in $cambios.planning) { Write-Host (" - Planning $($c.id) $($c.clave): " + $(if ($c.diasCambiados) { "horas por día [$($c.diasAntes)] -> [$($c.diasAhora)]" } else { '' }) + $(if ($c.recursoAntes -ne $c.recursoAhora) { " recurso $($c.recursoAntes) -> $($c.recursoAhora)" } else { '' })) }
foreach ($c in $cambios.horas) { Write-Host " - Detalle de horas $($c.clave): Horas finales $($c.antes) -> $($c.ahora)" }
foreach ($c in $cambios.textos) { Write-Host " - $($c.hoja) [$($c.clave)] $($c.columna): '$($c.antes)' -> '$($c.ahora)'" }
foreach ($c in $cambios.filas) { Write-Host " - $($c.hoja) $($c.id): fila $($c.tipo)" }
foreach ($c in $cambios.noConservables) { Write-Host " - NO SE CONSERVA: $c" }
foreach ($c in $cambios.hojasPropias) { Write-Host " - Hoja propia '$c': se copiará tal cual al regenerar." }
if ($n -eq 0 -and $cambios.noConservables.Count -eq 0) { Write-Host ' (solo cambios de formato o en celdas calculadas; no se conservan)' }

if ($Aplicar) {
    $rutaAj = Join-Path $Trabajo 'ajustes.json'
    $AJUSTES = Read-JsonFile $rutaAj; if ($null -eq $AJUSTES) { $AJUSTES = [ordered]@{} }
    foreach ($k in @('tareas', 'fijadas', 'textos', 'filasManuales')) { if (-not $AJUSTES.Contains($k)) { $AJUSTES[$k] = [ordered]@{} } }
    if (-not $AJUSTES.Contains('historial')) { $AJUSTES['historial'] = New-Object System.Collections.Generic.List[object] }
    foreach ($c in $cambios.planning) {
        if ($c.diasCambiados) { $AJUSTES.fijadas[$c.clave] = [ordered]@{ dias = $c.dias; recurso = $c.recursoAhora; desde = $cambios.revisado } }
        elseif ($c.recursoAntes -ne $c.recursoAhora) {
            if (-not $AJUSTES.tareas.Contains($c.clave)) { $AJUSTES.tareas[$c.clave] = [ordered]@{} }
            $AJUSTES.tareas[$c.clave]['recurso'] = $c.recursoAhora
        }
    }
    foreach ($c in $cambios.horas) {
        if ($AJUSTES.fijadas.Contains($c.clave)) { continue }
        if (-not $AJUSTES.tareas.Contains($c.clave)) { $AJUSTES.tareas[$c.clave] = [ordered]@{} }
        $AJUSTES.tareas[$c.clave]['horas'] = [int](To-Num $c.ahora); $AJUSTES.tareas[$c.clave]['motivo'] = "editado en Detalle de horas ($($cambios.revisado))"
    }
    foreach ($c in $cambios.textos) {
        if (-not $AJUSTES.textos.Contains($c.hoja)) { $AJUSTES.textos[$c.hoja] = [ordered]@{} }
        if (-not $AJUSTES.textos[$c.hoja].Contains($c.clave)) { $AJUSTES.textos[$c.hoja][$c.clave] = [ordered]@{} }
        $AJUSTES.textos[$c.hoja][$c.clave][$c.columna] = $c.ahora
    }
    foreach ($c in $cambios.filas) {
        if (-not $AJUSTES.filasManuales.Contains($c.hoja)) { $AJUSTES.filasManuales[$c.hoja] = [ordered]@{} }
        $AJUSTES.filasManuales[$c.hoja][$c.id] = $c.fila
    }
    $AJUSTES['aplicadoDesde'] = $hash
    $AJUSTES.historial.Add([ordered]@{ fecha = $cambios.revisado; archivo = $Archivo; cambios = $n; noConservables = $cambios.noConservables.Count })
    Write-JsonFile $rutaAj $AJUSTES
    Write-Host "Ajustes guardados en $rutaAj. Ejecute 2-Calcular, 3-Programar, 4-Validar y 6-Generar para regenerar con sus cambios."
}
if ($n -gt 0 -or $cambios.noConservables.Count -gt 0) { exit 1 }
exit 0
