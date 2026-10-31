<#
.SYNOPSIS
  Pasos 11-12 del procedimiento: genera estimacion_<slug>.xlsx replicando la estructura del modelo de salida
  (assets/Ejemplos Salidas/Ejemplo_proyecto.xlsx), recalcula y deja un snapshot para detectar cambios manuales.
.DESCRIPTION
  Protección de cambios manuales:
  - Si el Excel ya existe, primero ejecuta 5-Detectar-Cambios.ps1. Si el usuario editó el archivo y esos cambios
    no se incorporaron a ajustes.json (5-Detectar-Cambios.ps1 -Aplicar), se DETIENE sin tocar nada.
  - -DescartarCambios regenera igual (el archivo anterior queda en backup/).
  - Las hojas que no son estándar (creadas por el usuario) se copian tal cual al archivo nuevo.
  - Los textos editables y las filas manuales de Supuestos/Riesgos guardados en ajustes.json se vuelven a escribir.
.EXAMPLE
  .\6-Generar-Excel.ps1 -Trabajo 'Estimaciones\MigraMacro\_trabajo' -Salida 'Estimaciones\MigraMacro\estimacion_MigraMacro.xlsx'
#>
param(
    [Parameter(Mandatory = $true)][string]$Trabajo,
    [Parameter(Mandatory = $true)][string]$Salida,
    [switch]$DescartarCambios
)
. (Join-Path $PSScriptRoot 'Comun.ps1')

$PLAN = Read-JsonFile (Join-Path $Trabajo 'planning.json')
$WBS = Read-JsonFile (Join-Path $Trabajo 'wbs.json')
$ENT = Read-JsonFile (Join-Path $Trabajo 'entrada.json')
$ASIG = Read-JsonFile (Join-Path $Trabajo 'asignacion.json')
$AJUSTES = Read-JsonFile (Join-Path $Trabajo 'ajustes.json'); if ($null -eq $AJUSTES) { $AJUSTES = [ordered]@{} }
$VAL = Read-JsonFile (Join-Path $Trabajo 'validaciones.json')
if ($null -eq $PLAN -or $null -eq $VAL) { throw 'Faltan planning.json o validaciones.json. Ejecute 3-Programar.ps1 y 4-Validar-Planning.ps1.' }
$Salida = [IO.Path]::GetFullPath($Salida)
# Seguridad: solo se escribe un .xlsx de estimación, nunca sobre la entrada ni sobre otro tipo de archivo.
if ([IO.Path]::GetExtension($Salida) -ne '.xlsx') { throw "La salida debe ser un archivo .xlsx: $Salida" }
if ($ENT.archivo -and ([IO.Path]::GetFullPath([string]$ENT.archivo) -eq $Salida)) { throw 'La salida no puede ser el archivo de entrada.' }
$snapRuta = Join-Path $Trabajo 'snapshot.json'

# ---------- 1. Proteger cambios manuales ----------
$respaldo = $null
if (Test-Path -LiteralPath $Salida) {
    if (-not $DescartarCambios) {
        & (Join-Path $PSScriptRoot '5-Detectar-Cambios.ps1') -Trabajo $Trabajo -Archivo $Salida
        $cod = $LASTEXITCODE
        $hash = Get-FileHashSafe $Salida
        if ($cod -eq 1 -and (Get-Val $AJUSTES 'aplicadoDesde') -ne $hash) {
            Write-Host ''
            Write-Host 'DETENIDO: el Excel tiene cambios manuales que aún no están en ajustes.json.'
            Write-Host 'Revise _trabajo\cambios_manuales.json con el usuario y luego:'
            Write-Host '  - para conservarlos:  .\5-Detectar-Cambios.ps1 -Trabajo ... -Archivo ... -Aplicar   y vuelva a ejecutar 2, 3, 4 y 6'
            Write-Host '  - para descartarlos:  .\6-Generar-Excel.ps1 ... -DescartarCambios   (el archivo actual queda en backup\)'
            exit 5
        }
        if ($cod -eq 2) {
            Write-Host 'DETENIDO: no hay snapshot del archivo actual (se generó fuera de estos scripts), así que no se puede saber qué se editó a mano.'
            Write-Host 'Pida al usuario que confirme y use -DescartarCambios (el archivo actual queda en backup\), o incorpore sus cambios a ajustes.json.'
            exit 5
        }
    }
    $dirB = Join-Path (Split-Path -Parent $Salida) 'backup'
    if (-not (Test-Path -LiteralPath $dirB)) { New-Item -ItemType Directory -Path $dirB | Out-Null }
    $base = [IO.Path]::GetFileNameWithoutExtension($Salida) + '_' + (Get-Date).ToString('yyyyMMdd_HHmm')
    $respaldo = Join-Path $dirB "$base.xlsx"; $n = 1
    while (Test-Path -LiteralPath $respaldo) { $n++; $respaldo = Join-Path $dirB "${base}_$n.xlsx" }
    Copy-Item -LiteralPath $Salida -Destination $respaldo
    Write-Host "Respaldo: $respaldo"
}

# ---------- Datos ----------
$PROY = $WBS.proyecto
$tipo = [string]$PROY.tipo
$hojas = Get-HojasEstandar $tipo
$hP = "'Planning - $tipo'"
$TAREAS = @($PLAN.tareas); $RECS = @($PLAN.recursos); $LEY = @($PLAN.leyenda); $ND = [int]$PLAN.ultimoDia
$textos = Get-Val $AJUSTES 'textos' ([ordered]@{})
$filasMan = Get-Val $AJUSTES 'filasManuales' ([ordered]@{})
function Texto([string]$hoja, [string]$clave, [string]$col, [string]$defecto) {
    $h = Get-Val $textos $hoja; if ($null -eq $h) { return $defecto }
    $f = Get-Val $h $clave; if ($null -eq $f) { return $defecto }
    $v = Get-Val $f $col; if ($null -eq $v) { return $defecto }
    return [string]$v
}
$compDe = @{}; foreach ($c in $WBS.componentes) { $compDe[$c.id] = $c }
$colorDe = @{}; foreach ($l in $LEY) { $colorDe[$l.aplicativoOrigen] = $l.color }
$colorOla = @{ 'Transversal' = '#EDEDED'; 'Ola 1' = '#DDEBF7'; 'Ola 2' = '#E2EFDA'; 'Ola 3' = '#FFF2CC'; 'Ola 4' = '#FCE4D6' }
function ColorOla([string]$o) { if ($colorOla.ContainsKey($o)) { return $colorOla[$o] }; if ($o -match '^Ola') { return '#EDE7F6' }; return '#F2F2F2' }
function FechaDia([int]$d) { if ($d -lt 1) { return $null }; return (Get-IsoDate $PLAN.fechas[[Math]::Min($d, $ND) - 1]).ToOADate() }

$xl = Open-Excel
try {
    $wb = $xl.Workbooks.Add()
    $xl.Calculation = -4135   # manual mientras se escribe
    while ($wb.Worksheets.Count -lt $hojas.Count) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
    for ($i = 0; $i -lt $hojas.Count; $i++) { $wb.Worksheets.Item($i + 1).Name = $hojas[$i] }
    $HOJA = @{}; foreach ($h in $hojas) { $HOJA[$h] = $wb.Worksheets.Item($h) }

    function Style-Title($ws, [string]$texto, [int]$nCols) {
        Set-Cell $ws 1 1 $texto
        $rg = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item(1, $nCols))
        $rg.Interior.Color = ConvertTo-BGR '#17365D'; $rg.Font.Color = ConvertTo-BGR '#FFFFFF'; $rg.Font.Bold = $true; $rg.Font.Size = 15; $ws.Rows.Item(1).RowHeight = 26
    }
    function Style-Header($ws, [int]$fila, [int]$col, [string[]]$enc, [string]$fondo = '#1F4E78') {
        Set-Row $ws $fila $col $enc
        $rg = $ws.Range($ws.Cells.Item($fila, $col), $ws.Cells.Item($fila, $col + $enc.Count - 1))
        $rg.Interior.Color = ConvertTo-BGR $fondo; $rg.Font.Color = ConvertTo-BGR '#FFFFFF'; $rg.Font.Bold = $true; $rg.WrapText = $true; $rg.VerticalAlignment = -4108
        $ws.Rows.Item($fila).RowHeight = 32
    }
    function Style-Borders($ws, [int]$f1, [int]$c1, [int]$f2, [int]$c2) {
        if ($f2 -lt $f1) { return }
        $rg = $ws.Range($ws.Cells.Item($f1, $c1), $ws.Cells.Item($f2, $c2)); $rg.Borders.LineStyle = 1; $rg.Borders.Weight = 2
    }
    function Style-Sub($ws, [int]$fila, [string]$texto) { Set-Cell $ws $fila 1 $texto; $c = $ws.Cells.Item($fila, 1); $c.Font.Bold = $true; $c.Font.Size = 12; $c.Font.Color = ConvertTo-BGR '#17365D' }
    function Fill($ws, [int]$f, [int]$c, [string]$hex) { $ws.Cells.Item($f, $c).Interior.Color = ConvertTo-BGR $hex }
    function Add-CF($rg, [string]$formula, [string]$fondo, [string]$fuente) {
        $fc = $rg.FormatConditions.Add(2, 0, $formula)
        if ($fondo) { $fc.Interior.Color = ConvertTo-BGR $fondo }
        if ($fuente) { $fc.Font.Color = ConvertTo-BGR $fuente }
    }
    function Widths($ws, [double[]]$anchos) { for ($i = 0; $i -lt $anchos.Count; $i++) { $ws.Columns.Item($i + 1).ColumnWidth = $anchos[$i] } }

    # ---------- 8. Supuestos (primero: define el nombre Feriados) ----------
    $ws = $HOJA['Supuestos']
    Style-Title $ws "Supuestos — $($PROY.nombre)" 7
    $encS = @('ID', 'Categoría', 'Supuesto', 'Valor / regla', 'Punto de validación', 'Impacto si no se cumple', 'Estado')
    Style-Header $ws 3 1 $encS
    $par = $WBS.parametros
    $sup = New-Object System.Collections.Generic.List[object]
    $sup.Add(@('Base de horas', 'Modelo paramétrico del arnés con los parámetros de la entrada', "Productividad $(Get-Val $par 'Productividad LOC/h (sin IA)' 45) LOC/h; contingencia por componente (máx. 20 %)", 'Cierre de 02 Análisis', 'Re-estimar componentes afectados', 'Assumption'))
    $sup.Add(@('Capacidad', "Carga diaria por recurso: $(if ((Get-Val $par 'Carga diaria máxima' 'Jornada') -eq 'Jornada') { 'jornada del proyecto' } else { 'capacidad productiva del rol' }) ajustada por dedicación", "Jornada $($PROY.jornada) h; nadie supera su capacidad", 'Semanal', 'Mayor duración', 'Constraint'))
    $sup.Add(@('Calendario', 'Días hábiles desde la fecha de inicio, sin feriados ni vacaciones declarados', "Inicio $($PROY.fechaInicio); $(@($PLAN.feriados).Count) feriados", 'Kick-off', 'Desplazamiento de hitos', 'Constraint'))
    $sup.Add(@('Preparación', 'Entendimiento (01.1) y relevamiento (01.2) agregados por el arnés', 'Horas marcadas como Asumida', 'G0', 'Retraso del inicio de análisis', 'Decision'))
    $sup.Add(@('Certificación', "Certificación a cargo de: $($PROY.certificacion)", "Mínimo $($PROY.diasCertMin) días hábiles por bloque", 'Inicio de 06', 'Mayor duración de certificación', 'Decision'))
    if ($PROY.alcance -ne 'Certificación') { $sup.Add(@('Despliegue', "Agrupación: $($PROY.agrupacion)", 'Plan de pase, ensayo y ejecución por pase', 'G5', 'Re-planificar ventanas', 'Decision')) }
    if ($PROY.alcance -eq 'Hypercare') { $sup.Add(@('Seguimiento', 'Hypercare posterior al pase', "$($PROY.diasHyper) días hábiles", 'G5', 'Soporte adicional', 'Decision')) }
    $sup.Add(@('Asignación', 'Un desarrollador principal por componente; revisión de código por el Líder Técnico', 'Ver hoja Equipo', 'Inicio de cada ola', 'Re-asignación y nueva programación', 'Decision'))
    $sup.Add(@('Secuencia', "Modelo de entrega: $($PROY.modelo)", 'Ver Estrategia de oleadas', 'Inicio de cada ola', 'Re-secuenciar', 'Decision'))
    if ($PROY.usaIA) { $sup.Add(@('Factores IA', 'Uso de asistente IA en el desarrollo', 'Factores por tipo de tarea de la hoja Parametros (sin calibrar hasta el piloto)', 'Fin del piloto (Ola 1)', 'Recalibrar factores y re-estimar', 'Assumption')) }
    foreach ($x in @($WBS.excluidas)) { if ($x) { $sup.Add(@('Tarea excluida', "$($x.tarea) ($($x.clave))", $(if ($x.motivo) { $x.motivo } else { 'No aplica a esta migración' }), 'Kick-off', "Re-incluir ($($x.horas) h)", 'Decision')) } }
    if ($PROY.fueraAlcance) { $sup.Add(@('Exclusión', 'Fuera de alcance declarado en la entrada', $PROY.fueraAlcance, 'Kick-off', 'Cambio de alcance', 'Constraint')) }
    $filasSup = [ordered]@{}; $i = 0
    foreach ($s in $sup) { $i++; $id = 'ASM-{0:D3}' -f $i; $filasSup[$id] = @($id) + $s }
    foreach ($id in (Get-Val $filasMan 'Supuestos' ([ordered]@{})).Keys) { $f = $filasMan['Supuestos'][$id]; $filasSup[$id] = @($encS | ForEach-Object { [string](Get-Val $f $_ '') }); $filasSup[$id][0] = $id }
    $fila = 4; foreach ($id in $filasSup.Keys) { Set-Row $ws $fila 1 $filasSup[$id]; $fila++ }
    Style-Borders $ws 4 1 ($fila - 1) 7
    $fila += 1; Style-Sub $ws $fila 'Feriados (rango Feriados: puede agregar fechas en las filas vacías)'; $fila++
    Style-Header $ws $fila 1 @('#', 'Fecha', 'Descripción'); $f1 = $fila + 1
    $desc = @{}; foreach ($c in $ENT.calendario) { if ($c['Fecha']) { $desc[[string]$c['Fecha']] = [string](Get-Val $c 'Descripción' '') } }
    for ($i = 0; $i -lt 30; $i++) {
        Set-Cell $ws ($f1 + $i) 1 ($i + 1)
        if ($i -lt @($PLAN.feriados).Count) { $fe = [string]$PLAN.feriados[$i]; Set-Cell $ws ($f1 + $i) 2 ((Get-IsoDate $fe).ToOADate()); Set-Cell $ws ($f1 + $i) 3 $desc[$fe] }
    }
    $ws.Range("B${f1}:B$($f1 + 29)").NumberFormat = 'dd/mm/yyyy'
    Style-Borders $ws $f1 1 ($f1 + 29) 3
    [void]$wb.Names.Add('Feriados', "=Supuestos!`$B`$${f1}:`$B`$$($f1 + 29)")
    Widths $ws @(10, 14, 60, 40, 22, 30, 12)

    # ---------- 1. Resumen ejecutivo (define FechaInicio) ----------
    $ws = $HOJA['Resumen ejecutivo']
    $tit = $(if ($PROY.cliente) { "$($PROY.cliente) — " } else { '' }) + "$($PROY.nombre): estimación ordenada por oleadas ($($PROY.version))"
    Style-Title $ws $tit 11
    Style-Header $ws 3 1 @('Campo', 'Valor')
    $resumen = @(
        @('Artefacto', $Salida), @('Fuente de horas', "Modelo paramétrico del arnés + entrada $(Split-Path -Leaf $ENT.archivo)"),
        @('Plantilla de formato', 'assets/Ejemplos Salidas/Ejemplo_proyecto.xlsx'), @('Alcance', "$($PROY.alcanceResumen) (hasta $($PROY.alcance))"),
        @('Modelo de entrega', "$($PROY.modelo) · $($PROY.agrupacion)"), @('Fecha de inicio (FechaInicio)', (Get-IsoDate $PROY.fechaInicio).ToOADate()),
        @('Fecha fin estimada', "=$hP!B6"), @('Jornada (h/día)', [int]$PROY.jornada), @('Equipo (personas)', $RECS.Count),
        @('Fecha objetivo', $(if ($PROY.fechaObjetivo) { (Get-IsoDate $PROY.fechaObjetivo).ToOADate() } else { '' })),
        @('Holgura (días hábiles; negativo = brecha)', '=IF(B13="","",NETWORKDAYS(B10,B13,Feriados)-1)')
    )
    for ($i = 0; $i -lt $resumen.Count; $i++) { Set-Row $ws (4 + $i) 1 $resumen[$i] }
    $ws.Range('B9:B10').NumberFormat = 'dd/mm/yyyy'; $ws.Range('B13').NumberFormat = 'dd/mm/yyyy'
    $c = $ws.Range('B9'); $c.Interior.Color = ConvertTo-BGR '#FFF2CC'; $c.Font.Color = ConvertTo-BGR '#0000FF'; $c.Font.Bold = $true
    [void]$wb.Names.Add('FechaInicio', "='Resumen ejecutivo'!`$B`$9")
    Style-Borders $ws 4 1 (3 + $resumen.Count) 2
    $ws.Tab.Color = ConvertTo-BGR '#17365D'

    # ---------- 4. Planning ----------
    $ws = $HOJA["Planning - $tipo"]
    $nR = $RECS.Count; $nA = $LEY.Count
    $bloque = [Math]::Max(11, [Math]::Max($nR + 2, $nA + 1))
    $HR = $bloque + 7; $D1 = $HR + 1; $NF = $D1 + $TAREAS.Count - 1
    $LC = Get-ColLetter (10 + $ND); $tot = $nR + 2
    Set-Row $ws 1 1 @('Duración', '', 'Leyenda de aplicativos origen'); Set-Row $ws 1 7 @('Cap. h/día', 'Rol', 'Recurso', 'Total horas')
    for ($d = 1; $d -le $ND; $d++) { Set-Cell $ws 1 (10 + $d) $d }
    Set-Row $ws 2 1 @('Horas', "=J$tot"); Set-Row $ws 3 1 @('Días', "=COUNTIF(K${tot}:${LC}${tot},`">0`")"); Set-Row $ws 4 1 @('Semanas', '=ROUND(B3/5,1)')
    Set-Row $ws 5 1 @('Inicio', '=FechaInicio'); Set-Row $ws 6 1 @('Fin', '=WORKDAY(B5,B3-1,Feriados)'); $ws.Range('B5:B6').NumberFormat = 'dd/mm/yyyy'
    Set-Cell $ws 8 1 'Verde: igual a cap. (jornada completa)'; Set-Cell $ws 9 1 'Amarillo: menor a cap. (carga parcial)'; Set-Cell $ws 10 1 'Rojo: mayor a cap. (sobreasignación)'
    Fill $ws 8 1 '#C6EFCE'; Fill $ws 9 1 '#FFEB9C'; Fill $ws 10 1 '#FFC7CE'
    for ($i = 0; $i -lt $nA; $i++) { Set-Cell $ws (2 + $i) 3 $LEY[$i].aplicativoOrigen; $ws.Range("C$(2 + $i):F$(2 + $i)").Interior.Color = ConvertTo-BGR $LEY[$i].color }
    for ($i = 0; $i -lt $nR; $i++) { Set-Row $ws (2 + $i) 7 @([int]$RECS[$i].cap, $RECS[$i].rol, $RECS[$i].nombre, "=SUM(K$(2 + $i):${LC}$(2 + $i))") }
    if ($nR -gt 0) { $ws.Range("K2:${LC}$($nR + 1)").Formula = "=SUMIF(`$I`$${D1}:`$I`$${NF},`$I2,K`$${D1}:K`$${NF})" }
    Set-Cell $ws $tot 9 'TOTAL >>'; $ws.Range("J${tot}:${LC}${tot}").Formula = "=SUM(J2:J$($nR + 1))"; $ws.Range("I${tot}:${LC}${tot}").Font.Bold = $true
    $rgCarga = $ws.Range("K2:${LC}$($nR + 1)")
    Add-CF $rgCarga '=K2>$G2' '#FFC7CE' '#9C0006'; Add-CF $rgCarga '=K2=$G2' '#C6EFCE' '#006100'; Add-CF $rgCarga '=(K2>0)*(K2<$G2)' '#FFEB9C' '#9C5700'
    $fH = $bloque + 4; Set-Cell $ws $fH 10 'Hitos >>'
    foreach ($g in @($PLAN.hitos | Group-Object { [int]$_.dia })) { Set-Cell $ws $fH (10 + [int]$g.Name) ((@($g.Group | ForEach-Object { $_.codigo })) -join ' ') }
    Add-CF $ws.Range("K${fH}:${LC}${fH}") "=K$fH<>`"`"" '#070F26' '#FFC400'
    Set-Cell $ws ($bloque + 5) 10 'Semanas >>'
    for ($d = 1; $d -le $ND; $d += 5) { Set-Cell $ws ($bloque + 5) (10 + $d) ('S' + ([Math]::Floor(($d - 1) / 5) + 1)) }
    $ws.Rows.Item($bloque + 5).Font.Bold = $true
    $rgF = $ws.Range("K$($bloque + 6):${LC}$($bloque + 6)"); $rgF.Formula = "=WORKDAY(FechaInicio,K`$$HR-1,Feriados)"; $rgF.NumberFormat = 'dd/mm'; $rgF.Font.Size = 8; $rgF.Orientation = 90
    $encP = @('ID', 'Ola', 'Fase', 'Aplicativo origen', 'Aplicativo destino', 'Plataforma', 'Componente', 'Tarea', 'Recurso asignado', 'Horas')
    Set-Row $ws $HR 1 $encP; for ($d = 1; $d -le $ND; $d++) { Set-Cell $ws $HR (10 + $d) $d }
    $rgH = $ws.Range("A${HR}:${LC}${HR}"); $rgH.Interior.Color = ConvertTo-BGR '#070F26'; $rgH.Font.Color = ConvertTo-BGR '#FFFFFF'; $rgH.Font.Bold = $true; $rgH.Font.Name = 'Arial'; $rgH.Font.Size = 10; $rgH.WrapText = $true
    $ws.Rows.Item($HR).RowHeight = 42
    $fila = $D1
    foreach ($t in $TAREAS) {
        $pref = $(if ($t.origen -eq 'Asumida') { '[Asumida] ' } elseif ($t.origen -eq 'Contingencia') { '[Contingencia] ' } else { '' })
        $comp = $(if ($compDe.ContainsKey($t.componente)) { "$($t.componente) – $($t.componenteNombre)" } else { $t.componenteNombre })
        Set-Row $ws $fila 1 @($t.id, $t.ola, $t.fase, $t.aplicativoOrigen, $t.aplicativo, $t.plataforma, $comp, ($pref + $t.tarea), $t.recurso)
        Fill $ws $fila 2 (ColorOla $t.ola)
        if ($colorDe.ContainsKey($t.aplicativoOrigen)) { Fill $ws $fila 4 $colorDe[$t.aplicativoOrigen] }
        foreach ($k in $t.dias.Keys) { $h = [int](To-Num $t.dias[$k]); if ($h -gt 0 -and [int]$k -le $ND) { Set-Cell $ws $fila (10 + [int]$k) $h } }
        $fila++
    }
    if ($TAREAS.Count -gt 0) {
        $ws.Range("J${D1}:J${NF}").Formula = "=SUM(K${D1}:${LC}${D1})"
        $rgD = $ws.Range("K${D1}:${LC}${NF}"); $rgD.NumberFormat = '0;-0;"-"'; $rgD.Font.Name = 'Arial'; $rgD.Font.Size = 9
        foreach ($l in $LEY) { Add-CF $rgD ("=(`$D$D1=`"" + $l.aplicativoOrigen.Replace('"', '""') + "`")*(K$D1>0)") $l.color '#000000' }
        Style-Borders $ws $D1 1 $NF 10
    }
    Widths $ws @(7, 11, 24, 22, 14, 11, 26, 60, 22, 8)
    $ws.Range("K:${LC}").ColumnWidth = 4.5
    if ($TAREAS.Count -gt 0) { [void]$ws.Range("A${HR}:J${NF}").AutoFilter() }
    $ws.Tab.Color = ConvertTo-BGR '#070F26'
    $rP = { param([string]$col) "$hP!`$${col}`$${D1}:`$${col}`$${NF}" }
    $rI = & $rP 'I'; $rJ = & $rP 'J'; $rB = & $rP 'B'; $rD = & $rP 'D'; $rE = & $rP 'E'

    # ---------- 7. Detalle de horas ----------
    $ws = $HOJA['Detalle de horas']
    Style-Title $ws "Detalle de horas (trazabilidad) — $($PROY.nombre)" 19
    $encDH = @('ID origen', 'Código componente', 'Componente', 'Aplicativo origen', 'Aplicativo destino', 'Plataforma', 'Complejidad', 'Código tarea', 'Tarea', 'Driver / fórmula', 'Horas base', 'Factor IA', 'Horas finales', 'Fase', 'Recurso', 'Ola', 'ID planning', 'Origen', 'Observación')
    Style-Header $ws 3 1 $encDH
    $fila = 4
    foreach ($t in $TAREAS) {
        $hb = To-Num $t.horasBase; $fa = To-Num $t.factorIA 1
        $calc = [int][Math]::Ceiling([Math]::Round($hb * $fa, 1))
        $hf = $(if ($t.origen -eq 'Paramétrica' -and -not $t.fijada -and $t.tipo -eq 'normal' -and $calc -eq [int]$t.horas) { "=ROUNDUP(ROUND(K${fila}*L${fila},1),0)" } else { [int]$t.horas })
        $cc = $(if ($compDe.ContainsKey($t.componente)) { $t.componente } else { $t.componenteNombre })
        Set-Row $ws $fila 1 @($t.clave, $cc, $t.componenteNombre, $t.aplicativoOrigen, $t.aplicativo, $t.plataforma, $t.complejidad, $t.codigo, $t.tarea, $t.driver, $hb, $fa, $hf, $t.fase, $t.recurso, $t.ola, $t.id, $t.origen, (Texto 'Detalle de horas' $t.clave 'Observación' ([string]$t.observacion)))
        $fila++
    }
    $fT = $fila; Set-Cell $ws $fT 1 'Total'; Set-Cell $ws $fT 11 "=SUM(K4:K$($fT - 1))"; Set-Cell $ws $fT 13 "=SUM(M4:M$($fT - 1))"; $ws.Rows.Item($fT).Font.Bold = $true
    Style-Borders $ws 4 1 $fT 19
    $ws.Range("K4:K$fT").NumberFormat = '#,##0.0'; $ws.Range("M4:M$fT").NumberFormat = '#,##0'
    $dh = { param([string]$col) "'Detalle de horas'!`$${col}`$4:`$${col}`$$($fT - 1)" }
    $dhB = & $dh 'B'; $dhK = & $dh 'K'; $dhM = & $dh 'M'; $dhO = & $dh 'O'; $dhR = & $dh 'R'; $dhF = & $dh 'F'
    $fila = $fT + 2; Style-Sub $ws $fila 'Cambios y decisiones respecto a la entrada'; $fila++
    $dec = @()
    foreach ($c in $WBS.componentes) { if ($c.motivoComplejidad -match 'forzada') { $dec += "$($c.id): complejidad forzada en la entrada ($($c.complejidad))." }; if ($null -ne $c.override -and (To-Num $c.override) -gt 0) { $dec += "$($c.id): horas override de la entrada ($($c.override) h)." } }
    foreach ($t in $TAREAS) { if ($t.fijada) { $dec += "$($t.id) $($t.clave): horas por día fijadas por el usuario." } elseif ($t.observacion -match 'Ajuste manual') { $dec += "$($t.id) $($t.clave): $($t.observacion)" } }
    if ($dec.Count -eq 0) { $dec = @('Sin cambios respecto a la entrada.') }
    foreach ($x in $dec) { Set-Cell $ws $fila 1 "• $x"; $fila++ }
    Widths $ws @(22, 12, 26, 20, 16, 12, 10, 9, 60, 28, 9, 8, 9, 22, 20, 11, 9, 13, 40)

    # ---------- 2. Estrategia de oleadas ----------
    $ws = $HOJA['Estrategia de oleadas']
    Style-Title $ws "Estrategia de oleadas — $($PROY.nombre)" 20
    $encEO = @('Orden', 'Ola', 'Código', 'Componente', 'Aplicativo origen', 'Plataforma', 'Aplicativo destino', 'Estrategia', 'Complejidad (puntaje)', 'LOC', 'Pantallas', 'Integraciones', 'Horas base', 'Horas finales', 'Desarrollador(es)', 'Listo para certificar (día)', 'Despliegue (fecha)', 'Motivo de la secuencia', 'Riesgos / dependencias', 'Información faltante o por validar')
    Style-Header $ws 3 1 $encEO
    $bloqueDe = @{}; foreach ($b in $PLAN.bloques) { foreach ($id in $b.componentes) { $bloqueDe[$id] = $b } }
    $paseDe = @{}; foreach ($pa in $PLAN.pases) { foreach ($bn in $pa.bloques) { $paseDe[$bn] = $pa } }
    $fila = 4; $orden = 0
    foreach ($c in @($WBS.componentes | Sort-Object { [int]($_.ola -replace '\D', '') }, orden, lote)) {
        $orden++; $b = $bloqueDe[$c.id]; $pa = $(if ($b) { $paseDe[$b.nombre] } else { $null })
        $riesgos = @(); if ($c.dependencia -ne 'Ninguna') { $riesgos += "Dependencia externa $($c.dependencia)" }; if ($c.riesgo -eq 'Alto') { $riesgos += 'Riesgo técnico alto' }
        $listo = $(if ($b -and $b.listo) { [int]$b.listo } else { '' }); $dep = $(if ($pa -and $pa.dia) { FechaDia ([int]$pa.dia) } else { '' })
        $motivoDef = $(if ($c.motivo) { $c.motivo } else { "Criterio dominante: $($c.criterioDominante)" })
        Set-Row $ws $fila 1 @($orden, $c.ola, $c.id, $c.nombre, $c.aplicativoOrigen, $c.plataforma, $c.aplicativo, $c.estrategia, "$($c.complejidad) ($($c.puntaje))", $c.loc, $c.pantallas, $c.integraciones,
            "=SUMIF($dhB,C$fila,$dhK)", "=SUMIF($dhB,C$fila,$dhM)", $c.desarrollador, $listo, $dep,
            (Texto 'Estrategia de oleadas' $c.id 'Motivo de la secuencia' $motivoDef), (Texto 'Estrategia de oleadas' $c.id 'Riesgos / dependencias' ($riesgos -join '; ')),
            (Texto 'Estrategia de oleadas' $c.id 'Información faltante o por validar' $c.faltante))
        if ($colorDe.ContainsKey($c.aplicativoOrigen)) { Fill $ws $fila 5 $colorDe[$c.aplicativoOrigen] }
        $fila++
    }
    $fEO = $fila - 1; Style-Borders $ws 4 1 $fEO 20; $ws.Range("Q4:Q$fEO").NumberFormat = 'dd/mm/yyyy'; $ws.Range("M4:N$fEO").NumberFormat = '#,##0'
    $fila += 1; Style-Sub $ws $fila 'Base común por aplicativo destino'; $fila++
    foreach ($app in @($TAREAS | Where-Object { $_.componente -like 'Base común*' } | ForEach-Object { $_.aplicativo } | Sort-Object -Unique)) { Set-Row $ws $fila 1 @('', '', '', "Base común $app", '', '', $app, '', '', '', '', '', "=SUMIF($dhB,D$fila,$dhK)", "=SUMIF($dhB,D$fila,$dhM)"); $fila++ }
    $fila += 1; Style-Sub $ws $fila 'Actividades reutilizables o agrupables'; $fila++
    Style-Header $ws $fila 1 @('Actividad', 'Componentes / aplicativos', 'Cómo se aprovecha en el plan'); $fila++
    foreach ($x in @(Get-Val $ASIG 'reutilizables' @())) { Set-Row $ws $fila 1 @([string](Get-Val $x 'actividad' ''), [string](Get-Val $x 'componentes' ''), [string](Get-Val $x 'aprovechamiento' '')); $fila++ }
    $fila += 1; Style-Sub $ws $fila 'Criterios de secuenciación aplicados'; $fila++
    $i = 0; foreach ($x in @(Get-Val $ASIG 'criterios' @())) { $i++; Set-Cell $ws $fila 1 "$i. $x"; $fila++ }
    Widths $ws @(7, 10, 10, 28, 20, 13, 16, 12, 13, 8, 9, 11, 9, 9, 18, 11, 12, 40, 30, 30)

    # ---------- 3. Matriz de entregables ----------
    $ws = $HOJA['Matriz de entregables']
    Style-Title $ws "Matriz de entregables — $($PROY.nombre)" 15
    $encME = @('Ola', 'Entregable', 'Componentes incluidos', 'Aplicativos destino', 'Plataforma', 'Objetivo del entregable', 'Inicio construcción (día)', 'Listo para certificar (día)', 'Certificación (días)', 'Fin de certificación (fecha)', 'Fecha de despliegue', 'Fin de seguimiento', 'Horas totales', 'Criterio de salida', 'Riesgos o dependencias')
    Style-Header $ws 3 1 $encME
    $fila = 4
    foreach ($b in $PLAN.bloques) {
        $cs = @($b.componentes | ForEach-Object { $compDe[$_] })
        $ini = @($TAREAS | Where-Object { $b.componentes -contains $_.componente -and $_.fase -like '04*' -and $_.inicio -gt 0 } | ForEach-Object { $_.inicio } | Measure-Object -Minimum).Minimum
        $pa = $paseDe[$b.nombre]
        $se = @($TAREAS | Where-Object { $_.bloque -eq $(if ($pa) { $pa.nombre } else { '' }) -and $_.fase -like '08*' } | ForEach-Object { $_.fin } | Measure-Object -Maximum).Maximum
        $horas = $(if ($PROY.agrupacion -eq 'Por aplicativo destino') { "=SUMIFS($rJ,$rB,A$fila,$rE,`"*`"&D$fila&`"*`")" } else { "=SUMIF($rB,A$fila,$rJ)" })
        Set-Row $ws $fila 1 @($b.ola, $b.nombre, (($cs | ForEach-Object { $_.id }) -join ', '), (@($b.aplicativos) -join ', '), (@($cs | ForEach-Object { $_.plataforma } | Sort-Object -Unique) -join ', '),
            (Texto 'Matriz de entregables' $b.nombre 'Objetivo del entregable' "Entregable certificable de $($b.nombre)"), $(if ($ini) { [int]$ini } else { '' }), $(if ($b.listo) { [int]$b.listo } else { '' }), [int]$b.diasCert,
            $(if ($b.finCert) { FechaDia ([int]$b.finCert) } else { '' }), $(if ($pa -and $pa.dia) { FechaDia ([int]$pa.dia) } else { '' }), $(if ($se) { FechaDia ([int]$se) } else { '' }), $horas,
            (Texto 'Matriz de entregables' $b.nombre 'Criterio de salida' 'Pruebas unitarias completas, certificación aprobada sin defectos críticos'), (Texto 'Matriz de entregables' $b.nombre 'Riesgos o dependencias' ''))
        $ws.Range("A${fila}:O${fila}").Interior.Color = ConvertTo-BGR '#DDEBF7'
        $fila++
    }
    if (@($TAREAS | Where-Object { $_.ola -eq 'Pase único' }).Count -gt 0) { Set-Row $ws $fila 1 @('Pase único', 'Pase único', 'Todos', '', '', 'Despliegue conjunto', '', '', '', '', $(FechaDia ([int]($PLAN.pases[0].dia))), '', "=SUMIF($rB,A$fila,$rJ)", '', ''); $fila++ }
    Style-Borders $ws 4 1 ($fila - 1) 15; $ws.Range("J4:L$fila").NumberFormat = 'dd/mm/yyyy'; $ws.Range("M4:M$fila").NumberFormat = '#,##0'
    Set-Cell $ws ($fila + 1) 1 'Nota: las horas de cada ola incluyen análisis, diseño, construcción, pruebas, certificación y despliegue de sus componentes; las transversales se cuentan aparte (ola Transversal).'
    Widths $ws @(10, 16, 26, 20, 14, 34, 11, 11, 11, 13, 13, 13, 11, 40, 30)

    # ---------- 5. Capacidad por recurso ----------
    $ws = $HOJA['Capacidad por recurso']
    $nS = [int][Math]::Ceiling($ND / 5.0)
    Style-Title $ws "Capacidad por recurso — $($PROY.nombre)" ($nS + 4)
    $semanas = @(); for ($s = 1; $s -le $nS; $s++) { $semanas += "S$s" }
    $sec = @(4, (6 + $nR), (8 + 2 * $nR), (10 + 3 * $nR))
    Style-Sub $ws ($sec[0] - 1) '1. Horas asignadas por semana'; Style-Header $ws $sec[0] 1 (@('Recurso', 'Rol', 'Cap. h/día') + $semanas + @('Total'))
    Style-Sub $ws ($sec[1] - 1) '2. Capacidad disponible por semana (h)'; Style-Header $ws $sec[1] 1 (@('Recurso', 'Rol', 'Cap. h/día') + $semanas + @('Total'))
    Style-Sub $ws ($sec[2] - 1) '3. Utilización'; Style-Header $ws $sec[2] 1 (@('Recurso', 'Rol', 'Ventana') + $semanas)
    Style-Sub $ws ($sec[3] - 1) '4. Alertas'; Style-Header $ws $sec[3] 1 @('Recurso', 'Máx. h/día', 'Días > cap', 'Semanas > 100 %', 'Día inicio', 'Día fin', 'Días con carga', 'Días sin carga en ventana', 'Estado')
    $colTot = 4 + $nS; $LT2 = Get-ColLetter ($colTot - 1)
    $finRec = @{}; $iniRec = @{}
    for ($i = 0; $i -lt $nR; $i++) {
        $r = $RECS[$i]; $pr = 2 + $i; $f1 = $sec[0] + 1 + $i; $f2 = $sec[1] + 1 + $i; $f3 = $sec[2] + 1 + $i; $f4 = $sec[3] + 1 + $i
        Set-Row $ws $f1 1 @($r.nombre, $r.rol, [int]$r.cap); Set-Row $ws $f2 1 @($r.nombre, $r.rol, [int]$r.cap); Set-Row $ws $f3 1 @($r.nombre, $r.rol)
        for ($s = 1; $s -le $nS; $s++) {
            $ca = Get-ColLetter (10 + 5 * ($s - 1) + 1); $cb = Get-ColLetter (10 + [Math]::Min(5 * $s, $ND))
            Set-Cell $ws $f1 (3 + $s) "=SUM($hP!${ca}${pr}:${cb}${pr})"
            $capS = 0; for ($d = 5 * ($s - 1) + 1; $d -le [Math]::Min(5 * $s, $ND); $d++) { $capS += [int]$r.capDia[$d - 1] }
            Set-Cell $ws $f2 (3 + $s) $capS
            $cl = Get-ColLetter (3 + $s); Set-Cell $ws $f3 (3 + $s) "=IFERROR(${cl}${f1}/${cl}${f2},0)"
        }
        Set-Cell $ws $f1 $colTot "=SUM(D${f1}:${LT2}${f1})"; Set-Cell $ws $f2 $colTot "=SUM(D${f2}:${LT2}${f2})"
        $ds = @($TAREAS | Where-Object { $_.recurso -eq $r.nombre -and $_.inicio -gt 0 }); $ini = ($ds | ForEach-Object { $_.inicio } | Measure-Object -Minimum).Minimum; $fin = ($ds | ForEach-Object { $_.fin } | Measure-Object -Maximum).Maximum
        $sc = 0; if ($ini) { for ($d = [int]$ini; $d -le [int]$fin; $d++) { if ([int]$r.capDia[$d - 1] -gt 0 -and @($ds | Where-Object { $_.dias.Contains([string]$d) }).Count -eq 0) { $sc++ } } }
        $rowP = "$hP!`$K`$${pr}:`$${LC}`$${pr}"
        Set-Cell $ws $f3 3 "=IFERROR($(Get-ColLetter $colTot)${f1}/(C${f1}*(F${f4}-E${f4}+1)),0)"
        Set-Row $ws $f4 1 @($r.nombre, "=MAX($rowP)", "=COUNTIF($rowP,`">`"&C$f1)", "=COUNTIF(D${f3}:${LT2}${f3},`">1`")",
            "=IFERROR(MATCH(TRUE,INDEX($rowP>0,0),0),`"`")", "=IFERROR(LOOKUP(2,1/($rowP>0),$hP!`$K`$1:`$${LC}`$1),`"`")", "=COUNTIF($rowP,`">0`")", $sc,
            "=IF(OR(B${f4}>C${f1},C${f4}>0,D${f4}>0),`"ALERTA: sobreasignación`",IF(H${f4}>0,`"Revisar vacíos`",`"OK`"))")
        $iniRec[$r.nombre] = "'Capacidad por recurso'!E$f4"; $finRec[$r.nombre] = "'Capacidad por recurso'!F$f4"
    }
    if ($nR -gt 0) {
        $rgU = $ws.Range("C$($sec[2] + 1):${LT2}$($sec[2] + $nR)"); $rgU.NumberFormat = '0%;-0%;"-"'
        Add-CF $rgU "=C$($sec[2] + 1)>1" '#FFC7CE' '#9C0006'; Add-CF $rgU "=C$($sec[2] + 1)>=0.8" '#C6EFCE' '#006100'; Add-CF $rgU "=C$($sec[2] + 1)>0" '#FFEB9C' '#9C5700'
        foreach ($s0 in $sec) { Style-Borders $ws ($s0 + 1) 1 ($s0 + $nR) $(if ($s0 -eq $sec[3]) { 9 } else { $colTot }) }
    }
    Set-Cell $ws ($sec[3] + $nR + 2) 1 'Notas: las ventanas de cada recurso son escalonadas (entra cuando hay trabajo y sale al terminar). Los días sin carga dentro de la ventana se justifican en Validaciones.'
    Widths $ws @(24, 22, 10); $ws.Range("D:$(Get-ColLetter $colTot)").ColumnWidth = 7

    # ---------- 6. Equipo ----------
    $ws = $HOJA['Equipo']
    $encEq = @('Recurso', 'Rol', 'Estado', 'Cap. h/día', 'Rol en el proyecto', 'Fases aplicables', 'Componentes / aplicativos asignados', 'Horas asignadas', 'Horas asumidas', 'Horas de contingencia', 'Día inicio', 'Día fin', 'Utilización en ventana')
    if ($PROY.tarifas) { $encEq += @('Tarifa', "Costo ($($PROY.moneda))") }
    $encEq += 'Racional de la asignación'
    Style-Title $ws "Equipo — $($PROY.nombre)" $encEq.Count
    Style-Header $ws 3 1 $encEq
    $fila = 4
    foreach ($r in $RECS) {
        $ts = @($TAREAS | Where-Object { $_.recurso -eq $r.nombre })
        $fases = @($ts | ForEach-Object { ($_.fase -split ' ')[0] } | Sort-Object -Unique) -join ', '
        $asignados = @($WBS.componentes | Where-Object { $_.desarrollador -eq $r.nombre } | ForEach-Object { $_.id }) -join ', '
        $vals = @($r.nombre, $r.rol, $(if ($r.propuesto) { 'Propuesto' } else { 'Confirmado' }), [int]$r.cap, (Texto 'Equipo' $r.nombre 'Rol en el proyecto' $r.rol), $fases, $asignados,
            "=SUMIF($rI,A$fila,$rJ)", "=SUMIFS($dhM,$dhO,A$fila,$dhR,`"Asumida`")", "=SUMIFS($dhM,$dhO,A$fila,$dhR,`"Contingencia`")",
            $(if ($iniRec.ContainsKey($r.nombre)) { "=$($iniRec[$r.nombre])" } else { '' }), $(if ($finRec.ContainsKey($r.nombre)) { "=$($finRec[$r.nombre])" } else { '' }),
            "=IFERROR(H$fila/(D$fila*(L$fila-K$fila+1)),0)")
        if ($PROY.tarifas) { $vals += @((To-Num $r.tarifa), "=H$fila*N$fila") }
        $defR = $(if ($ts.Count -eq 0) { 'Sin tareas: retirar o justificar' } elseif ($r.propuesto) { 'Dimensionado por el arnés con autorización del usuario' } else { 'Recurso de la entrada' })
        $vals += (Texto 'Equipo' $r.nombre 'Racional de la asignación' $defR)
        Set-Row $ws $fila 1 $vals; $fila++
    }
    Set-Cell $ws $fila 1 'Total'; Set-Cell $ws $fila 8 "=SUM(H4:H$($fila - 1))"; Set-Cell $ws $fila 9 "=SUM(I4:I$($fila - 1))"; Set-Cell $ws $fila 10 "=SUM(J4:J$($fila - 1))"
    if ($PROY.tarifas) { Set-Cell $ws $fila 15 "=SUM(O4:O$($fila - 1))" }
    $ws.Rows.Item($fila).Font.Bold = $true; Style-Borders $ws 4 1 $fila $encEq.Count; $ws.Range("M4:M$fila").NumberFormat = '0%'
    Widths $ws @(24, 22, 11, 8, 24, 20, 26, 10, 10, 10, 8, 8, 10, 10, 12, 40)

    # ---------- 9. Riesgos ----------
    $ws = $HOJA['Riesgos']
    Style-Title $ws "Riesgos y dependencias — $($PROY.nombre)" 12
    $encR = @('ID', 'Categoría', 'Riesgo / dependencia', 'Evidencia / disparador', 'Probabilidad', 'Impacto', 'Exposición', 'Estrategia', 'Mitigación / decisión requerida', 'Responsable', 'Hito', 'Estado')
    Style-Header $ws 3 1 $encR
    function Expo([string]$p, [string]$i) {
        if ($p -eq 'Alta' -and $i -eq 'Alta') { return 'Crítica' }
        if (($p -eq 'Alta' -and $i -eq 'Media') -or ($p -eq 'Media' -and $i -eq 'Alta')) { return 'Alta' }
        if (($p -eq 'Media' -and $i -eq 'Media') -or ($p -eq 'Alta' -and $i -eq 'Baja') -or ($p -eq 'Baja' -and $i -eq 'Alta')) { return 'Media' }
        return 'Baja'
    }
    $rsk = [ordered]@{}; $i = 0
    function Add-Rsk($cat, $txt, $evid, $p, $im, $estr, $mit, $resp, $hito) { $script:i++; $id = 'RSK-{0:D3}' -f $script:i; $script:rsk[$id] = @($id, $cat, $txt, $evid, $p, $im, (Expo $p $im), $estr, $mit, $resp, $hito, 'Abierto') }
    foreach ($x in $ENT.riesgos) { Add-Rsk (Get-Val $x 'Categoría' '') (Get-Val $x 'Riesgo / dependencia' '') (Get-Val $x 'Evidencia / disparador' 'Entrada') (Get-Val $x 'Probabilidad' 'Media') (Get-Val $x 'Impacto' 'Media') 'Mitigar' (Get-Val $x 'Mitigación propuesta' '') (Get-Val $x 'Responsable' '') '' }
    foreach ($c in $WBS.componentes) {
        if ($c.faltante) { Add-Rsk 'Alcance' "Información faltante en $($c.id): $($c.faltante)" 'Inventario' 'Media' 'Media' 'Mitigar' 'Validar con el cliente antes del análisis del componente' $WBS.lider 'G0' }
        if ($c.dependencia -eq 'Crítica') { Add-Rsk 'Dependencias' "Dependencia externa crítica en $($c.id)" 'Inventario' 'Alta' 'Alta' 'Mitigar' 'Confirmar fecha y responsable de la dependencia' $WBS.lider '' }
    }
    $hDev = @{}; $totDev = 0; foreach ($t in $TAREAS) { if ($t.fase -match '^0[2-5] ') { $hDev[$t.recurso] = [int]$hDev[$t.recurso] + [int]$t.horas; $totDev += [int]$t.horas } }
    foreach ($k in $hDev.Keys) { if ($totDev -gt 0 -and $hDev[$k] / $totDev -gt 0.4 -and $hDev.Count -gt 1) { Add-Rsk 'Capacidad' "Concentración de trabajo en $k ($([Math]::Round(100 * $hDev[$k] / $totDev)) % de 02-05)" 'Planning' 'Media' 'Alta' 'Mitigar' 'Evaluar un segundo recurso o redistribuir componentes' $WBS.lider '' } }
    if ($PROY.usaIA) { Add-Rsk 'AI-SDLC' 'Factores IA sin calibrar con datos reales' 'Parametros' 'Media' 'Media' 'Mitigar' 'Medir en el piloto (Ola 1) y recalibrar' $WBS.lider 'H1' }
    if ($PROY.certificacion -eq 'QA cliente') { Add-Rsk 'Calidad' 'Certificación a cargo del cliente: disponibilidad de su QA' 'Proyecto' 'Media' 'Alta' 'Transferir' 'Acordar calendario de certificación con el cliente' '' '' }
    if ($PROY.ambientes -notmatch 'Certificación|QA') { Add-Rsk 'Técnico' 'No se declaró ambiente de certificación/QA' 'Proyecto' 'Media' 'Alta' 'Mitigar' 'Confirmar ambientes antes de 06 Certificación' '' 'G0' }
    foreach ($id in (Get-Val $filasMan 'Riesgos' ([ordered]@{})).Keys) { $f = $filasMan['Riesgos'][$id]; $rsk[$id] = @($encR | ForEach-Object { [string](Get-Val $f $_ '') }); $rsk[$id][0] = $id }
    $fila = 4; foreach ($id in $rsk.Keys) { Set-Row $ws $fila 1 $rsk[$id]; $fila++ }
    if ($fila -gt 4) {
        $rgX = $ws.Range("G4:G$($fila - 1)")
        Add-CF $rgX '=G4="Crítica"' '#FF0000' '#FFFFFF'; Add-CF $rgX '=G4="Alta"' '#FFC7CE' '#9C0006'; Add-CF $rgX '=G4="Media"' '#FFEB9C' '#9C5700'; Add-CF $rgX '=G4="Baja"' '#C6EFCE' '#006100'
        Style-Borders $ws 4 1 ($fila - 1) 12
    }
    Widths $ws @(9, 14, 48, 22, 11, 10, 11, 11, 40, 20, 8, 10)

    # ---------- 10. Validaciones ----------
    $ws = $HOJA['Validaciones']
    Style-Title $ws "Validaciones — $($PROY.nombre)" 6
    Style-Header $ws 3 1 @('#', 'Validación', 'Resultado', 'Esperado', 'Estado', 'Detalle / justificación')
    $fila = 4
    $vivas = @{
        3  = "=COUNTBLANK($rI)"; 4 = "=COUNTIF($rJ,`"<=0`")+COUNTBLANK($rJ)"; 6 = "=COUNTBLANK($rB)+COUNTBLANK($rD)+COUNTBLANK($rE)"
        7  = "=SUM($rJ)-'Detalle de horas'!M$fT"; 10 = "=MAX($hP!K2:${LC}$($nR + 1))"; 11 = "=COUNTIF($rB,`"Transversal`")"
        12 = "=COUNTIF($(& $rP 'C'),`"01.1*`")*COUNTIF($(& $rP 'C'),`"01.2*`")"
    }
    $todos = @($VAL) + @([ordered]@{ n = 21; control = 'El archivo respeta la plantilla (10 hojas en orden, FechaInicio y Feriados)'; valor = 0; esperado = '0'; estado = 'OK'; detalle = '' })
    foreach ($v in @($todos | Sort-Object { [int]$_.n })) {
        $n = [int]$v.n
        $res = $(if ($vivas.ContainsKey($n)) { $vivas[$n] } else { $v.valor })
        $estado = switch ($n) {
            10 { "=IF(C$fila<=$($PROY.jornada),`"OK`",`"REVISAR`")" }
            11 { "=IF(C$fila>=8,`"OK`",`"REVISAR`")" }
            12 { "=IF(C$fila>0,`"OK`",`"REVISAR`")" }
            default { if ([string]$v.esperado -eq '0') { "=IF(C$fila=0,`"OK`",`"REVISAR`")" } else { [string]$v.estado } }
        }
        Set-Row $ws $fila 1 @($n, $v.control, $res, [string]$v.esperado, $estado, (Texto 'Validaciones' ([string]$n) 'Detalle / justificación' ([string]$v.detalle)))
        $fila++
    }
    $rgE = $ws.Range("E4:E$($fila - 1)")
    Add-CF $rgE '=E4="OK"' '#C6EFCE' '#006100'; Add-CF $rgE '=E4="REVISAR"' '#FFC7CE' '#9C0006'
    Style-Borders $ws 4 1 ($fila - 1) 6
    $fila += 1; Style-Sub $ws $fila 'Cobertura por componente y fase'; $fila++
    $fasesCob = @('02', '03', '04', '05'); if ($PROY.alcance -ne 'Certificación') { $fasesCob += '07' }
    Style-Header $ws $fila 1 (@('Componente') + $fasesCob); $fc0 = $fila + 1; $fila++
    foreach ($c in @($WBS.componentes | Where-Object { $_.tieneTareas -and $_.estrategia -ne 'Retirar' })) {
        Set-Cell $ws $fila 1 $c.id
        for ($j = 0; $j -lt $fasesCob.Count; $j++) { Set-Cell $ws $fila (2 + $j) "=COUNTIFS($(& $rP 'G'),`$A$fila&`" *`",$(& $rP 'C'),`"$($fasesCob[$j])*`")" }
        $fila++
    }
    if ($fila -gt $fc0) { $rgC = $ws.Range("B${fc0}:$(Get-ColLetter (1 + $fasesCob.Count))$($fila - 1)"); Add-CF $rgC "=B$fc0=0" '#FFC7CE' '#9C0006'; Add-CF $rgC "=B$fc0>0" '#C6EFCE' '#006100'; Style-Borders $ws $fc0 1 ($fila - 1) (1 + $fasesCob.Count) }
    Widths $ws @(12, 60, 12, 26, 11, 60)

    # ---------- KPIs del Resumen (dependen de las demás hojas) ----------
    $ws = $HOJA['Resumen ejecutivo']
    Style-Header $ws 3 4 @('KPI', 'Valor')
    $r0 = if ($PROY.claridad -eq 'Alta') { 0.15 } elseif ($PROY.claridad -eq 'Baja') { 0.5 } else { 0.3 }
    $kpi = New-Object System.Collections.Generic.List[object]
    $kpi.Add(@('Componentes', "=COUNTA('Estrategia de oleadas'!C4:C$fEO)"))
    foreach ($pl in @($WBS.componentes | ForEach-Object { $_.plataforma } | Sort-Object -Unique)) { $kpi.Add(@("  $pl", "=COUNTIF('Estrategia de oleadas'!F4:F$fEO,`"$pl`")")) }
    $kpi.Add(@('Olas', @($WBS.componentes | ForEach-Object { $_.ola } | Where-Object { $_ } | Sort-Object -Unique).Count))
    $kpi.Add(@('Duración (días hábiles)', "=$hP!B3")); $kpi.Add(@('Duración (semanas)', "=$hP!B4")); $kpi.Add(@('Horas totales', "=$hP!B2")); $kpi.Add(@('Días-persona (horas ÷ jornada)', '=ROUND(E{TOT}/$B$11,1)'))
    foreach ($o in @('Paramétrica', 'Entrada', 'Asumida', 'Contingencia')) { $kpi.Add(@("  Horas $o", "=SUMIF($dhR,`"$o`",$dhM)")) }
    $kpi.Add(@('Rango optimista (h)', '=ROUND(E{TOT}*0.9,0)')); $kpi.Add(@('Rango pesimista (h)', "=ROUND(E{TOT}*$(1 + $r0),0)"))
    $kpi.Add(@('Rango optimista (días-persona)', '=ROUND(E{TOT}*0.9/$B$11,1)')); $kpi.Add(@('Rango pesimista (días-persona)', "=ROUND(E{TOT}*$(1 + $r0)/`$B`$11,1)"))
    $kpi.Add(@('Máx. h de un recurso en un día', "=MAX($hP!K2:${LC}$($nR + 1))")); $kpi.Add(@('Días sin carga en ventanas (devs)', (@($VAL | Where-Object { $_.n -eq 15 })[0]).valor))
    $fTot = 0; for ($i = 0; $i -lt $kpi.Count; $i++) { if ($kpi[$i][0] -eq 'Horas totales') { $fTot = 4 + $i } }
    for ($i = 0; $i -lt $kpi.Count; $i++) { $v = $kpi[$i][1]; if ($v -is [string]) { $v = $v.Replace('{TOT}', [string]$fTot) }; Set-Row $ws (4 + $i) 4 @($kpi[$i][0], $v) }
    Style-Borders $ws 4 4 (3 + $kpi.Count) 5; $ws.Range("E4:E$(3 + $kpi.Count)").NumberFormat = '#,##0.#'
    Style-Header $ws 3 7 @('Recurso', 'Horas', '% participación', 'Día inicio', 'Día fin', 'Días-persona')
    $fila = 4
    foreach ($r in $RECS) { Set-Row $ws $fila 7 @($r.nombre, "=SUMIF($rI,G$fila,$rJ)", "=IFERROR(H$fila/H$(4 + $RECS.Count),0)", "=$($iniRec[$r.nombre])", "=$($finRec[$r.nombre])", "=ROUND(H$fila/`$B`$11,1)"); $fila++ }
    Set-Row $ws $fila 7 @('Total', "=SUM(H4:H$($fila - 1))", '', '', '', "=ROUND(H$fila/`$B`$11,1)"); $ws.Range("I4:I$fila").NumberFormat = '0%'; $ws.Range("L4:L$fila").NumberFormat = '#,##0.0'; Style-Borders $ws 4 7 $fila 12
    $fila = [Math]::Max($fila, [Math]::Max(3 + $kpi.Count, 15)) + 2
    Style-Sub $ws $fila 'Horas por componente'; $fila++
    Style-Header $ws $fila 1 @('Componente', 'Aplicativo origen', 'Ola', 'Plataforma', 'Aplicativo destino', 'Horas base', 'Horas finales'); $fila++
    $fi = $fila
    foreach ($c in $WBS.componentes) { Set-Row $ws $fila 1 @($c.id, $c.aplicativoOrigen, $c.ola, $c.plataforma, $c.aplicativo, "=SUMIF($dhB,A$fila,$dhK)", "=SUMIF($dhB,A$fila,$dhM)"); if ($colorDe.ContainsKey($c.aplicativoOrigen)) { Fill $ws $fila 2 $colorDe[$c.aplicativoOrigen] }; $fila++ }
    Style-Borders $ws $fi 1 ($fila - 1) 7
    $fila++; Style-Sub $ws $fila 'Esfuerzo por fase'; $fila++; Style-Header $ws $fila 1 @('Fase', 'Horas', 'Días-persona'); $fila++; $fi = $fila
    $rC = & $rP 'C'
    foreach ($fa in @($TAREAS | ForEach-Object { $_.fase } | Where-Object { $_ } | Sort-Object -Unique)) { Set-Row $ws $fila 1 @($fa, "=SUMIF($rC,A$fila,$rJ)", "=ROUND(B$fila/`$B`$11,1)"); $fila++ }
    Set-Row $ws $fila 1 @('Total', "=SUM(B${fi}:B$($fila - 1))", "=ROUND(B$fila/`$B`$11,1)"); $ws.Range("C${fi}:C$fila").NumberFormat = '#,##0.0'; Style-Borders $ws $fi 1 $fila 3; $fila++
    $fila++; Style-Sub $ws $fila 'Horas por aplicativo origen'; $fila++; Style-Header $ws $fila 1 @('Aplicativo origen', 'Horas'); $fila++; $fi = $fila
    foreach ($l in $LEY) { Set-Row $ws $fila 1 @($l.aplicativoOrigen, "=SUMIF($rD,A$fila,$rJ)"); Fill $ws $fila 1 $l.color; $fila++ }
    Style-Borders $ws $fi 1 ($fila - 1) 2
    $fila++; Style-Sub $ws $fila 'Horas por plataforma'; $fila++; Style-Header $ws $fila 1 @('Plataforma', 'Horas'); $fila++; $fi = $fila
    foreach ($pl in @($TAREAS | ForEach-Object { $_.plataforma } | Where-Object { $_ } | Sort-Object -Unique)) { Set-Row $ws $fila 1 @($pl, "=SUMIF($dhF,A$fila,$dhM)"); $fila++ }
    Style-Borders $ws $fi 1 ($fila - 1) 2
    $fila++; Style-Sub $ws $fila 'Hitos y gates (criterio de salida para pasar a la fase siguiente)'; $fila++; Style-Header $ws $fila 1 @('Gate', 'Criterio de salida', 'Día hábil', 'Fecha', 'Aprueba'); $fila++; $fi = $fila
    foreach ($h in $PLAN.hitos) { Set-Row $ws $fila 1 @("$($h.codigo) – $($h.nombre)", $h.criterio, [int]$h.dia, "=WORKDAY(FechaInicio,C$fila-1,Feriados)", [string]$h.aprueba); $fila++ }
    Set-Row $ws $fila 1 @('Fin del plan', 'Última tarea programada', $ND, "=WORKDAY(FechaInicio,C$fila-1,Feriados)", ''); $fila++
    $ws.Range("D${fi}:D$fila").NumberFormat = 'dd/mm/yyyy'; Style-Borders $ws $fi 1 ($fila - 1) 5
    $fila++; Style-Sub $ws $fila 'Supuestos y riesgos relevantes'; $fila++
    $vi = @(); foreach ($id in @($filasSup.Keys | Select-Object -First 4)) { $vi += "$id $($filasSup[$id][2])" }; foreach ($id in @($rsk.Keys | Select-Object -First 4)) { $vi += "$id $($rsk[$id][2])" }
    foreach ($x in $vi) { Set-Cell $ws $fila 1 "• $x"; $fila++ }
    Widths $ws @(34, 46, 4, 30, 12, 4, 24, 10, 12, 10, 10)

    # ---------- Cierre ----------
    $xl.Calculation = -4105
    $xl.Visible = $true; $xl.WindowState = -4143
    $wsP = $HOJA["Planning - $tipo"]; [void]$wsP.Activate(); [void]$wsP.Range("K$D1").Select(); $xl.ActiveWindow.FreezePanes = $true
    [void]$HOJA['Resumen ejecutivo'].Activate()
    $xl.Visible = $false
    if (Test-Path -LiteralPath $Salida) { Remove-Item -LiteralPath $Salida }
    $wb.SaveAs($Salida, 51)

    # Hojas propias del usuario: se copian desde el archivo anterior.
    $copiadas = @()
    if ($respaldo) {
        $wbOld = $xl.Workbooks.Open($respaldo, 0, $true)
        foreach ($w in $wbOld.Worksheets) {
            if ($hojas -contains $w.Name) { continue }
            $w.Copy([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)); $copiadas += $w.Name
        }
        $wbOld.Close($false)
    }
    $xl.CalculateFull()
    $wb.Save()

    # Verificación y snapshot (se lee del archivo guardado para comparar luego con lo que edite el usuario).
    $estados = @($HOJA['Validaciones'].Range("E4:E$(3 + $todos.Count)").Value2)
    $revisar = @($estados | Where-Object { $_ -eq 'REVISAR' }).Count
    $snap = [ordered]@{ archivo = $Salida; generado = (Get-Date).ToString('yyyy-MM-dd HH:mm'); hojasEstandar = $hojas; claves = [ordered]@{}; planning = [ordered]@{}; tablas = [ordered]@{}; filas = [ordered]@{} }
    foreach ($t in $TAREAS) { $snap.claves[$t.id] = $t.clave }
    $snap.planning = (Read-PlanningTareas $HOJA["Planning - $tipo"]).tareas
    $te = Get-TablasEditables
    foreach ($h in $te.Keys) {
        $snap.tablas[$h] = [ordered]@{}
        foreach ($f in (Read-TablaHasta $HOJA[$h] $te[$h].filaEnc $te[$h].clave)) { $o = [ordered]@{}; foreach ($col in $te[$h].editables + @('Horas base', 'Factor IA', 'Horas finales')) { if ($f.Contains($col)) { $o[$col] = $f[$col] } }; $snap.tablas[$h][[string]$f[$te[$h].clave]] = $o }
    }
    foreach ($h in (Get-HojasFilasManuales)) {
        $snap.filas[$h] = [ordered]@{}
        foreach ($f in (Read-TablaHasta $HOJA[$h] 3 'ID')) { $o = [ordered]@{}; foreach ($k in $f.Keys) { if ($k -ne '_fila') { $o[$k] = $f[$k] } }; $snap.filas[$h][[string]$f['ID']] = $o }
    }
    $wb.Close($true)
    $snap['hash'] = Get-FileHashSafe $Salida
    Write-JsonFile $snapRuta $snap
}
finally { Close-Excel $xl }

Write-Host "Excel generado: $Salida"
if ($copiadas.Count -gt 0) { Write-Host "Hojas del usuario conservadas: $($copiadas -join ', ')" }
Write-Host "Validaciones en REVISAR: $revisar"
if ($revisar -gt 0) { exit 4 }
exit 0
