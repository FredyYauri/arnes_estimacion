# Excel desde PowerShell 5.1 (COM) — reglas técnicas

Entorno verificado: Windows, PowerShell 5.1, Excel en español, sin Python. Todo se hace con `Excel.Application` (COM).

## Lectura (solo lectura)

```powershell
$xl = New-Object -ComObject Excel.Application; $xl.Visible = $false; $xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open($ruta, 0, $true)          # ReadOnly = $true
$data = $ws.UsedRange.Value2                         # arreglo 2D 1-based: $data[$fila, $col]
$wb.Close($false); $xl.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
```
- Leer rangos completos con `UsedRange.Value2` (rápido); nunca celda por celda en bucles grandes.
- Fechas llegan como número serial OLE: `[DateTime]::FromOADate($v)`.
- Volcar a JSON (`_trabajo/entrada.json`) y trabajar sobre el JSON.

## Escritura

- Crear libro nuevo: `$wb = $xl.Workbooks.Add()`; ajustar n.º de hojas y nombres en orden; `SaveAs($ruta, 51)` (xlsx).
- Antes de `Validation.Add` o `FreezePanes`: `$xl.Visible = $true; $xl.WindowState = -4143` (si la ventana está minimizada fallan con 0x800A03EC).
- `0x800AC472` = Excel ocupado → reintentar con pausa corta.
- Números: `$cell.Formula = $n.ToString([Globalization.CultureInfo]::InvariantCulture)`. Textos: `$cell.Value2 = [string]$t`. No asignar `[double]` a `Value2` ni arreglos `object[,]` a `Range.Value2` (falla la conversión).
- Fórmulas: `Range.Formula` en **inglés** con separador `,` (SUM, SUMIF, WORKDAY, COUNTIF…). Para escribir muchas filas, asignar la fórmula de la primera fila a todo el rango (`$ws.Range("I19:I175").Formula = '=SUM(J19:DO19)'`) y Excel ajusta las referencias relativas.
- Escribir en bloque cuando sea posible: armar el texto TSV y pegar con `[Windows.Forms.Clipboard]::SetText($tsv)` + `$ws.Range('A19').PasteSpecial()` o `$ws.Paste()`, o escribir columna por columna con fórmula/valor. Para la matriz de horas diarias, escribir solo celdas > 0.
- Una función PowerShell que devuelve un `Range` debe usar `return , $ws.Range(...)` (si no, PowerShell lo enumera).
- No nombrar funciones como alias existentes (`sp`, `sc`, `fc`, `gc`, `ls`, `cp`…): el alias gana y el script queda esperando entrada.
- Variables no distinguen mayúsculas: `$m` pisa a `$M`.

## Formatos condicionales

- `FormatConditions.Add(2, 0, $formula)` (xlExpression) interpreta la fórmula en el **idioma local**. Usar fórmulas **sin funciones** para que funcionen en cualquier idioma: comparaciones y productos, p. ej. `=($D19="Warrants-Autom.xlsm")*(K19>0)`, `=K2>$G2`, `=K15<>""`, `=G4="Crítica"`.
- Si inevitablemente se necesita una función, usar nombres en español y `;` (p. ej. `=Y(IZQUIERDA($C19;2)="04";J19>0)`), o mejor una columna auxiliar.
- Para colores: `$fc.Interior.Color` y `$fc.Font.Color` usan BGR → convertir `#RRGGBB` con `R + 256*G + 65536*B`.

## Nombres y validación de datos

- `Names.Add('FechaInicio', "='Resumen ejecutivo'!`$B`$9")` (sin funciones; si llevan funciones, usar `;`).
- Listas desplegables: `Validation.Add(3, 1, 1, "=lstRoles")` apuntando a un nombre definido sobre un rango (evita separadores de lista locales).

## Presentación

- Fijar anchos de columna **antes** de `AutoFit` de filas (si no, filas altísimas).
- Leyendas en filas fijas: calcular la fila de encabezados después de la leyenda para no pisarla.
- `TEXT(fecha,"mmm")` depende del idioma: usar `CHOOSE(MONTH(...),"ene",...)` si se necesitan meses.
- Captura visual para revisión: `Range.CopyPicture(1,2)` + `[Windows.Forms.Clipboard]::GetImage().Save(png)`.

## Cierre y verificación

```powershell
$xl.CalculateFull(); $wb.Save(); $wb.Close($true); $xl.Quit()
[Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
```
Reabrir en solo lectura y leer la hoja Validaciones (columna Estado) para el informe final. Verificar que no quedan procesos `EXCEL` huérfanos creados por el script.

## Ejecución de scripts con tildes

`& ([scriptblock]::Create([IO.File]::ReadAllText($ruta,[Text.Encoding]::UTF8)))` — guardar los .ps1 en UTF-8.
