# Validaciones obligatorias

Se ejecutan dos veces: **(a)** sobre `planning.json` antes de escribir el Excel (script) y **(b)** como fórmulas vivas en la hoja Validaciones del archivo final. La estimación no se entrega con un "REVISAR" sin justificación en la columna Detalle.

Notación: `P` = hoja Planning, filas de tarea `D1:N`, días `K:LC` (D = aplicativo origen, E = aplicativo destino, I = recurso, J = horas); `DH` = Detalle de horas; `EO` = Estrategia de oleadas.

| # | Validación (checklist del arnés) | Resultado (fórmula en el Excel) | Esperado |
|---|---|---|---|
| 1 | Todas las funcionalidades están incluidas | `=COUNTA(EO!C4:C<n>)` | = n.º de filas del Inventario |
| 2 | Cada funcionalidad tiene análisis, diseño, construcción y pruebas unitarias | `=COUNTIF(Validaciones!<matriz 02..05>,0)` (excluye Retirar/Mantener) | 0 |
| 3 | Cada tarea tiene responsable | `=COUNTBLANK(P!I D1:I N)` | 0 |
| 4 | Cada tarea tiene horas > 0 | `=COUNTIF(P!J D1:J N,"<=0")+COUNTBLANK(P!J D1:J N)` | 0 |
| 5 | No hay horas negativas, decimales ni inconsistentes | `=SUMPRODUCT(--(P!K D1:LC N<0))+SUMPRODUCT(--(P!K D1:LC N<>INT(P!K D1:LC N)))` | 0 |
| 6 | Toda tarea tiene ola, aplicativo origen y aplicativo/entregable destino | `=COUNTBLANK(P!B D1:B N)+COUNTBLANK(P!D D1:D N)+COUNTBLANK(P!E D1:E N)` | 0 |
| 7 | Total Planning = Total Detalle de horas | `=SUM(P!J D1:J N)-SUM(DH!<Horas finales>)` | 0 |
| 8 | Horas agregadas identificadas | `=SUMIF(DH!<Origen>,"Asumida",<HF>)+SUMIF(DH!<Origen>,"Contingencia",<HF>)` vs filas con prefijo `[Asumida]`/`[Contingencia]` en P | Iguales |
| 9 | No hay recursos sobre su capacidad diaria | `=SUMPRODUCT(--(P!K2:LC(R+1)>P!G2:G(R+1)))` | 0 |
| 10 | Máximo de horas de un recurso en un día | `=MAX(P!K2:LC(R+1))` | ≤ Jornada |
| 11 | Actividades transversales identificadas | `=COUNTIF(P!B D1:B N,"Transversal")` | ≥ 8 |
| 12 | Preparación presente (01.1 y 01.2) | `=COUNTIF(P!C D1:C N,"01.1*")*COUNTIF(P!C D1:C N,"01.2*")` | > 0 |
| 13 | Certificación por bloque con duración mínima | n.º de bloques con días de 06 ≥ días_cert (calculado por el script; valor) | = n.º de bloques |
| 14 | Las olas coinciden con lo pedido | `=COUNTA('Matriz de entregables'!A4:A<n>)` | = n.º de olas pedidas (1 = entrega única) |
| 15 | Días sin carga de desarrolladores dentro de su ventana | `=SUM('Capacidad por recurso'!<Días sin carga devs>)` | ≤ 10 % de días-ventana (justificar) |
| 16 | Dependencias incumplidas | valor del script (tareas que inician antes del fin de una predecesora) | 0 |
| 17 | Orden de fases por componente (02 → 03 → 04 → 05 → 06 → 07 → 08) | valor del script | 0 |
| 18 | Tareas con huecos en su ejecución (continuidad) | valor del script | 0 |
| 19 | Construcciones antes de cerrar los transversales prerrequisito | valor del script | 0 |
| 20 | Fecha fin ≤ fecha objetivo (si existe) | `=IF(<objetivo>="","N/A",<fin><=<objetivo>)` | VERDADERO o justificar brecha |
| 21 | El archivo respeta la plantilla | 10 hojas en orden, nombres `FechaInicio` y `Feriados` existen (script) | OK |
| 22 | Cada aplicativo origen tiene color y entrada en la leyenda | n.º de aplicativos origen distintos en P!D sin regla de color (script) | 0 |

Estado: `=IF(<condición>,"OK","REVISAR")`. Para resultados de script, escribir el valor numérico en Resultado y dejar la fórmula de Estado.

## Controles del script de pre-validación (sobre `planning.json`)

```
para cada tarea: horas == suma(dias); horas entero > 0; recurso no vacío; fase válida
para cada recurso y día: suma ≤ cap(recurso, día) y ≤ jornada; 0 si vacaciones
para cada tarea: días con horas consecutivos en días hábiles con capacidad (sin huecos)
para cada dependencia (a → b): inicio(b) > fin(a)  (o ≥ si mismo recurso y hay capacidad)
para cada componente: fin(fase k) ≤ inicio(fase k+1)
construcción: inicio(CO/BC-03) > fin(TRV-01,02,04,05,06,07)
certificación: días(06 del bloque) ≥ días_cert ; inicio(06) > fin(PU-02 del bloque)
totales: Σ planning == Σ detalle ; Σ por origen consistente
```

Los controles los ejecuta [`scripts/4-Validar-Planning.ps1`](../scripts/4-Validar-Planning.ps1). Guardar el resultado en `_trabajo/validaciones.txt` (control, valor, esperado, estado) y resumirlo en la respuesta final.
