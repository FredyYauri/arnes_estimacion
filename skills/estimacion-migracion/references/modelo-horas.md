# Modelo paramétrico de horas

Horas finales por tarea = `ROUNDUP( HorasBase × FactorIA , 0 )` (entero, mínimo 1 si la tarea aplica).
Las reglas de % (CO-07, CO-10, PU-xx, CE-01) se calculan sobre **horas base** (antes del factor IA).
Los valores de la hoja `Parametros` de la entrada prevalecen sobre esta tabla (buscar por nombre exacto).

## Parámetros por defecto

| Parámetro (nombre exacto) | Defecto |
|---|---|
| Productividad LOC/h (sin IA) | 45 |
| Factor complejidad Baja / Media / Alta | 0,8 / 1,0 / 1,3 |
| Horas por pantalla Baja / Media / Alta | 4 / 8 / 16 |
| Horas por integración Baja / Media / Alta | 6 / 12 / 24 |
| Horas por reporte Baja / Media / Alta | 2 / 4 / 8 |
| Horas por tabla Baja / Media / Alta | 1 / 2 / 4 |
| Factor IA Análisis | 0,5 |
| Factor IA Diseño | 0,75 |
| Factor IA Construcción | 0,6 |
| Factor IA Revisión de código | 1,0 |
| Factor IA Preparar casos PU | 0,9 |
| Factor IA Ejecutar PU | 0,5 |
| Factor IA Manuales | 0,4 |
| Factor IA Otros | 1,0 |
| % Gestión del proyecto | 8 |
| % Errores y secretos | 10 |
| % Revisión de código | 5 |
| % Preparar casos PU | 5 |
| % Ejecutar PU | 15 |
| Contingencia conocido / integración / requerimientos poco claros / dependencia externa / componente IA | 3 / 8 / 15 / 15 / 15 |
| Redondeo a bloques de 4 h | No |
| Carga diaria máxima | Capacidad productiva del rol |
| Acompañamiento LT en días libres | No |
| Soporte dev en certificación (h/día) | 4 |
| Ejecución QA (h/día) | 6 |
| Coordinación LT en certificación (h/día) | 2 |
| Hypercare dev (h/día por aplicativo) | 2 |
| Hypercare LT (h/día) | 2 |
| Duración máxima de ola (semanas) | 8 |
| Incluir retiro del legacy | Sí |

Las filas con "Baja / Media / Alta" corresponden en la hoja `Parametros` a tres parámetros separados (`Factor complejidad Baja`, `Factor complejidad Media`, …); las contingencias, a `Contingencia conocido`, `Contingencia integración`, etc.
Si "Uso de asistente IA" = No, todos los factores IA = 1,0.

## Tareas por componente

`c` = complejidad (B/M/A); `fc` = factor complejidad; `n_pant`, `n_int`, `n_rep`, `n_tab` = conteos del Inventario.

| Código | Horas base | Factor IA |
|---|---|---|
| AN-01 | B 4 / M 8 / A 16; + 1 h por cada 1.000 LOC sobre 3.000 | Análisis |
| AN-02 | B 2 / M 3 / A 6 | Otros |
| AN-03 | B 2 / M 4 / A 8 | Análisis |
| AN-04 | Simple 4 / Compleja 12 | Análisis |
| DI-01 | B 3 / M 8 / A 16 | Diseño |
| DI-02 | n_pant × (B 1 / M 2 / A 3) | Diseño |
| DI-03 | n_tab × (B 0,5 / M 1 / A 2); + 4 si migración Compleja | Diseño |
| CO-01 | B 8 / M 16 / A 24 | Construcción |
| CO-02 | LOC / Productividad × fc (LOC es bloqueante en la entrada) | Construcción |
| CO-03 | n_pant × Horas por pantalla(c) | Construcción |
| CO-04 | n_tab × Horas por tabla(c) | Construcción |
| CO-05 | n_int × Horas por integración(c) | Construcción |
| CO-06 | n_rep × Horas por reporte(c) | Construcción |
| CO-07 | % Errores y secretos × Σ(CO-02..CO-06), mín. 1 | Construcción |
| CO-08 | B 4 / M 8 / A 12 | Construcción |
| CO-09 | Simple 8 / Compleja 24 | Construcción |
| CO-10 | % Revisión de código × Σ(CO-01..CO-09), mín. 1 | Revisión de código |
| PU-01 | % Preparar casos PU × Σ(CO-01..CO-09), mín. 2 | Preparar casos PU |
| PU-02 | % Ejecutar PU × Σ(CO-01..CO-09), mín. 2 (Upgrade: × 1,5) | Ejecutar PU |
| DE-01 | B 2 / M 4 / A 8 | Manuales |
| SE-01 | B 1 / M 2 / A 4 | Otros |

**Horas override** del Inventario: reemplaza Σ(02..05) del componente; repartir proporcionalmente a las horas paramétricas finales; origen `Entrada`.

## Base común por aplicativo destino

| Código | Horas base | Factor IA |
|---|---|---|
| BC-01 | 8 | Análisis |
| BC-02 | 8 | Diseño |
| BC-03 | 40 (aplicativo Nuevo: 80) | Construcción |
| BC-04 | 24 | Construcción |
| BC-05 | 3 | Revisión de código |
| BC-06 | 10 | Ejecutar PU |

## Transversales

`n_comp` = componentes con tareas; `n_app` = aplicativos destino distintos.

| Código | Horas base (factor IA = 1,0 salvo indicación) | Origen |
|---|---|---|
| TRV-ENT | 16 + 2 × n_comp (máx. 40) | Asumida |
| TRV-REL | 16 + 4 × n_app (máx. 40) | Asumida |
| TRV-AMB | 8 × n_app | Asumida |
| TRV-IA | 4 × n_app (solo si Uso IA = Sí) | Asumida |
| TRV-01 | 4 + 0,5 × n_comp | Paramétrica |
| TRV-02 | 4 × n_app | Paramétrica |
| TRV-06 | 16 (+ 8 si hay ≥ 3 plataformas destino) | Paramétrica |
| TRV-05 | 8 | Paramétrica |
| TRV-07 | 8 | Paramétrica |
| TRV-04 | 8 + 0,5 × n_comp | Paramétrica |
| TRV-08 | 8 + 2 × (componentes con migración de datos) | Paramétrica |
| TRV-QA | 8 + 1 × n_comp | Paramétrica |
| TRV-SEC | 16 | Paramétrica |
| TRV-UX | 24 | Paramétrica |
| TRV-GES | % Gestión × Σ(horas finales de todas las demás tareas). Si resulta < 1 h por día hábil del proyecto, completar hasta 1 h/día y marcar la diferencia como Asumida | Paramétrica |
| TRV-ACO | 1 h por día hábil sin otra tarea del Líder Técnico dentro de su ventana (solo si "Acompañamiento LT en días libres" = Sí) | Asumida |

## Certificación, despliegue y seguimiento (por bloque)

`H04` = horas base de 04 Construcción de los componentes del bloque; `n_dev` = desarrolladores del bloque; `n_app_b` = aplicativos del bloque.

- **Días de certificación** = máx(Duración mínima de certificación, ROUNDUP(0,2 × H04 / Ejecución QA (h/día))).
- CE-01 = 10 % × H04 (mín. 4). CE-02 = días_cert × Ejecución QA (h/día).
- CE-03 = días_cert × Soporte dev en certificación (h/día), por cada desarrollador del bloque.
- CE-04 = días_cert × Coordinación LT en certificación (h/día).
- CE-05 = 8 + 2 × componentes del bloque (si datos sensibles).
- DE-02 = 8 por pase. DE-03 = 8 × n_app_b. DE-04 = 4 × n_app_b (Plataformas) + 2 por desarrollador.
- SE-02 = días_hypercare × (Hypercare dev × n_app_b + Hypercare LT). SE-03 = 8 + 2 × n_app_b.

Todas con origen Paramétrica y factor IA Otros. Si "Certificación a cargo de" = QA cliente, no generar CE-01/CE-02 (registrar como decisión declarada en la entrada).

## Contingencia

Por componente, % = **el mayor** de los que apliquen (no se suman), con tope 20 %:

| Condición | % |
|---|---|
| Trabajo conocido y repetitivo (Reutilización Alta y Riesgo técnico Bajo) | Contingencia conocido |
| Destino existente o Integraciones > 0 | Contingencia integración |
| "Información faltante" con texto, Reglas = Alto o Claridad de requerimientos del proyecto = Baja | Contingencia requerimientos poco claros |
| Dependencia externa Parcial (10) o Crítica (Contingencia dependencia externa) | 10 / 15 |
| Plataforma o integración con componente de IA nuevo | Contingencia componente IA |
| Ninguna de las anteriores | Contingencia conocido |

Horas de contingencia = ROUNDUP(% × Σ horas finales 02-05 del componente). Se planifican en una fila propia `CO-CT Contingencia (x %) – <motivo> (<ID>)` en 04 Construcción, con el mismo recurso, antes de CO-10. Origen `Contingencia`.

## Rango de la estimación

- Optimista = Total × 0,90.
- Pesimista = Total × (1 + r), con r = 0,15 (claridad Alta), 0,30 (Media), 0,50 (Baja).
Mostrar ambos en el Resumen ejecutivo junto con el total base.

## Ejemplo de cálculo (verificación rápida)

Componente Media, LOC 1.438, 2 pantallas, 2 integraciones, 3 tablas, sin migración de datos, IA = Sí:
- CO-02 base = 1438/45 × 1,0 = 31,96 → × 0,6 = 19,2 → **20 h**
- CO-03 base = 2 × 8 = 16 → × 0,6 → **10 h**; CO-05 base = 2 × 12 = 24 → **15 h**; CO-04 base = 3 × 2 = 6 → **4 h**
- CO-07 base = 10 % × (31,96 + 16 + 6 + 24) = 7,8 → × 0,6 → **5 h**
- Σ base CO-01..CO-09 (con CO-01 = 16 y CO-08 = 8) ≈ 109,8 → CO-10 = 5,5 → **6 h**; PU-01 = 5,5 × 0,9 → **5 h**; PU-02 = 16,5 × 0,5 → **9 h**

## Calibración

Al cerrar cada proyecto, comparar horas reales vs estimadas por código de tarea y complejidad, y ajustar `Productividad LOC/h`, horas por pantalla/integración y factores IA en la hoja `Parametros` de la plantilla de entrada (registrar fecha y proyecto fuente en la descripción del parámetro).
