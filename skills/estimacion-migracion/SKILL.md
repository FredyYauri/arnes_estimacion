---
name: estimacion-migracion
description: 'Arnés de estimación para migraciones, upgrades y mejoras evolutivas de aplicaciones legacy (VBA/Excel, Access, VB6, scripts, sistemas legados) hacia .NET web/desktop, APIs, Power Platform, SharePoint, Power BI, móvil o datos. Úsalo para: leer la plantilla de entrada, clasificar complejidad, construir el WBS por oleadas y fases 00-08, calcular horas con modelo paramétrico y factor IA, asignar roles y capacidad, programar el Gantt diario sin sobreasignación, generar el Excel de estimación con la plantilla estándar (Ejemplo_proyecto.xlsx) y ejecutar las validaciones obligatorias.'
argument-hint: 'Ruta de la entrada completada (Plantilla_Entrada_Estimacion.xlsx)'
---

# Estimación de migraciones y upgrades

**Objetivo:** datos de entrada + reglas de estimación + capacidad de recursos + plantilla de salida = estimación ordenada, realista y reutilizable.

**Alcance:** mejoras evolutivas, migraciones y upgrades.

## Recursos del arnés

| Recurso | Uso |
|---|---|
| [Plantilla de entrada](./assets/Plantilla_Entrada_Estimacion.xlsx) | Se entrega al usuario para completar. Hojas: Instrucciones, Proyecto, Inventario, Recursos, Calendario, Restricciones, Riesgos, Parametros, Catalogos |
| [Modelo de salida](./assets/Ejemplos%20Salidas/Ejemplo_proyecto.xlsx) | Ejemplo real del Excel final (10 hojas): única referencia de estructura (hojas, columnas, nombres definidos), fórmulas, estilos y formatos condicionales. Solo lectura |
| [Entradas](./references/entradas.md) | Campos bloqueantes, políticas precargadas, derivados y controles de la entrada (nada se supone) |
| [Catálogo de fases y tareas](./references/catalogo-fases-tareas.md) | Fases 00-08, tareas base, WBS y dependencias estándar |
| [Modelo de horas](./references/modelo-horas.md) | Fórmulas paramétricas, factor IA, redondeo, contingencia y origen de horas |
| [Complejidad y priorización](./references/complejidad-priorizacion.md) | Puntaje de complejidad y armado de oleadas |
| [Roles y capacidad](./references/roles-capacidad.md) | Catálogo de roles, capacidad diaria, matriz de asignación y equipo propuesto |
| [Reglas de programación](./references/reglas-programacion.md) | Algoritmo del Gantt diario y reglas duras |
| [Validaciones](./references/validaciones.md) | Checklist obligatorio y cómo calcular cada control |
| [Excel COM](./references/excel-com.md) | Reglas técnicas para leer/escribir Excel desde PowerShell 5.1 |
| [Scripts](./scripts/README.md) | Scripts fijos que calculan, programan, validan y generan el Excel, y protegen los cambios manuales del usuario |

El modelo de salida es [Ejemplo_proyecto.xlsx](./assets/Ejemplos%20Salidas/Ejemplo_proyecto.xlsx) (en `assets/Ejemplos Salidas/`). Es la fuente de verdad de la estructura del Excel final: se replican sus hojas, columnas, fórmulas, nombres definidos y formatos. Solo lectura: sus datos son de otro proyecto y nunca se copian a una estimación nueva.

## Estructura de trabajo por proyecto

```
Estimaciones/<slug-proyecto>/
├── fuentes/                     # información del proyecto aportada por el usuario (relevamientos, inventarios, planes)
├── entrada_<slug>.xlsx          # copia de la entrada completada (no se modifica)
├── estimacion_<slug>.xlsx       # salida final (siempre el mismo nombre)
├── backup/                      # versiones previas: estimacion_<slug>_<yyyyMMdd_HHmm>.xlsx
└── _trabajo/                    # entrada.json, asignacion.json, ajustes.json, wbs.json, planning.json, validaciones, snapshot.json
```

Todo el cálculo lo hacen los [scripts del arnés](./scripts/README.md); el agente no escribe scripts propios ni asigna horas a mano. Si la salida ya existe, `6-Generar-Excel.ps1` la respalda en `backup/` y **se detiene si el usuario la editó a mano** hasta que decida conservar o descartar esos cambios.

## Procedimiento

1. **Preparar carpeta** con la estructura anterior. Copiar la entrada a `entrada_<slug>.xlsx` si viene de otra ubicación.
2. **Leer y validar la entrada**: `scripts/1-Leer-Entrada.ps1` la lee en solo lectura, la vuelca a `_trabajo/entrada.json` y aplica los controles de [entradas](./references/entradas.md).
3. **Faltantes**: si el script sale con código 2, devolver la tabla de `_trabajo/faltantes.json` (Hoja · Fila/ID · Campo · Problema) y **detenerse**; ningún dato del proyecto se supone. Si informa cobertura de recursos insuficiente, preguntar si esperar o dimensionar el equipo.
4. **Asignar** los recursos de la entrada con [roles-capacidad](./references/roles-capacidad.md). Si Recursos no cubre lo necesario, preguntar; solo con autorización explícita dimensionar el equipo y marcarlo "Propuesto".
5. **Definir olas** con [complejidad-priorizacion](./references/complejidad-priorizacion.md): si el usuario pidió n olas, repartir los componentes en exactamente n olas (piloto y quick wins primero, componentes grandes temprano con recurso dedicado, agrupación por aplicativo destino y componentes reutilizables); si no lo mencionó, es una sola entrega y todo va en la Ola 1.
6. **Guardar las decisiones** de los pasos 4 y 5 en `_trabajo/asignacion.json` (formato en [scripts](./scripts/README.md#asignacionjson-lo-arma-el-agente)): número de olas (`numeroOlas`), ola, desarrollador y orden de cada componente, responsables por rol, recursos Propuestos autorizados, actividades reutilizables y criterios de secuenciación. Es el único paso de criterio del agente.
7. **Calcular** con `scripts/2-Calcular-Estimacion.ps1`: clasifica la complejidad ([complejidad-priorizacion](./references/complejidad-priorizacion.md), respetando "Complejidad forzada"), arma el WBS ([catalogo-fases-tareas](./references/catalogo-fases-tareas.md): preparación 01.1/01.2, transversales, base común, certificación/despliegue/seguimiento según la agrupación; *Retirar* solo genera retiro y *Mantener* no genera tareas) y calcula horas ([modelo-horas](./references/modelo-horas.md): base → factor IA → redondeo → contingencia en fila separada, con origen `Paramétrica`, `Entrada`, `Asumida` o `Contingencia`). La hoja `Parametros` de la entrada prevalece sobre la referencia.
8. **Ajustes del usuario**: si existe `_trabajo/ajustes.json` (cambios manuales que el usuario decidió conservar), los scripts 2 y 3 los aplican automáticamente.
9. **Programar**: `scripts/3-Programar.ps1` aplica [reglas-programacion](./references/reglas-programacion.md) y escribe `_trabajo/planning.json` (horas enteras por tarea y día hábil). Nunca asignar a mano.
10. **Pre-validar**: `scripts/4-Validar-Planning.ps1` ejecuta los controles de [validaciones](./references/validaciones.md). Si algo queda en REVISAR, corregir la causa en `asignacion.json` (reasignar, reordenar o mover la oleada) y volver al paso 7.
11. **Generar el Excel**: `scripts/6-Generar-Excel.ps1` escribe `estimacion_<slug>.xlsx` replicando la estructura del [modelo de salida](./assets/Ejemplos%20Salidas/Ejemplo_proyecto.xlsx): 10 hojas, fórmulas vivas (SUMIF/WORKDAY/COUNTIF), nombres `FechaInicio` y `Feriados`, formatos condicionales por fase y por carga. Si sale con código 5, el usuario editó el Excel: mostrarle `_trabajo/cambios_manuales.json` y preguntar si conserva (`5-Detectar-Cambios.ps1 -Aplicar` y repetir pasos 7 a 11) o descarta (`-DescartarCambios`). Nunca regenerar sobre cambios del usuario sin su respuesta.
12. **Validar el archivo final**: el script recalcula, lee la hoja Validaciones e informa cuántas quedan en REVISAR; toda REVISAR debe quedar justificada en la columna Detalle.
13. **Informar** al usuario: KPIs, rango de la estimación, equipo, oleadas e hitos, validaciones, supuestos clave, información faltante y preguntas abiertas.

## Reglas que nunca se rompen

- Ningún recurso supera su carga diaria máxima (por defecto, la Jornada del proyecto: 8 h/día, configurable).
- Una tarea iniciada se ejecuta de corrido hasta terminar (sin días vacíos en su ejecución).
- Ninguna tarea inicia antes de que terminen sus predecesoras.
- Toda funcionalidad incluye análisis, diseño, construcción y pruebas unitarias.
- Cada oleada (o bloque de despliegue) incluye certificación (mínimo configurado, por defecto 5 días hábiles), despliegue y seguimiento, salvo que el alcance llegue solo hasta certificación.
- Ningún dato del proyecto se supone: si falta, se pide. "Asumida" solo identifica horas que agrega el arnés por regla (preparación 01.1/01.2, ajustes de gestión), nunca datos de la entrada.
- La preparación (01.1 y 01.2) existe siempre, aunque no esté en la entrada, y se marca como Asumida.
- Las actividades transversales no se duplican por componente: se planifican una vez y se reutilizan.
- Horas enteras, sin negativos ni vacíos; cada tarea con responsable, fase, aplicativo/entregable y horas.
