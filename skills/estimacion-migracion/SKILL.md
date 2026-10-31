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
| [Modelo de salida](./assets/Ejemplo_proyecto.xlsx) | Ejemplo real del Excel final (10 hojas): referencia de estructura, fórmulas, estilos y formatos condicionales. Solo lectura |
| [Entradas](./references/entradas.md) | Campos bloqueantes, políticas precargadas, derivados y controles de la entrada (nada se supone) |
| [Catálogo de fases y tareas](./references/catalogo-fases-tareas.md) | Fases 00-08, tareas base, WBS y dependencias estándar |
| [Modelo de horas](./references/modelo-horas.md) | Fórmulas paramétricas, factor IA, redondeo, contingencia y origen de horas |
| [Complejidad y priorización](./references/complejidad-priorizacion.md) | Puntaje de complejidad y armado de oleadas |
| [Roles y capacidad](./references/roles-capacidad.md) | Catálogo de roles, capacidad diaria, matriz de asignación y equipo propuesto |
| [Reglas de programación](./references/reglas-programacion.md) | Algoritmo del Gantt diario y reglas duras |
| [Plantilla de salida](./references/plantilla-salida.md) | Hojas, columnas, fórmulas y estilos del Excel final (réplica de `Ejemplo_proyecto.xlsx`) |
| [Validaciones](./references/validaciones.md) | Checklist obligatorio y cómo calcular cada control |
| [Excel COM](./references/excel-com.md) | Reglas técnicas para leer/escribir Excel desde PowerShell 5.1 |

El modelo de salida es [Ejemplo_proyecto.xlsx](./assets/Ejemplo_proyecto.xlsx) (en `assets/`). Solo lectura: sus datos son de otro proyecto y nunca se copian a una estimación nueva.

## Estructura de trabajo por proyecto

```
Estimaciones/<slug-proyecto>/
├── fuentes/                     # información del proyecto aportada por el usuario (relevamientos, inventarios, planes)
├── entrada_<slug>.xlsx          # copia de la entrada completada (no se modifica)
├── estimacion_<slug>.xlsx       # salida final (siempre el mismo nombre)
├── backup/                      # versiones previas: estimacion_<slug>_<yyyyMMdd_HHmm>.xlsx
└── _trabajo/                    # entrada.json, wbs.json, planning.json, scripts generados, log de validaciones
```

Si la salida ya existe, cópiala a `backup/` con sello de fecha antes de regenerarla (no sobrescribir respaldos existentes).

## Procedimiento

1. **Preparar carpeta** con la estructura anterior. Copiar la entrada a `entrada_<slug>.xlsx` si viene de otra ubicación.
2. **Leer la entrada** en modo solo lectura (Excel COM, ver [excel-com](./references/excel-com.md)) y volcarla a `_trabajo/entrada.json`.
3. **Validar la entrada** según [entradas](./references/entradas.md): bloqueantes, políticas, listas, coherencia de fechas y recursos. Cualquier faltante → devolver la tabla de faltantes y **detenerse**; ningún dato del proyecto se supone.
4. **Clasificar complejidad** de cada componente del Inventario con [complejidad-priorizacion](./references/complejidad-priorizacion.md). Respetar "Complejidad forzada" si viene informada.
5. **Construir el WBS** con [catalogo-fases-tareas](./references/catalogo-fases-tareas.md): Proyecto → Oleada → Aplicativo destino → Componente → Fase → Tarea. Incluir siempre: preparación (01.1/01.2), transversales (02/03), base común por aplicativo destino cuando aplique, y certificación/despliegue/seguimiento según la agrupación elegida. Componentes con estrategia *Retirar* solo generan tareas de retiro; *Mantener* no genera tareas.
6. **Calcular horas** con [modelo-horas](./references/modelo-horas.md): horas base → factor IA → redondeo hacia arriba a entero → contingencia (fila separada). Etiquetar el origen de cada hora: `Paramétrica`, `Entrada`, `Asumida` o `Contingencia`. Los valores de la hoja `Parametros` de la entrada prevalecen sobre los de la referencia.
7. **Asignar** los recursos de la entrada con [roles-capacidad](./references/roles-capacidad.md). Si Recursos no cubre lo necesario, preguntar; solo con autorización explícita dimensionar el equipo y marcarlo "Propuesto".
8. **Definir oleadas** con [complejidad-priorizacion](./references/complejidad-priorizacion.md): piloto y quick wins primero, componentes grandes temprano con recurso dedicado, agrupación por aplicativo destino y componentes reutilizables.
9. **Programar** con [reglas-programacion](./references/reglas-programacion.md): generar `_trabajo/programar.ps1`, ejecutarlo y obtener `_trabajo/planning.json` (horas enteras por tarea y día hábil). Nunca asignar a mano.
10. **Pre-validar** el `planning.json` con los controles de [validaciones](./references/validaciones.md) (capacidad, dependencias, continuidad, fases, totales). Si algo falla, corregir la causa (reasignar, reordenar o mover la oleada) y volver al paso 9.
11. **Generar el Excel** `estimacion_<slug>.xlsx` según [plantilla-salida](./references/plantilla-salida.md) y [excel-com](./references/excel-com.md): 10 hojas, fórmulas vivas (SUMIF/WORKDAY/COUNTIF), nombres `FechaInicio` y `Feriados`, formatos condicionales por fase y por carga.
12. **Validar el archivo final**: reabrir, recalcular (`Application.CalculateFull()`), leer la hoja Validaciones y confirmar que todo está en "OK" o justificado. Guardar el log en `_trabajo/validaciones.txt`.
13. **Informar** al usuario: KPIs, rango de la estimación, equipo, oleadas e hitos, validaciones, supuestos clave, información faltante y preguntas abiertas.

## Reglas que nunca se rompen

- Ningún recurso supera su carga diaria máxima (por defecto, la capacidad productiva del rol; tope absoluto 8 h).
- Una tarea iniciada se ejecuta de corrido hasta terminar (sin días vacíos en su ejecución).
- Ninguna tarea inicia antes de que terminen sus predecesoras.
- Toda funcionalidad incluye análisis, diseño, construcción y pruebas unitarias.
- Cada oleada (o bloque de despliegue) incluye certificación (mínimo configurado, por defecto 5 días hábiles), despliegue y seguimiento, salvo que el alcance llegue solo hasta certificación.
- Ningún dato del proyecto se supone: si falta, se pide. "Asumida" solo identifica horas que agrega el arnés por regla (preparación 01.1/01.2, ajustes de gestión), nunca datos de la entrada.
- La preparación (01.1 y 01.2) existe siempre, aunque no esté en la entrada, y se marca como Asumida.
- Las actividades transversales no se duplican por componente: se planifican una vez y se reutilizan.
- Horas enteras, sin negativos ni vacíos; cada tarea con responsable, fase, aplicativo/entregable y horas.
