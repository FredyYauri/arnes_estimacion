# Reglas de programación (Gantt diario)

El calendario se calcula con un script PowerShell generado en `_trabajo/programar.ps1` que lee `_trabajo/wbs.json` y escribe `_trabajo/planning.json`. **Nunca** asignar horas diarias manualmente.

## Reglas duras

1. Días = días hábiles numerados 1..N desde `FechaInicio`; fecha(d) = `WORKDAY(FechaInicio, d-1, Feriados)`. Fines de semana y feriados no existen en la numeración.
2. Carga diaria de un recurso ≤ su capacidad efectiva ([entradas](./entradas.md#hoja-recursos)) y ≤ Jornada. Vacaciones o fuera de disponibilidad = capacidad 0 (no cuenta como hueco).
3. Horas enteras por celda (1..cap). Sin decimales, negativos ni ceros escritos.
4. **Continuidad**: una tarea iniciada recibe ≥ 1 h en cada día en que su recurso tiene capacidad > 0, hasta terminar.
5. **Dependencias** finish-to-start ([catálogo](./catalogo-fases-tareas.md#dependencias-estándar-finish-to-start)). Misma persona: la sucesora puede empezar el mismo día si queda capacidad. Distinta persona: la sucesora empieza al día hábil siguiente.
6. Un recurso puede tener tareas en paralelo solo si la suma del día ≤ su capacidad (p. ej., soporte en certificación 4 h + análisis de la siguiente ola 2 h).
7. Construcción (fase 04) de cualquier componente no inicia antes de cerrar TRV-01, 02, 04, 05, 06 y 07, ni antes de la base común de su aplicativo destino.
8. Certificación: inicia el día hábil siguiente al último PU-02 del bloque y dura ≥ días_cert ([modelo-horas](./modelo-horas.md#certificación-despliegue-y-seguimiento-por-bloque)). No cae en fechas de "Congelamiento del cliente".
9. Despliegue: después de la certificación del bloque (o de todos los bloques si es Pase único) y dentro de una "Ventana de despliegue" si existen.
10. Seguimiento (hypercare) inicia el día hábil siguiente al pase y dura `Duración de hypercare`.
11. Restricciones de la entrada: "Dependencia externa" fija la fecha mínima de inicio de la construcción del componente; "Hito fijo" se valida al final (holgura o brecha); "Orden obligatorio" agrega una dependencia.

## Tipos de tarea

| Tipo | Ejemplos | Cómo se programa |
|---|---|---|
| `normal` | AN, DI, CO, PU, BC, TRV de análisis/diseño, DE-01..DE-03, SE-03, SE-01 | Cola por recurso; se asigna `min(restante, capacidad libre)` cada día |
| `tasa_fija` | CE-02, CE-03, CE-04, SE-02 | `h/día` constante durante `días` consecutivos desde que está lista. Si un día la capacidad libre no alcanza, ese día se asigna lo que haya (≥ 1) y la tarea se extiende un día (duración ≥ mínima) |
| `distribuida` | TRV-GES | Se programa al final: horas repartidas del día 1 al último día del proyecto (≥ 1 h/día; 2 h el primer día de cada semana y en días de hito hasta agotar las horas) |
| `relleno` | TRV-ACO | Al final: 1 h en cada día de la ventana activa del LT sin otra tarea (solo si el parámetro está activo) |

## Algoritmo (list scheduling por día)

```
calcular prioridad de cada tarea = (ruta crítica restante en días, ola, prioridad de negocio, posición en la serie del desarrollador)
para d = 1, 2, ... mientras queden tareas normales/tasa_fija pendientes:
    para cada recurso r:
        libre = capacidad(r, d)
        1) tareas tasa_fija de r activas o listas en d   → asignar su h/día (no exceder libre)
        2) tareas normales de r ya iniciadas (orden de inicio) → asignar min(restante, libre), mínimo 1 h cada una
        3) tareas normales de r listas (preds cumplidas según regla 5), por prioridad → asignar min(restante, libre) mientras libre > 0
    marcar terminadas las tareas con restante = 0 (registrar día fin)
al terminar: programar TRV-GES (distribuida) y TRV-ACO (relleno)
```

- Si en el paso 2 no alcanza 1 h por tarea iniciada, es un error de diseño: limitar el número de tareas que un recurso inicia en paralelo (máximo 2: una principal + una de tasa fija o de la ola siguiente).
- Serie por desarrollador: el AN-01 del siguiente componente de su serie depende del último CO del componente anterior (puede avanzar mientras espera la revisión CO-10 del LT); su CO-01/CO-02 depende del PU-02 del anterior.
- Si un recurso queda ocioso dentro de su ventana, intentar adelantar análisis/diseño de la ola siguiente asignada a él. No inventar tareas: si no hay nada, el día queda como "disponible" y se reporta en Capacidad por recurso.

## Balanceo y ajustes (iterar antes de generar el Excel)

- Fecha objetivo incumplida → (a) mover componentes del recurso más cargado a otro con la misma plataforma, (b) informar la brecha y recomendar recursos adicionales (solo se agregan si el usuario lo autoriza). Nunca subir la carga diaria por encima del tope.
- Ola que excede su duración máxima → partirla en dos olas.
- Recurso con > 20 % de días sin carga en su ventana → reasignar o acortar su ventana (entra cuando hay trabajo y sale al terminar).

## Formato de `planning.json`

```json
{
  "fechaInicio": "2026-11-02",
  "feriados": ["2026-12-08", "2026-12-25"],
  "ultimoDia": 84,
  "hitos": [{"codigo": "H0", "nombre": "Preparación completa", "dia": 10, "criterio": "..."}],
  "leyenda": [{"aplicativoOrigen": "Warrants-Autom.xlsm", "color": "#BDD7EE"}, {"aplicativoOrigen": "Transversal", "color": "#E7E6E6"}],
  "recursos": [{"nombre": "Líder Técnico 1", "rol": "Líder Técnico", "cap": 5, "propuesto": true}],
  "tareas": [
    {"id": "T001", "codigo": "TRV-GES", "ola": "Transversal", "fase": "00 Gestión", "aplicativoOrigen": "Transversal", "aplicativo": "Todos",
     "plataforma": "Transversal", "componente": "Proyecto", "tarea": "Gestión del proyecto, ...",
     "recurso": "Jefe de Proyecto 1", "horas": 95, "origen": "Paramétrica", "preds": [],
     "inicio": 1, "fin": 84, "dias": {"1": 2, "2": 1}}
  ]
}
```

IDs `T001..Tnnn` en el orden final del Planning: Transversal primero (por fase), luego por Ola → Aplicativo destino → Componente (orden de la serie) → Fase → día de inicio; al final los bloques de certificación/despliegue/seguimiento.

## Técnica (PowerShell 5.1)

- Usar `[ordered]@{}` y listas `System.Collections.Generic.List[object]`; `ConvertTo-Json -Depth 10`.
- Leer/escribir JSON con `[IO.File]::ReadAllText/WriteAllText(..., [Text.Encoding]::UTF8)`.
- Ejecutar scripts con tildes: `& ([scriptblock]::Create([IO.File]::ReadAllText($ruta,[Text.Encoding]::UTF8)))`.
- No nombrar funciones como alias existentes (`sp`, `sc`, `fc`, `gc`, `ls`, etc.).
