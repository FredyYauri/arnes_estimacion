# Scripts del arnés

Scripts fijos y versionados que ejecutan el cálculo de la estimación. El agente **no escribe ni modifica scripts** en cada corrida: solo prepara `asignacion.json` (decisiones de criterio) y ejecuta estos pasos. Así, la misma entrada produce siempre el mismo resultado.

Requisitos: Windows, PowerShell 5.1 y Excel instalado (COM). Ejecutar desde esta carpeta o con ruta completa; los .ps1 están en UTF-8.

## Flujo

| Paso | Script | Lee | Escribe |
|---|---|---|---|
| 1 | `1-Leer-Entrada.ps1 -Entrada <entrada.xlsx>` | entrada (solo lectura) | `_trabajo/entrada.json`, `_trabajo/faltantes.json` (sale con código 2 si faltan datos) |
| 2 | el agente arma `_trabajo/asignacion.json` | entrada.json | asignacion.json |
| 3 | `2-Calcular-Estimacion.ps1 -Trabajo <_trabajo>` | entrada, asignacion, ajustes | `wbs.json` (complejidad, WBS, horas) |
| 4 | `3-Programar.ps1 -Trabajo <_trabajo>` | wbs, entrada, ajustes | `planning.json` (Gantt diario) |
| 5 | `4-Validar-Planning.ps1 -Trabajo <_trabajo>` | planning, wbs | `validaciones.json`, `validaciones.txt` (código 4 si hay REVISAR) |
| 6 | `6-Generar-Excel.ps1 -Trabajo <_trabajo> -Salida <estimacion.xlsx>` | todo lo anterior | Excel final, `snapshot.json`, respaldo en `backup/` |

`5-Detectar-Cambios.ps1` se usa cuando el usuario editó el Excel a mano (ver abajo). `6-Generar-Excel.ps1` lo ejecuta solo antes de sobrescribir.

Si el paso 5 da REVISAR, se corrige la causa (asignación, orden, ola) en `asignacion.json` y se repiten los pasos 3 a 5. Nunca se editan `wbs.json` ni `planning.json` a mano.

## Cambios manuales en el Excel (no se pierden)

Cada vez que se genera el Excel queda un `snapshot.json` con lo que se escribió. Antes de sobrescribir, `6-Generar-Excel.ps1` compara el archivo con ese snapshot:

- **Sin cambios del usuario** → regenera (con respaldo en `backup/`).
- **Con cambios no incorporados** → se **detiene** sin tocar nada (código 5) y deja la lista en `_trabajo/cambios_manuales.json`. El agente muestra la lista al usuario y le pregunta:
  - **Conservar** → `5-Detectar-Cambios.ps1 -Trabajo ... -Archivo ... -Aplicar` guarda los cambios en `ajustes.json`; luego se repiten los pasos 3 a 6 y los cambios aparecen en el nuevo Excel.
  - **Descartar** → `6-Generar-Excel.ps1 ... -DescartarCambios` (el archivo editado queda en `backup/`).

Qué se conserva con `-Aplicar`:

| Dónde editó el usuario | Qué se guarda en `ajustes.json` | Efecto al regenerar |
|---|---|---|
| Planning: horas por día o recurso de una tarea | `fijadas[<clave>]` | La tarea queda fija tal cual; el resto se programa alrededor. Las validaciones avisan si rompe una regla |
| Planning: solo el recurso | `tareas[<clave>].recurso` | Se reprograma con el nuevo recurso |
| Detalle de horas: Horas finales | `tareas[<clave>].horas` (origen `Entrada`) | Se reprograma con esas horas |
| Motivo de la secuencia, criterios de salida, racional, justificaciones, observaciones | `textos[<hoja>][<clave>][<columna>]` | Se vuelven a escribir |
| Filas nuevas o editadas en Supuestos o Riesgos | `filasManuales[<hoja>][<ID>]` | Se vuelven a escribir con su ID |
| Hojas nuevas creadas por el usuario | (nada) | Se copian tal cual al archivo nuevo |

No se conservan: filas nuevas en Planning (el trabajo nuevo se agrega en la entrada), formatos y cambios en celdas calculadas. `5-Detectar-Cambios.ps1` los informa como "NO SE CONSERVA" para que el usuario decida.

Recomendación al usuario: para notas libres, crear una hoja propia (se conserva siempre).

## `asignacion.json` (lo arma el agente)

Contiene solo decisiones de criterio que el modelo no puede derivar solo. Todo lo demás sale de la entrada.

```json
{
  "numeroOlas": 1,
  "componentes": {
    "M01": { "ola": "Ola 1", "desarrollador": "Ana", "orden": 1, "motivo": "Piloto de la plataforma Web .NET", "revisor": "Luis", "componenteIA": false },
    "APP2": { "ola": "Ola 1", "desarrolladores": ["Ana", "Pedro", "Rosa"], "orden": 2, "motivo": "Aplicativo grande (40 módulos) en 3 lotes en paralelo" }
  },
  "roles": {
    "Jefe de Proyecto": "Carla", "Líder Técnico": "Luis", "Arquitecto de Aplicaciones": "Luis",
    "Especialista de Plataformas": "Pedro", "Developer Sr": "Ana", "Developer": "Ana",
    "Analista Calidad Sr": "Rosa", "Analista Calidad Ssr": "Rosa", "Seguridad": "", "UX": "", "Datos": ""
  },
  "propuestos": [
    { "nombre": "Developer .NET Ssr 1", "rol": "Developer .NET Ssr", "plataformas": "Web .NET", "dedicacion": 100 }
  ],
  "reutilizables": [ { "actividad": "Servicio de notificaciones", "componentes": "M01, M04", "aprovechamiento": "M01 lo construye; M04 lo reutiliza" } ],
  "tareasExcluidas": [ { "clave": "M01|DI-02", "motivo": "Las pantallas se migran sin rediseño" } ],
  "tareasExtra": [
    { "componente": "M02", "codigo": "EX-01", "fase": "04", "tarea": "Convertir jobs JCL a pipelines", "horas": 24, "recurso": "Ana", "origen": "Asumida", "motivo": "Pedido del cliente" }
  ],
  "criterios": [ "Piloto por plataforma en la Ola 1", "Componentes Alta con desarrollador dedicado desde el inicio" ]
}
```

- `numeroOlas`: cuántas olas pidió el usuario. Si no se mencionaron olas, `1` (una sola entrega, todo en "Ola 1"). Manda sobre "Modelo de entrega" / "Número de olas" de la entrada.
- `componentes`: obligatorio para cada componente con tareas (`desarrollador` debe existir en Recursos o en `propuestos`). `ola` ("Ola 1".."Ola n") puede omitirse si viene "Ola sugerida" o si es una sola entrega.
- Filas con "Módulos (n)" > 1 se parten en lotes `<ID>-L1..Lk`. `desarrolladores` reparte los lotes entre esas personas (uno por persona, en rueda); `lotes` fija la cantidad. Sin ninguno de los dos: ⌈n ÷ "Módulos por lote"⌉ con el `desarrollador` indicado. Un lote puede tener su propia entrada (`"APP2-L3": {...}`).
- `tareasExcluidas` / `tareasExtra`: adaptan el catálogo a esta migración (ver [catálogo](../references/catalogo-fases-tareas.md#adaptar-el-catálogo-a-cada-migración)). `predecesoras` es opcional en las extra.
- `gates`: opcional, solo si el pedido ajusta los gates del Resumen (G0 Inicio, G1 Análisis cerrado, G2 Diseño aprobado, G3 Listo para certificar, G4 Certificado, G5 Pase, G6 Cierre). Por código: `{ "G1": { "criterio": "...", "aprueba": "..." }, "G2": { "omitir": true } }`; `"G3"` aplica a todos los bloques, `"G3.2"` solo al bloque 2.
- `roles`: opcional; si falta un rol, el script busca en Recursos según el catálogo (p. ej. sin Jefe de Proyecto, la gestión la toma el Líder Técnico).
- `propuestos`: solo con autorización explícita del usuario (quedan marcados "Propuesto").

## `ajustes.json` (cambios manuales del usuario)

Lo escribe `5-Detectar-Cambios.ps1 -Aplicar`; también puede editarse a pedido del usuario. Las claves de tarea son estables entre regeneraciones: `TRV-GES`, `<ID componente>|CO-02`, `BC|<aplicativo>|BC-03`, `<bloque>|CE-02`, etc. (columna "ID origen" de Detalle de horas).

```json
{
  "tareas":  { "M01|CO-02": { "horas": 30, "motivo": "estimado por el equipo" } },
  "fijadas": { "M02|PU-02": { "dias": { "41": 4, "42": 4 }, "recurso": "Ana" } },
  "textos":  { "Estrategia de oleadas": { "M01": { "Motivo de la secuencia": "Pedido del cliente" } } },
  "filasManuales": { "Riesgos": { "RSK-050": { "ID": "RSK-050", "Categoría": "Datos", "Riesgo / dependencia": "..." } } },
  "aplicadoDesde": "<hash del Excel cuyos cambios ya se incorporaron>",
  "historial": []
}
```

Los días de `fijadas` son días hábiles del proyecto (1 = fecha de inicio): si cambia la fecha de inicio o el calendario, revise las tareas fijadas.

## Códigos de salida

0 = OK · 1 = hay cambios manuales (5) · 2 = faltan datos en la entrada (1) o no hay snapshot (5) · 3 = tareas sin programar (3) · 4 = validaciones en REVISAR (4 y 6) · 5 = generación detenida para proteger cambios manuales (6).
