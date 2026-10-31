# Catálogo de fases, tareas y WBS

## Fases (orden obligatorio)

| Código | Fase | Etiqueta en Planning (columna Fase) | Color de referencia (no se usa en el Gantt) |
|---|---|---|---|
| 00 | Gestión | `00 Gestión` | #808080 (texto blanco) |
| 01.1 | Preparación – Entendimiento inicial | `01.1 Preparación – Entendimiento inicial` | #FFC400 |
| 01.2 | Preparación – Relevamiento y preparación técnica | `01.2 Preparación – Relevamiento y preparación técnica` | #FFC400 |
| 02 | Análisis | `02 Análisis` | #00CB5D |
| 03 | Diseño | `03 Diseño` | #00DFED |
| 04 | Construcción | `04 Construcción` | #0072BC (texto blanco) |
| 05 | Pruebas unitarias | `05 Pruebas unitarias` | #FF7A00 (texto blanco) |
| 06 | Certificación | `06 Certificación` | #7030A0 (texto blanco) |
| 07 | Despliegue | `07 Despliegue` | #C00000 (texto blanco) |
| 08 | Seguimiento en producción | `08 Seguimiento en producción` | #375623 (texto blanco) |

## Estructura de desglose (WBS)

```
Proyecto
└── Oleada / Entregable            (Transversal | Ola 1..n | Pase/Bloque de despliegue)
    └── Aplicativo destino         (Todos | <aplicativo>)
        └── Componente             (módulo / programa / requerimiento / aplicativo origen, de cualquier tecnología legacy | Base común <aplicativo> | TRV)
            └── Fase               (00..08)
                └── Tarea          (código del catálogo + nombre + "(<componente>)")
```

En Planning: columna B = Ola, C = Fase, D = Aplicativo origen (define el color de la barra y la leyenda), E = Aplicativo destino, F = Plataforma, G = Componente, H = Tarea, I = Recurso.

## Tareas transversales (una sola vez por proyecto) — Ola = "Transversal", Aplicativo = "Todos"

| Código | Fase | Tarea | Rol por defecto | Origen |
|---|---|---|---|---|
| TRV-GES | 00 | Gestión del proyecto, backlog, seguimiento, riesgos, reportes y coordinación | Jefe de Proyecto (o Líder Técnico si no hay) | Paramétrica |
| TRV-ACO | 00 | Acompañamiento técnico (solo si parámetro activo; 1 h en días sin otra tarea del LT) | Líder Técnico | Asumida |
| TRV-ENT | 01.1 | Kick-off y entendimiento funcional/técnico del alcance, dependencias, integraciones y prioridades | Líder Técnico | Asumida |
| TRV-REL | 01.2 | Relevamiento técnico: dependencias, accesos, permisos, restricciones y confirmación de la estrategia de oleadas | Líder Técnico | Asumida |
| TRV-AMB | 01.2 | Ambientes, accesos, repositorios y pipelines por aplicativo destino | Especialista de Plataformas (o Developer Sr) | Asumida |
| TRV-IA | 01.2 | Configurar el asistente IA por repositorio: instrucciones, prompts de migración y reglas de sanitización (solo si Uso IA = Sí) | Developer Sr | Asumida |
| TRV-01 | 02 | Relevar usuarios, roles y permisos actuales de los componentes | Developer (o Analista Calidad Sr) | Paramétrica |
| TRV-02 | 02 | Relevar roles y perfiles de los aplicativos destino | Developer | Paramétrica |
| TRV-06 | 03 | Arquitectura de referencia: patrones de integración, errores, configuración, secretos y observabilidad | Arquitecto de Aplicaciones (o Líder Técnico) | Paramétrica |
| TRV-05 | 03 | Eventos de auditoría y trazabilidad | Líder Técnico | Paramétrica |
| TRV-07 | 03 | Modelo de tablas comunes de soporte (auditoría, parámetros, bitácora) | Líder Técnico / Modelador Sr | Paramétrica |
| TRV-04 | 03 | Matriz de roles por pantalla y operación | Líder Técnico | Paramétrica |
| TRV-08 | 03 | Estrategia de carga inicial y tratamiento del histórico (si algún componente migra datos) | Arquitecto de Datos / Data Engineer Sr / Líder Técnico | Paramétrica |
| TRV-QA | 03 | Estrategia y plan de pruebas de certificación (si QA proveedor o Mixta) | Analista Calidad Sr | Paramétrica |
| TRV-SEC | 03 | Threat modeling y lineamientos de desarrollo seguro (si datos sensibles) | Analista de hacking Sr/Medium | Paramétrica |
| TRV-UX | 03 | Guía UX/UI y sistema de diseño (si aplicativo destino nuevo o ≥10 pantallas en total) | Especialista UX/UI Sr | Paramétrica |

Prerrequisitos: TRV-01, TRV-02, TRV-04, TRV-05, TRV-06 y TRV-07 terminan antes de cualquier tarea de 04 Construcción. TRV-08 termina antes del primer despliegue con datos.

## Base común por aplicativo destino

Se crea cuando un aplicativo destino recibe ≥2 componentes o es **Nuevo**. Ola = la primera ola que usa ese aplicativo. Componente = `Base común <aplicativo>`.

| Código | Fase | Tarea |
|---|---|---|
| BC-01 | 02 | Analizar impacto sobre el aplicativo destino |
| BC-02 | 03 | Definir estructura común y habilitación de funciones por perfil y por ola |
| BC-03 | 04 | Construir estructura común (navegación, layout, configuración, logging) |
| BC-04 | 04 | Implementar modelo base de permisos y auditoría |
| BC-05 | 04 | Revisión de código de la base común (Líder Técnico) |
| BC-06 | 05 | Pruebas unitarias de la base común |

Dependencia: BC-03/BC-04 terminan antes de la construcción de los componentes de ese aplicativo; BC-06 antes de sus pruebas unitarias.

## Tareas por componente (estrategia Migrar / Refactorizar / Reconstruir / Reemplazar por estándar)

Nombre en Planning: `<Tarea> (<ID componente>)`.

| Código | Fase | Tarea | Rol por defecto | Condición |
|---|---|---|---|---|
| AN-01 | 02 | Análisis del componente origen: extraer/sanitizar código, revisar lógica, objetos ocultos, integraciones y datos | Developer asignado | Siempre |
| AN-02 | 02 | Validar con negocio las reglas existentes | Developer (Alta: Líder Técnico) | Siempre |
| AN-03 | 02 | Analizar el aplicativo destino y el impacto | Developer asignado | Destino existente sin base común |
| AN-04 | 02 | Análisis de datos: fuentes, calidad, volumen e histórico | Developer / Data Engineer | Migración de datos ≠ Ninguna |
| DI-01 | 03 | Diseñar la solución técnica | Developer (Alta: Líder Técnico o Arquitecto) | Siempre |
| DI-02 | 03 | Diseñar prototipos de pantallas en el destino | Developer o Especialista UX/UI | Pantallas > 0 |
| DI-03 | 03 | Diseñar modelo de datos, mapeo y carga | Developer / Modelador | Tablas > 0 |
| CO-01 | 04 | Preparar el aplicativo destino para recibir la funcionalidad | Developer asignado | Destino sin base común |
| CO-02 | 04 | Implementar la lógica de negocio, cálculos y procesos | Developer asignado | Siempre |
| CO-03 | 04 | Implementar pantallas y navegación | Developer asignado | Pantallas > 0 |
| CO-04 | 04 | Migrar estructuras de datos y parametrización (tablas, archivos, configuración) | Developer asignado | Tablas > 0 |
| CO-05 | 04 | Implementar integraciones y reemplazar automatizaciones del origen | Developer asignado | Integraciones > 0 |
| CO-06 | 04 | Implementar reportes y salidas | Developer (Power BI si aplica) | Reportes > 0 |
| CO-07 | 04 | Manejo explícito de errores y externalización de rutas y secretos | Developer asignado | Siempre |
| CO-08 | 04 | Implementar permisos y auditoría | Developer asignado | Requiere permisos y auditoría = Sí |
| CO-09 | 04 | Scripts de migración/carga de datos e histórico | Developer / Data Engineer | Migración de datos ≠ Ninguna |
| CO-CT | 04 | Contingencia (x %) – motivo | Developer asignado | Contingencia > 0 |
| CO-10 | 04 | Revisar código y ejecutar análisis estático | Líder Técnico | Siempre |
| PU-01 | 05 | Preparar casos y datos de prueba unitaria contra resultados del legacy | Developer asignado | Siempre |
| PU-02 | 05 | Implementar y ejecutar pruebas unitarias y corregir | Developer asignado | Siempre |
| DE-01 | 07 | Manual de usuario y runbook del componente | Developer asignado | Alcance ≥ Producción |
| SE-01 | 08 | Retiro del componente legacy (desactivar, archivar y comunicar) | Developer asignado | Parámetro "Incluir retiro del legacy" = Sí y alcance = Hypercare |

## Adaptar el catálogo a cada migración

El catálogo es la **base**, no una lista cerrada: ni este catálogo ni las tareas del Excel de ejemplo tienen que repetirse igual en todas las migraciones. Según el pedido, los documentos y la tecnología origen/destino, el agente ajusta el WBS en `asignacion.json`:

- `tareasExcluidas`: tareas del catálogo que no aplican (p. ej. `"M01|DI-02"` si no hay rediseño de pantallas), cada una con su motivo. Se informan en Supuestos como "Tarea excluida". No se pueden excluir TRV-ENT, TRV-GES ni la contingencia.
- `tareasExtra`: tareas propias de la migración que el catálogo no cubre (p. ej. conversión de JCL a pipelines, migración de reportes Crystal, pruebas de performance). Cada una con componente (o `TRV`), código, fase, tarea, horas, recurso, origen (`Asumida` o `Entrada`) y motivo. Sin predecesoras explícitas, va después de la última tarea del componente en esa fase o antes.

Toda tarea extra o excluida debe poder explicarse con el pedido o con un dato de la entrada; si no, se pregunta.

Se permite agrupar en una sola fila tareas del mismo componente, fase y recurso (p. ej. AN-01+AN-03, CO-02+CO-03+CO-04, PU-01+PU-02) para reducir filas; la trazabilidad por código queda en la hoja "Detalle de horas" (columna ID planning).

Estrategia **Retirar**: solo AN-02 (confirmar retiro con negocio) y SE-01. Estrategia **Mantener**: sin tareas (se lista en Estrategia de oleadas con 0 h).
Tipo **Mejora evolutiva**: usar las mismas tareas; LOC = líneas a crear o modificar; AN-01 se reduce a la mitad (no hay código origen que extraer).
Tipo **Upgrade** (versión de framework/plataforma): AN-01 = inventario de dependencias incompatibles; CO-02 = adecuación de código (LOC afectadas); agregar PU de regresión (PU-02 × 1,5).

## Tareas por bloque de certificación/despliegue

Bloque = ola, aplicativo destino de la ola, o pase único según "Agrupación de certificación y despliegue".

| Código | Fase | Tarea | Rol | Condición |
|---|---|---|---|---|
| CE-01 | 06 | Diseño de casos de certificación | Analista Calidad Ssr | QA proveedor / Mixta |
| CE-02 | 06 | Ejecución de certificación, registro de defectos y evidencias | Analista Calidad Ssr / Tester | QA proveedor / Mixta |
| CE-03 | 06 | Soporte en certificación y corrección de defectos | Developer(s) del bloque | Siempre |
| CE-04 | 06 | Coordinación de certificación, congelamiento y gestión de defectos críticos | Líder Técnico / Analista Calidad Sr | Siempre |
| CE-05 | 06 | Pruebas de seguridad y validación de remediaciones | Analista de hacking | Datos sensibles = Sí |
| DE-02 | 07 | Plan de pase y comunicación | Jefe de Proyecto | Alcance ≥ Producción |
| DE-03 | 07 | Preparar y ensayar paquetes, scripts DDL y carga inicial | Especialista de Plataformas / Líder Técnico | Alcance ≥ Producción |
| DE-04 | 07 | Ejecutar el pase y validar post-despliegue | Especialista de Plataformas + Developer | Alcance ≥ Producción |
| SE-02 | 08 | Hypercare: estabilización, monitoreo y soporte inicial | Developer(s) + Líder Técnico | Alcance = Hypercare |
| SE-03 | 08 | Transferencia a soporte y mantenimiento | Líder Técnico | Alcance = Hypercare |

## Dependencias estándar (finish-to-start)

```
TRV-ENT → TRV-REL / TRV-AMB / TRV-IA → (análisis de componentes puede iniciar desde el 2.º día de 01.2)
TRV-01, TRV-02, TRV-04, TRV-05, TRV-06, TRV-07 → toda tarea 04
BC-02 → BC-03 → BC-04 → BC-05 → BC-06 ;  BC-04 → CO-xx de componentes del aplicativo
Por componente: AN-* → DI-* → CO-01 → CO-02..CO-09 → CO-CT → CO-10 → PU-01 → PU-02
PU-02 de todos los componentes del bloque → CE-01..CE-05 → DE-02..DE-04 → SE-02 → SE-03 / SE-01
DE-01 puede ejecutarse en paralelo a la certificación del bloque
```

PU-01 puede comenzar mientras el LT ejecuta CO-10 (preparar casos), pero PU-02 inicia después de CO-10.
