# Roles, capacidad y asignación

## Catálogo de roles

**Carga diaria por defecto** = la Jornada del proyecto (8 h, configurable en Proyecto → "Jornada (h/día)") para todos los recursos, ajustada por la Dedicación % y la Capacidad h/día de cada recurso si se informa. Solo si `Carga diaria máxima` = "Capacidad productiva del rol" se usa la columna **Cap. planificable** de esta tabla como tope por rol (valor superior del rango sugerido). Nunca superar la jornada.

| Rol | Actividades principales | Cap. diaria sugerida | Cap. planificable (h/día) | Modalidad recomendada |
|---|---|---|---|---|
| Jefe de Proyecto *(agregado por el arnés)* | Planificación, seguimiento, riesgos, reportes, comunicación con stakeholders, coordinación de pases | 2 a 4 h | 4 | Continua, todo el proyecto |
| Líder Técnico | Liderazgo técnico, entendimiento, análisis técnico, diseño, revisión de código, bloqueos, estándares, validación de despliegues | 4 a 5 h | 5 | Continua, 50-70 % |
| Arquitecto de Aplicaciones | Arquitectura, patrones, componentes reutilizables, integraciones, atributos de calidad, decisiones de alto impacto | 3 a 4 h | 4 | Preparación, análisis, diseño; revisiones por hito |
| Especialista de Plataformas | Ambientes, accesos, permisos, infraestructura, despliegues, monitoreo, automatización | 4 a 5 h | 5 | Puntual: ambientes, despliegue, monitoreo |
| Developer .NET Sr | Componentes complejos, APIs, integraciones, revisión de código, mentoría, incidencias críticas, PU | 6 h | 6 | Continua en construcción |
| Developer .NET Ssr | Funcionalidades, APIs, integraciones medias, PU, defectos, documentación | 6 h | 6 | Continua en construcción |
| Developer .NET Nearshore | Funcionalidades .NET bajo coordinación del equipo | 6 h | 6 | Continua; +10 % coordinación (asumida) |
| Analista Calidad Sr | Estrategia de pruebas, cobertura, riesgos, coordinación QA, defectos críticos | 4 a 5 h | 5 | Desde análisis; mayor en certificación |
| Analista Calidad Ssr | Diseño y ejecución de casos, defectos, evidencias | 5 a 6 h | 6 | Alta en certificación y regresión |
| Tester | Pruebas manuales, regresión, evidencias | 6 h | 6 | Pruebas, certificación, estabilización |
| Developer Power Apps Sr / Ssr | Power Apps, Power Automate, Dataverse, integraciones | 6 h | 6 | Continua en diseño/construcción |
| Developer SharePoint Sr / Ssr | Sitios, listas, bibliotecas, permisos, flujos | 6 h | 6 | Continua |
| Developer Power BI Sr | Modelo semántico, DAX complejo, RLS, optimización | 5 a 6 h | 6 | Intensa en análisis/modelado |
| Developer Power BI Ssr | Reportes, dashboards, DAX, validación | 6 h | 6 | Continua en construcción |
| Especialista UX Sr / Ssr | Journeys, arquitectura de información, wireframes, usabilidad | 4-5 / 5-6 h | 5 / 6 | Análisis y diseño |
| Especialista UI Sr / Ssr | Sistema de diseño, guía visual, pantallas, responsive | 4-5 / 5-6 h | 5 / 6 | Diseño; revisiones en construcción |
| Developer iOS Sr / Ssr | Desarrollo iOS, APIs, publicación | 6 h | 6 | Continua |
| Developer Android Sr / Ssr | Desarrollo Android, APIs, publicación | 6 h | 6 | Continua |
| Arquitecto de Datos Sr | Arquitectura, gobierno, seguridad y calidad de datos | 3 a 4 h | 4 | Análisis y diseño; hitos |
| Modelador Sr | Modelo conceptual/lógico/físico, diccionario de datos | 5 a 6 h | 6 | Análisis y diseño |
| Data Engineer Sr / Medium | Pipelines, ETL/ELT, cargas, calidad | 6 h | 6 | Continua en construcción de datos |
| Data Engineer Junior | Tareas de datos baja/media complejidad bajo supervisión | 5 h | 5 | Bajo supervisión |
| Analista de hacking / desarrollo seguro Jr | Validaciones básicas, seguimiento de hallazgos | 4 a 5 h | 5 | Puntual en certificación |
| Analista de hacking / desarrollo seguro Medium | Pruebas de seguridad, vulnerabilidades, remediaciones | 5 h | 5 | Diseño, construcción, certificación |
| Analista de hacking / desarrollo seguro Sr | Threat modeling, revisión de arquitectura, aprobación previa a producción | 3 a 4 h | 4 | Análisis, diseño, aprobación |

## Plataforma destino → rol de desarrollo

| Plataforma | Complejidad Alta | Baja / Media |
|---|---|---|
| Web .NET, Desktop .NET, API .NET | Developer .NET Sr | Developer .NET Ssr (o Nearshore) |
| Power Apps / Power Automate | Developer Power Apps Sr | Developer Power Apps Ssr |
| SharePoint | Developer SharePoint Sr | Developer SharePoint Ssr |
| Power BI | Developer Power BI Sr | Developer Power BI Ssr |
| iOS / Android | Developer iOS/Android Sr | Developer iOS/Android Ssr |
| Datos (ETL, carga, BD) | Data Engineer Sr | Data Engineer Medium |

**Plataforma no listada** (la plataforma destino viene del pedido o del documento, p. ej. Angular, Java, microfrontends, nube): rol `Developer <plataforma> Sr` para Alta y `Developer <plataforma> Ssr` para Baja/Media, con la misma capacidad que los developers de la tabla. Se asigna a quien tenga esa plataforma en "Plataformas / skills".

Si solo hay un nivel disponible en Recursos, usar ese y registrar la diferencia como riesgo si la complejidad no coincide.

## Matriz de asignación por fase (R = responsable, A = apoyo/revisión)

| Fase | JP | LT | Arq | Plat | Dev | QA Sr | QA Ssr/Tester | UX/UI | Datos | Seg |
|---|---|---|---|---|---|---|---|---|---|---|
| 00 Gestión | R | A | | | | | | | | |
| 01.1 Entendimiento | A | R | A | | | | | | | |
| 01.2 Relevamiento | | R | A | R (ambientes) | R (IA, repos) | | | | | |
| 02 Análisis | | A (Alta: R) | | | R | A | | A | R (datos) | |
| 03 Diseño | | A (Alta: R) | R (TRV-06) | | R | R (TRV-QA) | | R (pantallas) | R (modelo) | R (TRV-SEC) |
| 04 Construcción | | R (revisión) | A | | R | | | A | R | |
| 05 Pruebas unitarias | | A | | | R | | | | R | |
| 06 Certificación | | R (coord.) | | | R (soporte) | R | R | | | R |
| 07 Despliegue | R (plan) | R | | R | A | | A | | A | A |
| 08 Seguimiento | A | R | | A | R | | | | | |

## Equipo dimensionado por el arnés

Solo si Recursos no cubre lo necesario **y el usuario lo autoriza explícitamente**.

1. 1 Jefe de Proyecto (o la gestión la asume el LT si el proyecto tiene < 500 h).
2. 1 Líder Técnico.
3. Desarrolladores: por cada plataforma destino, `ceil(horas de construcción + pruebas de esa plataforma / (cap. planificable × 5 × semanas objetivo))`, mínimo 1. Si no hay fecha objetivo, apuntar a que ninguna ola supere su duración máxima. Componentes Alta > 200 h → un desarrollador dedicado.
4. QA: si "Certificación a cargo de" = QA proveedor o Mixta → 1 Analista Calidad Ssr por cada 3-4 desarrolladores; + 1 Analista Calidad Sr si hay ≥ 3 olas o ≥ 6 desarrolladores.
5. Especialista de Plataformas si n_app ≥ 2 o hay despliegue a producción; si no, sus tareas las toma el LT.
6. Arquitecto de Aplicaciones si hay ≥ 3 plataformas destino, un aplicativo Nuevo o integraciones ≥ 5 en total.
7. UX/UI si un aplicativo destino es Nuevo o hay ≥ 10 pantallas en total.
8. Seguridad si "Datos sensibles o exposición a internet" = Sí.

Nombres de instancia: `<Rol> <n>` (p. ej., "Developer .NET Sr 1"). Marcar todos como **Propuesto** en la hoja Equipo.

## Reglas de asignación

- Un componente tiene un desarrollador principal para todas sus tareas 02-05 (salvo validaciones/diseños de Alta que hace el LT o el Arquitecto, y la revisión de código del LT).
- Nadie revisa su propio código: si el LT desarrolla un componente, su revisión la hace el Arquitecto o un Developer Sr.
- Evitar que el LT sea cuello de botella: repartir transversales de análisis entre desarrolladores con revisión del LT.
- Nearshore: sumar 10 % de coordinación como tarea Asumida del LT.
- Ningún perfil queda en el equipo sin tareas; si ocurre, retirarlo o justificarlo en Equipo ("Sin tareas: …").
