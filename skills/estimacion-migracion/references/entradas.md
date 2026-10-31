# Entradas de la estimación

**Principio: ningún dato del proyecto se supone.** Todo lo que describe al proyecto (aplicativos, tecnologías, LOC, pantallas, integraciones, tablas, recursos, calendario) debe venir en la entrada. Si falta, el agente **se detiene**, lista todos los faltantes en un solo bloque y espera que el usuario los complete. No usa valores por defecto ni los registra como supuesto.

Fuente: `Estimaciones/<slug>/entrada_<slug>.xlsx` (copia de [Plantilla_Entrada_Estimacion.xlsx](../assets/Plantilla_Entrada_Estimacion.xlsx)). Encabezados en la fila 3, datos desde la fila 4 (`Proyecto` y `Parametros` son pares Campo/Valor desde la fila 4).

Niveles:
- **B – Bloqueante**: dato del proyecto. Vacío = detener y pedir. El valor 0 es válido si se escribe explícitamente (p. ej. 0 integraciones).
- **P – Política**: decisión de entrega que la plantilla trae **precargada** con el valor recomendado y visible. El agente usa lo que esté escrito; si la celda se vacía, es Bloqueante.
- **Op – Opcional**: enriquece la salida; vacío no afecta el cálculo.
- **D – Derivado**: el agente lo calcula a partir de otros datos de la entrada (no es una suposición).

## Qué NO se configura en el arnés

El arnés (agente, skill, referencias, plantilla) es genérico y no se modifica por proyecto. La cantidad de aplicativos, sus tecnologías origen y destino, plataformas, componentes y recursos se leen del Inventario y de Recursos de cada entrada; el agente deriva `n_app`, plataformas, base común por aplicativo y equipo requerido. Solo se toca el arnés para cambiar reglas corporativas (catálogos, parámetros calibrados, plantilla de salida).

## Hoja `Proyecto`

| Campo | Nivel | Regla / uso |
|---|---|---|
| Nombre del proyecto | B | Títulos, slug del archivo |
| Cliente / área solicitante | Op | Resumen |
| Tipo de trabajo | B | Migración / Upgrade / Mejora evolutiva |
| Objetivo | B | Resumen |
| Alcance (resumen) | B | Resumen |
| Fuera de alcance | Op | Hoja Supuestos (como exclusión declarada) |
| La estimación llega hasta | P | Certificación / Producción / Hypercare; recorta fases 07-08 |
| Tecnología origen predominante | D | Resumen de "Tecnología origen" del Inventario |
| Tecnología destino predominante | D | Resumen de "Plataforma destino" del Inventario |
| Ambientes disponibles | B | Riesgo si no hay Certificación/QA |
| Fecha tentativa de inicio | B | `FechaInicio` |
| Fecha objetivo | Op | Holgura o brecha contra el fin estimado |
| Jornada (h/día) | P | Horas por recurso y día (defecto 8, configurable). Carga diaria planificada y tope absoluto |
| Días laborables | P | `WORKDAY` |
| Uso de asistente IA en el desarrollo | P | Activa factores IA |
| Modelo de entrega | P | Entrega única (defecto) / Oleadas. Si el pedido indica n olas, manda el pedido |
| Número de olas | Op | Solo con Oleadas (≥ 2). Vacío o 1 = una sola entrega |
| Agrupación de certificación y despliegue | P | Por ola / Por aplicativo destino / Pase único |
| Certificación a cargo de | P | QA proveedor / QA cliente / Mixta |
| Duración mínima de certificación (días hábiles) | P | Regla de certificación |
| Duración de hypercare (días hábiles) | P | Fase 08 |
| Claridad de requerimientos | B | Alta / Media / Baja; contingencia y rango |
| Datos sensibles o exposición a internet | B | Incluye actividades y roles de seguridad |
| Usuarios involucrados | Op | Resumen / UX |
| Criterios de aceptación generales | Op | Resumen |
| Restricciones generales | Op | Riesgos |
| Tarifas por rol disponibles | Op | Activa costeo en Equipo |
| Moneda | Op (B si hay tarifas) | Costeo |
| Versión de la estimación | P | Título y control de cambios |

## Hoja `Inventario` (una fila por componente: módulo, programa, macro, requerimiento o aplicativo origen, de cualquier tecnología legacy)

Si se tiene el detalle por módulo, va una fila por módulo. Si solo se tienen totales de un aplicativo (p. ej. "40 módulos, 30.000 LOC, 120 pantallas"), va **una sola fila** con esos totales y la cantidad en "Módulos (n)": el arnés mide la complejidad por módulo promedio y parte la construcción en lotes para trabajar en paralelo.

| Columna | Nivel | Nota |
|---|---|---|
| ID | B | Único |
| Nombre | B | |
| Descripción / objetivo | Op | |
| Aplicativo origen | B | |
| Tecnología origen | B | |
| Aplicativo destino | B | |
| Destino existente o nuevo | B | |
| Plataforma destino | B | Texto libre: la indica el pedido o el documento de arquitectura/diseño. La lista de Catalogos es solo sugerencia |
| Estrategia | B | Migrar / Refactorizar / Reconstruir / Reemplazar por estándar / Retirar / Mantener (lista de Catalogos) |
| Tipo de trabajo | Op | Vacío = hereda el tipo de trabajo del proyecto (regla, no suposición) |
| Prioridad de negocio (1-3) | Op | Vacío = sin prioridad diferencial (no altera la secuencia) |
| Ola sugerida | Op | Vacío = la define el arnés |
| LOC | B* | Líneas lógicas del origen; en evolutivas, LOC a crear/modificar estimadas por el equipo |
| Pantallas / formularios | B* | 0 explícito si no tiene |
| Reportes / salidas | B* | 0 explícito si no tiene |
| Integraciones (n) | B* | 0 explícito si no tiene |
| Detalle de integraciones | B si Integraciones > 0 | Sistema y tipo de cada integración |
| Tablas / hojas / fuentes de datos (n) | B* | 0 explícito si no tiene |
| Migración de datos | B* | Ninguna / Simple / Compleja |
| Reglas de negocio | B* | Bajo / Medio / Alto |
| Riesgo técnico | B* | Bajo / Medio / Alto |
| Dependencia externa | B* | Ninguna / Parcial / Crítica |
| Reutilización | B* | Alta / Media / Baja |
| Requiere permisos y auditoría | B* | Sí / No |
| Usuarios / roles | Op | |
| Criterios de aceptación | Op | |
| Información faltante o por validar | Op | Lo que el usuario declare aquí se trata como contingencia, riesgo y pregunta al cliente |
| Complejidad forzada | Op | Prevalece sobre el cálculo |
| Horas override | Op | Reemplaza Σ(02-05) del componente; origen `Entrada` |
| Observaciones | Op | |
| Módulos (n) | Op | Solo para filas agregadas (un aplicativo con totales). Vacío o 1 = la fila es un solo componente |

\* No aplica a componentes con estrategia **Retirar** o **Mantener**.

## Hoja `Recursos`

| Columna | Nivel | Nota |
|---|---|---|
| Recurso (nombre o alias) | B | |
| Rol | B | Lista de Catalogos |
| Capacidad h/día | Op | Vacío = capacidad planificable del rol (regla del catálogo) |
| Dedicación % | B | |
| Disponible desde / hasta | Op | Vacío = todo el proyecto |
| Plataformas / skills | B para roles de desarrollo | Para asignar componentes por plataforma |
| Tarifa por hora | B si "Tarifas por rol disponibles" = Sí | |

Capacidad efectiva diaria = min(Capacidad h/día × Dedicación %, Jornada), redondeada hacia abajo (mínimo 1).

**Hoja vacía o insuficiente** (falta cubrir una plataforma, QA cuando "Certificación a cargo de" ≠ QA cliente, o seguridad cuando hay datos sensibles): el agente **pregunta** si debe (a) esperar que el usuario complete Recursos o (b) dimensionar el equipo según [roles-capacidad](./roles-capacidad.md#equipo-dimensionado-por-el-arnés). Solo con autorización explícita crea recursos marcados **Propuesto**.

## Hoja `Calendario`

Fecha, Tipo (Feriado / No laborable / Vacaciones / Ventana de despliegue / Congelamiento del cliente), Recurso (vacío = todos), Descripción.
- Hoja vacía → el agente pide confirmar "sin feriados ni ausencias en el periodo" antes de programar.
- Feriado / No laborable sin Recurso → rango `Feriados`. Vacaciones → capacidad 0 para ese recurso ese día.
- Ventana de despliegue → los pases solo caen en esas fechas. Congelamiento del cliente → sin pases ni certificación esas fechas.

## Hojas `Restricciones` y `Riesgos` (opcionales)

- Restricciones: Dependencia externa (fecha mínima de construcción del componente), Hito fijo, Ventana de despliegue, Recurso compartido (reduce capacidad), Orden obligatorio (A antes que B).
- Riesgos: se copian a la salida y se complementan con los detectados por el arnés.

## Hoja `Parametros` (política, precargada)

Pares Parámetro/Valor del [modelo de horas](./modelo-horas.md) y de [programación](./reglas-programacion.md). Se leen por nombre exacto. Si un parámetro está vacío → Bloqueante.

## Controles de la entrada (antes de estimar)

1. Todos los campos B y P completos; listas válidas según Catalogos.
2. IDs de Inventario únicos.
3. Fecha objetivo ≥ fecha de inicio; disponibilidad desde ≤ hasta.
4. Cada plataforma destino tiene al menos un recurso de desarrollo que la cubra.
5. Rol de calidad presente si "Certificación a cargo de" = QA proveedor o Mixta; rol de seguridad si hay datos sensibles.

Si algún control falla, responder con una tabla **Hoja · Fila/ID · Campo · Problema** y no continuar hasta que el usuario corrija la entrada.
