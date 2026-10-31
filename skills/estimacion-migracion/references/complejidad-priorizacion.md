# Complejidad y priorización

## Puntaje de complejidad (por componente)

Cada criterio vale 1 (Bajo), 2 (Medio) o 3 (Alto).

| Criterio | 1 – Bajo | 2 – Medio | 3 – Alto | Fuente en Inventario |
|---|---|---|---|---|
| Reglas de negocio | Simples | Varias reglas | Complejas o ambiguas | Reglas de negocio |
| Integraciones | Ninguna | Una | Múltiples (≥ 2) | Integraciones (n) |
| Datos | Consulta simple / sin migración | Lectura y escritura / migración Simple | Transformación o migración Compleja | Migración de datos + Tablas |
| Pantallas | Sin pantalla | 1 a 3 pantallas | ≥ 4 pantallas o flujos | Pantallas / formularios |
| Riesgo técnico | Conocido | Moderado | Tecnología o dependencia nueva | Riesgo técnico |
| Dependencia externa | Ninguna | Parcial | Crítica | Dependencia externa |
| Reutilización | Alta | Media | Baja | Reutilización |
| Tamaño (LOC) | < 1.000 | 1.000 – 3.000 | > 3.000 | LOC |

**Puntaje** = promedio de los 8 criterios.
- Baja: ≤ 1,6 · Media: > 1,6 y ≤ 2,3 · Alta: > 2,3.

**Reglas de escalamiento** (aplican después del promedio):
- Riesgo técnico Alto **y** Dependencia Crítica → Alta.
- LOC > 5.000 o Integraciones ≥ 4 → mínimo Media; si además Reglas = Alto → Alta.
- "Complejidad forzada" informada → prevalece; registrarla como decisión de la entrada.

**Filas agregadas** ("Módulos (n)" > 1): LOC, pantallas e integraciones se miden por módulo promedio (total ÷ n) para el puntaje y las reglas de escalamiento; así un aplicativo grande no sale Alta solo por su tamaño total. La fila se parte en lotes de módulos (`<ID>-L1`, `<ID>-L2`…) con las cantidades repartidas en partes iguales, de modo que las horas siguen saliendo de los totales. Cantidad de lotes: `lotes` en `asignacion.json`; si no, uno por cada desarrollador de `desarrolladores`; si no, ⌈n ÷ "Módulos por lote"⌉ (defecto 10). Migración de datos y permisos se estiman una sola vez (lote 1).

Registrar en Estrategia de oleadas: puntaje, complejidad resultante y criterio que más pesó.

## Prioridad de ejecución

Orden de prioridad (de mayor a menor) para asignar oleada y posición en la serie de cada desarrollador:

1. Dependencias obligatorias (Restricciones "Orden obligatorio", base común del aplicativo destino).
2. **Piloto por plataforma**: en la Ola 1, un componente representativo (Baja o Media) por plataforma destino para calibrar factores IA y patrones técnicos.
3. **Componentes grandes (Alta o > 200 h)**: inician lo antes posible con un desarrollador dedicado (son la ruta crítica).
4. **Riesgo técnico alto**: temprano, para validar el patrón (integraciones nuevas, dependencias críticas).
5. **Quick wins** (Baja, alto valor): Ola 1 para generar entregable temprano.
6. Prioridad de negocio (1 → 3).
7. Reutilización: el primero que implementa un patrón compartido (servicio de notificaciones, integración, COM, etc.) va antes que quienes lo reutilizan; registrar el patrón en "Actividades reutilizables o agrupables".

## Armado de oleadas

- **Cuántas olas**: las que pida el usuario. Si el pedido dice "en n olas" (o la entrada trae Modelo de entrega = Oleadas con "Número de olas" = n), se arman exactamente n olas y se guarda `numeroOlas: n` en `asignacion.json`. Si no se menciona, es **una sola entrega** (todo en "Ola 1"). El arnés no inventa ni parte olas por su cuenta.
- Con n olas:
  - **Quién agrupa**: si el pedido (o la columna "Ola sugerida") dice qué componentes van en cada ola, esa agrupación **se respeta tal cual**. Si genera un conflicto (dependencia que va al revés, recurso sobrecargado), no se cambia: se informa y se pregunta. Si el pedido dice que la agrupación la define el arnés, se reparten los componentes con los criterios de priorización de abajo. Si no dice nada, se pregunta.
  - Ninguna ola queda vacía.
  - Cada ola produce un entregable certificable (criterio de salida explícito en la Matriz de entregables).
  - Agrupar por aplicativo destino y patrones compartidos; un desarrollador hace en **serie** sus componentes de una ola (termina uno antes de iniciar el siguiente) y puede adelantar análisis/diseño de la ola siguiente mientras espera una revisión o dependencia.
  - Olas pueden solaparse si usan recursos distintos (frentes paralelos); la ola N+1 de un recurso comienza cuando ese recurso termina su ola N.
  - En "Motivo de la secuencia" se indica si la ola la definió el usuario o el arnés (y con qué criterio).
- Si "Agrupación de certificación y despliegue" = Pase único → certificaciones por ola/aplicativo y un solo despliegue al final; = Por ola → un bloque de certificación + despliegue + seguimiento por ola; = Por aplicativo destino → un bloque por aplicativo dentro de cada ola.

## Salidas de este paso

- Complejidad por componente (con puntaje y criterio dominante).
- Ola, posición en la serie y desarrollador asignado por componente.
- Lista de actividades reutilizables o agrupables (actividad, componentes, cómo se aprovecha).
- Criterios de secuenciación aplicados (texto numerado para la hoja Estrategia de oleadas).
