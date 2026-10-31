# Plantilla de salida (`estimacion_<slug>.xlsx`)

Réplica estructural y visual del modelo [Ejemplo_proyecto.xlsx](../assets/Ejemplo_proyecto.xlsx) (solo lectura; abrirlo para consultar estilos, fórmulas y formatos ante cualquier duda). Diez hojas, en este orden:

1. Resumen ejecutivo · 2. Estrategia de oleadas · 3. Matriz de entregables · 4. Planning - <Tipo de trabajo> · 5. Capacidad por recurso · 6. Equipo · 7. Detalle de horas · 8. Supuestos · 9. Riesgos · 10. Validaciones

Nombres definidos obligatorios: `FechaInicio` = 'Resumen ejecutivo'!$B$9 · `Feriados` = Supuestos!$B$<f1>:$B$<f1+29> (30 filas, admite agregar).

Todas las cifras derivadas son **fórmulas vivas** (SUMIF, SUMIFS, COUNTIF, WORKDAY, MAX); solo las horas diarias del Planning y los datos descriptivos son valores.

## Estilos comunes

| Elemento | Formato |
|---|---|
| Título (A1) | Relleno #17365D, fuente blanca negrita 15, Carlito/Calibri, alto 26 |
| Encabezado de tabla | Relleno #1F4E78, fuente blanca negrita 11, ajuste de texto, alto 30-46 |
| Subtítulo de sección | Fuente #17365D negrita 12, sin relleno |
| Celda editable (FechaInicio) | Relleno #FFF2CC, fuente azul #0000FF negrita, formato dd/mm/yyyy |
| Fila resaltada (Matriz de entregables) | Relleno #DDEBF7 |
| Horas | `#,##0` · porcentajes `0%;-0%;"-"` · horas diarias `0;-0;"-"` |
| Estado OK / REVISAR / ALERTA | Verde #C6EFCE/#006100 · Rojo #FFC7CE/#9C0006 · Amarillo #FFEB9C/#9C5700 |
| Bordes | Finos en todas las tablas |
| Pestaña Resumen | Color #17365D · Planning #070F26 |

## 4. Planning - <Tipo de trabajo>  (hoja principal)

Sea `R` = n.º de recursos, `A` = n.º de entradas de la leyenda (aplicativos origen + grupos grises), `bloque` = máx(11, R + 2, A + 1), `HR` = bloque + 7 (fila de encabezados), `D1` = HR + 1 (primera tarea), `N` = última fila de tarea, `LC` = columna del último día (K = día 1).

**Columnas de tarea**: A ID · B Ola · C Fase · **D Aplicativo origen** · E Aplicativo destino · F Plataforma · G Componente · H Tarea · I Recurso asignado · J Horas · K..LC días.
Aplicativo origen = columna "Aplicativo origen" del Inventario; en filas sin origen único usar: `Transversal` (TRV-*), `Base común` (BC-*), `Varios` (bloques de certificación, despliegue y seguimiento).

**Bloque resumen (filas 1..bloque)**

| Celda | Contenido |
|---|---|
| A1 / B1 | "Duración" · C1 "Leyenda de aplicativos origen" · G1 "Cap. h/día" · H1 "Rol" · I1 "Recurso" · J1 "Total horas" · K1..LC1 = 1..n (n.º de día) |
| A2..A6 / B2..B6 | Horas `=J<tot>` · Días `=COUNTIF(K<tot>:LC<tot>,">0")` · Semanas `=ROUND(B3/5,1)` · Inicio `=FechaInicio` · Fin `=WORKDAY(B5,B3-1,Feriados)` |
| A8..A11 | Leyenda de carga: "= cap. = jornada completa" · "< cap. = carga parcial" · "> cap. = sobreasignación" |
| C2..C(A+1) | Un aplicativo origen por fila (orden de aparición en el Planning), con relleno de su color; al final `Transversal`, `Base común` y `Varios` en gris. El texto puede ocupar C:F |
| Filas 2..R+1 | G = capacidad del recurso · H = rol (agrupador) · I = recurso · J `=SUM(K r:LC r)` · K..LC `=SUMIF($I$D1:$I$N,$I r,K$D1:K$N)` |
| Fila R+2 | I "TOTAL >>" · J/K..LC `=SUM(...)` de las filas de recursos |

**Paleta de aplicativos origen** (suave, fuente negra; se asigna en orden y se repite cíclicamente si hay más de 12):

| # | Color | # | Color | # | Color |
|---|---|---|---|---|---|
| 1 | #BDD7EE azul | 5 | #D9C3E9 lila | 9 | #E2D3B5 arena |
| 2 | #C6E0B4 verde | 6 | #B4DEDE turquesa | 10 | #DDD9C3 caqui |
| 3 | #F8CBAD salmón | 7 | #F4B6C2 rosa | 11 | #B7C9E2 pervinca |
| 4 | #FFE699 amarillo | 8 | #D6DCE4 gris azulado | 12 | #E4DFEC lavanda |

Grupos sin origen único: Transversal #E7E6E6 · Base común #D9D9D9 · Varios #EDEDED.

**Cabecera del calendario**

| Fila | Contenido |
|---|---|
| bloque+4 (hitos) | Código del hito (H0, H1.., Pase, Cierre) en la columna de su día; CF: relleno #070F26, fuente #FFC400 |
| bloque+5 (semanas) | "S1", "S2"… cada 5 columnas (K, P, U, …), negrita |
| bloque+6 (fechas) | `=WORKDAY(FechaInicio,K$HR-1,Feriados)`, formato `dd/mm`, tamaño 8, orientación 90° |
| HR (encabezados) | A ID · B Ola · C Fase · D Aplicativo origen · E Aplicativo destino · F Plataforma · G Componente · H Tarea · I Recurso asignado · J Horas · K..LC = 1..n. Relleno #070F26, fuente blanca negrita, Arial 10, alto 42 |

**Filas de tarea (D1..N)**: A `T001`… · B Ola (relleno por ola: Transversal #EDEDED, Ola 1 #DDEBF7, Ola 2 #E2EFDA, Ola 3 #FFF2CC, Ola 4 #FCE4D6, Ola 5+ #EDE7F6, Bloques de despliegue #F2F2F2) · C Fase (etiqueta exacta, sin color) · D relleno del color de su aplicativo origen · E..I texto · J `=SUM(K r:LC r)` · K..LC horas enteras (vacío si 0), Arial 9, formato `0;-0;"-"`.
Filas con origen Asumida o Contingencia: prefijo en H `[Asumida] ` o `[Contingencia] `.

**Formatos condicionales** (sin funciones, independientes del idioma):
- Barras del Gantt en K D1:LC N, una regla por entrada de la leyenda: `=($D D1="<aplicativo origen>")*(K D1>0)` → relleno del color asignado, fuente negra (duplicar comillas si el nombre las contiene).
- Carga en K2:LC(R+1): rojo `=K2>$G2`; verde `=K2=$G2`; amarillo `=(K2>0)*(K2<$G2)`.
- Hitos en K(bloque+4):LC(bloque+4): `=K<f><>""`.

Anchos: A 7 · B 11 · C 24 · D 22 · E 14 · F 11 · G 26 · H 60 · I 22 · J 8 · días 4,5. Inmovilizar paneles en K D1. Autofiltro en A HR:J N.

## 1. Resumen ejecutivo

- A1 título: `<Cliente> — <Nombre del proyecto>: estimación ordenada por oleadas (<versión>)`.
- A3:B: Campo/Valor → Artefacto (ruta de salida), Fuente de horas ("Modelo paramétrico del arnés + entrada <archivo>"), Plantilla de formato, Alcance, Modelo de entrega, **B9 Fecha de inicio (FechaInicio)**, B10 Fecha fin `='Planning - …'!B6`, Jornada, Equipo, Fecha objetivo y holgura (`=NETWORKDAYS(B10,<fecha objetivo>,Feriados)-1`, negativo = brecha).
- D3:E KPIs: componentes por plataforma (`COUNTIF` en Estrategia de oleadas), olas, duración en días (`=Planning!B3`) y semanas, horas totales, horas por origen (`=SUMIF('Detalle de horas'!<Origen>,"Paramétrica",<Horas finales>)`, ídem Entrada, Asumida, Contingencia), rango optimista y pesimista, máx. h de un recurso en un día, días sin carga en ventanas.
- G3:K por recurso: Recurso · Horas `=SUMIF(Planning!$I$D1:$I$N,G4,Planning!$J$D1:$J$N)` · % participación · Día inicio · Día fin (desde Capacidad por recurso) · fila Total.
- Tabla "Horas por componente" (Componente · Aplicativo origen · Ola · Plataforma · Aplicativo destino · Horas base · Horas finales vía SUMIF a Detalle de horas), "Horas por aplicativo origen" (`=SUMIF(Planning!$D$D1:$D$N,<app>,Planning!$J$D1:$J$N)`, celda del nombre con el color de la leyenda) y "Horas por plataforma".
- Tabla "Hitos principales": Hito · Criterio de salida · Día hábil · Fecha `=WORKDAY(FechaInicio,D-1,Feriados)`.
- Viñetas "Supuestos y riesgos relevantes" (5-8 líneas, detalle en sus hojas).

## 2. Estrategia de oleadas

Encabezados (fila 3): Orden · Ola · Código · Componente · Aplicativo origen · Plataforma · Aplicativo destino · Estrategia · Complejidad (puntaje) · LOC · Pantallas · Integraciones · Horas base · Horas finales · Desarrollador(es) · Listo para certificar (día) · Despliegue (fecha) · Motivo de la secuencia · Riesgos / dependencias · Información faltante o por validar.
Horas por SUMIF a Detalle de horas. Debajo: fila de base común por aplicativo; tabla "Actividades reutilizables o agrupables" (Actividad · Componentes/aplicativos · Cómo se aprovecha en el plan); lista "Criterios de secuenciación aplicados".

## 3. Matriz de entregables

Encabezados: Ola · Entregable · Componentes incluidos · Aplicativos destino · Plataforma · Objetivo del entregable · Inicio construcción (día) · Listo para certificar (día) · Certificación (días por aplicativo) · Fin de certificación (fecha) · Fecha de despliegue · Fin de seguimiento · Horas totales de la ola `=SUMIF(Planning!$B$D1:$B$N,A4,Planning!$J$D1:$J$N)` · Criterio de salida · Riesgos o dependencias. Filas con relleno #DDEBF7. Nota final explicando qué horas cuentan en cada ola.

## 5. Capacidad por recurso

Cuatro secciones, una fila por recurso, columnas S1..Sk + Total:
1. Horas asignadas por semana: `=SUM('Planning - …'!K r:O r)` (bloques de 5 columnas).
2. Capacidad disponible por semana: `cap × días hábiles con disponibilidad` (descontar vacaciones).
3. Utilización: `=IFERROR(h/cap,0)`; columna Ventana `=IFERROR(total/(cap*(fin-inicio+1)),0)`. CF: > 1 rojo, ≥ 0,8 verde, > 0 amarillo.
4. Alertas: Máx. h/día `=MAX(fila)` · Días > cap `=COUNTIF(fila,">"&cap)` · Semanas > 100 % · Día inicio `=IFERROR(MATCH(TRUE,INDEX(fila>0,0),0),"")` · Día fin `=IFERROR(LOOKUP(2,1/(fila>0),$K$1:$LC$1),"")` · Días con carga · Días sin carga en ventana · Estado `=IF(OR(B>cap,C>0,D>0),"ALERTA: sobreasignación",IF(H>0,"Revisar vacíos","OK"))`.
Notas al pie: ventanas escalonadas, justificación de días sin carga.

## 6. Equipo

Encabezados: Recurso · Rol · Estado (Propuesto/Confirmado) · Cap. h/día · Rol en el proyecto · Fases aplicables · Componentes / aplicativos asignados · Horas asignadas (SUMIF Planning) · Horas asumidas · Horas de contingencia (SUMIFS Detalle) · Día inicio · Día fin · Utilización en ventana · [Tarifa · Costo, si hay tarifas] · Racional de la asignación. Fila Total.

## 7. Detalle de horas (trazabilidad)

Una fila por tarea del catálogo (antes de agrupar). Encabezados: ID origen (`<ID comp>-<código>`) · Código componente · Componente · Aplicativo origen · Aplicativo destino · Plataforma · Complejidad · Código tarea · Tarea · Driver / fórmula (texto, p. ej. "1438 LOC / 45 × 1,0") · Horas base · Factor IA · Horas finales `=ROUNDUP(ROUND(base*factor,1),0)` · Fase · Recurso · Ola · ID planning · Origen (Paramétrica/Entrada/Asumida/Contingencia) · Observación.
Fila Total. Debajo: "Cambios y decisiones respecto a la entrada" (viñetas).

## 8. Supuestos

Encabezados: ID (ASM-001…) · Categoría · Supuesto · Valor / regla · Punto de validación · Impacto si no se cumple · Estado (Assumption / Constraint / Decision). Incluir siempre: base de horas (modelo y parámetros usados), capacidad, calendario, preparación asumida, certificación, despliegue, seguimiento, asignación, secuencia, factores IA y exclusiones declaradas (Fuera de alcance). No se registran datos del proyecto supuestos: no existen.
Debajo: tabla de feriados (# · Fecha · Descripción, 30 filas) = rango `Feriados`, con nota de edición.

## 9. Riesgos

Encabezados: ID (RSK-001…) · Categoría · Riesgo / dependencia · Evidencia / disparador · Probabilidad · Impacto · Exposición · Estrategia (Mitigar/Aceptar/Transferir/Evitar) · Mitigación / decisión requerida · Responsable · Hito · Estado.
Exposición: Alta×Alta = Crítica; Alta×Media o Media×Alta = Alta; Media×Media, Alta×Baja o Baja×Alta = Media; resto Baja. CF por igualdad exacta (`=G4="Crítica"` rojo #FF0000 texto blanco, Alta rojo claro, Media amarillo, Baja verde).
Fuentes: hoja Riesgos de la entrada + riesgos detectados (información faltante, dependencias críticas, concentración en un recurso, factores IA sin calibrar, ausencia de QA, ambientes).

## 10. Validaciones

Encabezados: # · Validación · Resultado (fórmula) · Esperado · Estado `=IF(cond,"OK","REVISAR")` · Detalle / justificación. Ver [validaciones](./validaciones.md).
Debajo: matriz de cobertura Componente × Fase (01.1, 01.2, 02..08) con `COUNTIFS(Planning!G,"*<ID>*",Planning!C,"<fase>*")`; CF: 0 rojo, > 0 verde. Las fases que no aplican por alcance se marcan "N/A".
