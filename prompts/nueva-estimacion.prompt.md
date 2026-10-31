---
description: "Genera una estimación completa (WBS, horas, equipo, oleadas, Gantt diario y Excel con la plantilla estándar) a partir de la plantilla de entrada completada."
agent: "Estimador de Migraciones"
argument-hint: "Ruta de la entrada completada, p. ej. Estimaciones/mi-proyecto/entrada_mi-proyecto.xlsx"
---
Genera la estimación del proyecto a partir de la entrada: ${input:entrada:Ruta de Plantilla_Entrada_Estimacion.xlsx completada}

Sigue el procedimiento completo del skill `estimacion-migracion` (pasos 1 a 13), sin saltarte la pre-validación ni la validación final.
Si el pedido indica un número de olas, úsalo; si no, es una sola entrega.
Si faltan datos bloqueantes, pregúntalos todos juntos antes de calcular.
Al terminar, entrega el resumen: KPIs, rango, equipo, oleadas e hitos, validaciones, supuestos clave, información faltante y preguntas para el cliente.
