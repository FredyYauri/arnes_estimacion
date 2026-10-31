---
description: "Revisa un Excel de estimación existente contra las reglas y validaciones obligatorias del arnés y reporta hallazgos (sin modificarlo)."
agent: "Estimador de Migraciones"
argument-hint: "Ruta del Excel de estimación a revisar"
---
Revisa la estimación ${input:archivo:Ruta del Excel de estimación} en **solo lectura**.

1. Verifica estructura contra `references/plantilla-salida.md` (hojas, columnas, nombres `FechaInicio` y `Feriados`, fórmulas vivas).
2. Ejecuta los controles de `references/validaciones.md` y de `references/reglas-programacion.md` (capacidad diaria, continuidad, dependencias, orden de fases, certificación mínima, horas agregadas marcadas, totales).
3. Contrasta horas por componente con el modelo de `references/modelo-horas.md` y marca desviaciones > 20 %.
4. Entrega una tabla de hallazgos: # · Hoja/celda o tarea · Hallazgo · Severidad (Alta/Media/Baja) · Recomendación. No modifiques el archivo; si el usuario pide corregir, respalda primero en `backup/`.
