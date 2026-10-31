---
description: "Completa la plantilla de entrada de estimación a partir de documentos libres (relevamientos, inventarios de macros, planes previos, correos, actas)."
agent: "Estimador de Migraciones"
argument-hint: "Nombre del proyecto (y opcionalmente rutas de documentos fuente)"
---
Prepara la entrada de estimación del proyecto **${input:proyecto:Nombre del proyecto}** a partir de: ${input:fuentes:Rutas de documentos fuente; si se omite, se usa Estimaciones/<proyecto>/fuentes/}

`<slug>` = nombre del proyecto tal como lo escribió el usuario, sin espacios ni caracteres especiales (p. ej. `MigraMacro`). Si no se indicaron fuentes, usar todos los archivos de `Estimaciones/<slug>/fuentes/`; si esa carpeta no existe o está vacía, crearla y pedir al usuario que coloque ahí la información.

1. Crea la carpeta `Estimaciones/<slug>/` si no existe y copia `.github/skills/estimacion-migracion/assets/Plantilla_Entrada_Estimacion.xlsx` a `Estimaciones/<slug>/entrada_<slug>.xlsx` (no modifiques la plantilla del arnés). Si `entrada_<slug>.xlsx` ya existe, no la reemplaces: respáldala en `Estimaciones/<slug>/backup/entrada_<slug>_<yyyyMMdd_HHmm>.xlsx` y complétala sobre la existente.
2. Lee las fuentes en solo lectura y completa las hojas Proyecto, Inventario, Recursos, Calendario, Restricciones y Riesgos según `references/entradas.md`.
3. No inventes valores: deja vacío lo que no esté en las fuentes y regístralo en "Información faltante o por validar" del componente o en una lista de preguntas.
4. Respeta las listas de la hoja Catalogos.
5. Devuelve: ruta del archivo, resumen de lo cargado (n.º de componentes por plataforma, recursos, fechas) y la lista de datos faltantes priorizada (bloqueantes primero).
