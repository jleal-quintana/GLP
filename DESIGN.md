# Design System - Capitulo IV Add-in

## Marca

Usar identidad Quintana Energy:

- Verde oliva: `#6B7B38`
- Verde bosque: `#33492D`
- Azul marino: `#1B4B6C`
- Gris azulado: `#DAE0E5`
- Verde lima: `#E2FF87`
- Crema: `#FFEAC6`
- Texto: `#1A1A1A`

Tipografia del taskpane: Montserrat para titulos y labels, Inter para cuerpo y datos. En Excel, Calibri como fallback seguro.

## UI

El taskpane es una herramienta de trabajo, no una landing. Una columna, controles compactos, labels tecnicos y estados claros.

Flujo principal:

1. Actualizar catalogo.
2. Filtrar por provincia.
3. Seleccionar una o varias areas.
4. Definir ano de inicio global y overrides.
5. Definir metodos globales y overrides por area.
6. Generar o actualizar hojas.

## Hojas

Por area:

- `{AREA}_HDP`
- `{AREA}_Prono`
- `{AREA}_Pozos`
- `{AREA}_Graficos`
- `{AREA}_Detalle`

Globales:

- `Resumen_Areas`
- `CapIV_Debug`
- `_CapIV_State` very hidden

## Debug

`CapIV_Debug` queda visible. Cada operacion relevante escribe timestamp, paso, estado y detalle/error. El objetivo es poder diagnosticar fallas de Office.js, descarga o escritura sin abrir DevTools.
