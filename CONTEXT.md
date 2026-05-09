# GLP Context

## Domain Terms

### Capitulo IV

Fuente publica de datos de produccion y pozos usada por GLP para armar historicos, proyecciones y graficos en Excel.

### Area

Unidad de seleccion del usuario. Debe tratarse por codigo exacto salvo que exista un alias explicito y auditado.

### Capitulo IV Resource

Archivo publicado por Capitulo IV, normalmente un CSV por concepto o por anio. Un recurso puede ser grande y debe procesarse sin cargarlo completo en memoria.

### Production Download

Proceso de leer recursos de produccion convencional y no convencional, filtrar por Area, normalizar filas y emitir progreso visible.

### Download Ledger

Hoja visible del workbook donde GLP deja evidencia incremental de lo descargado por recurso para que el usuario vea avance y pueda auditar coincidencias.

### Forecast Projection

Proceso de extender la serie historica de un Area usando supuestos editables. Debe producir el mismo criterio para el resumen consolidado y para las formulas visibles que quedan en Excel.
