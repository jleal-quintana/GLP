# GLP

Office Web Add-in para Excel, basado tecnicamente en `addin-decision-tree`, para descargar datos de Capitulo IV, armar historicos por area, proyectar y generar graficos.

Version publica: `https://jleal-quintana.github.io/GLP/`

Instalador para usuarios: [`tutorial/glp-installer.zip`](https://github.com/jleal-quintana/GLP/raw/main/tutorial/glp-installer.zip)

## Desarrollo

```bash
npm install
npm run dev
npm run start:desktop
npm run test
npm run validate
npm run build
```

Servidor local del taskpane: `https://localhost:3002/taskpane.html`.

## Funcionalidad

- Catalogo desde `Capitulo IV - Pozos`, con provincia y area.
- Filtro por provincia, busqueda por area/empresa y seleccion masiva de areas filtradas.
- Seleccion de multiples areas con ano de inicio global y editable por area.
- Descarga convencional + no convencional.
- Hojas por area: historico, pronostico de produccion, pronostico de pozos, graficos y detalle pozo-mes.
- Resumen consolidado con historico y suma de proyecciones individuales.
- Hoja visible `CapIV_Debug` con log paso a paso.
- Warning interactivo si aparecen meses faltantes en el medio de la serie; los faltantes iniciales se completan con 0.
- Modos: actualizar datos y regenerar area.

Los archivos `.xlsm` en la raiz son referencias funcionales, no dependencias del add-in.
