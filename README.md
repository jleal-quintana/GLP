# GLP - Capítulo IV

Add-in de Excel para consultar la información oficial de Capítulo IV, construir históricos por área, proyectar producción y pozos, y consolidar resultados.

Página de instalación publicada: `https://jleal-quintana.github.io/GLP/instalacion.html`.

## Lanzamiento local recomendado

Requisitos: Windows 10/11, Excel de escritorio de Microsoft 365, Node.js 20 o superior y conexión a internet.

```powershell
npm ci
npm run check
npm run start:desktop
```

El último comando levanta el panel en `https://localhost:3002/taskpane.html` y abre Excel con el add-in cargado. Para detenerlo:

```powershell
npm run stop:desktop
```

Si Windows todavía no confía en el certificado local, ejecutar una vez:

```powershell
npx office-addin-dev-certs install
```

## Flujo de uso

1. Esperar a que cargue el catálogo oficial.
2. Filtrar y seleccionar una o más áreas.
3. Definir año inicial, horizonte y métodos de proyección.
4. Opcionalmente, abrir **Ajustes por área** para sobrescribir parámetros concretos.
5. Elegir **Actualizar datos** para conservar supuestos editados en Excel o **Regenerar áreas** para reconstruir las hojas.
6. Revisar el resumen, los gráficos y `CapIV_Debug`.

### Actualizar un libro existente

Al abrir un archivo generado anteriormente, usar **Actualizar libro** al comienzo del panel. GLP recupera las áreas y la configuración desde `_CapIV_State`, vuelve a consultar la serie oficial completa e incorpora meses nuevos o correcciones sin duplicar registros. Los supuestos editados en las hojas de pronóstico se conservan.

## Salida en el libro

- `{AREA}_HDP`: histórico mensual.
- `{AREA}_Prono`: pronóstico de producción y supuestos editables.
- `{AREA}_Pozos`: pronóstico de actividad.
- `{AREA}_Graficos`: gráficos del área.
- `{AREA}_Detalle`: detalle pozo-mes.
- `Resumen_Areas`: consolidado que sigue los cambios hechos en cada pronóstico.
- `CapIV_Descarga`: trazabilidad de recursos y filas descargadas.
- `CapIV_Debug`: log visible de ejecución.
- `_CapIV_State`: estado interno oculto.

## Calidad

```powershell
npm run check
```

El comando verifica TypeScript, pruebas automáticas, ambos manifests y el build de producción.

## Publicación web

Los metadatos se consultan directamente en Datos Argentina. Los CSV de producción oficiales redirigen actualmente de HTTPS a HTTP; una página HTTPS no puede consumirlos de forma segura. Para una publicación pública se debe configurar `CAPIV_PROXY_BASE_URL` con un proxy HTTPS que acepte una URL destino codificada, por ejemplo:

```text
https://proxy.ejemplo/?url={url}
```

También se admite una base sin plantilla, a la que GLP agrega `?url=`. El proxy debe limitarse por allowlist a `datos.gob.ar` y `datos.energia.gob.ar`, conservar streaming y no registrar datos sensibles. En GitHub Pages, crear la variable de repositorio `CAPIV_PROXY_BASE_URL`; el workflow ya la incorpora al build.

Sin ese proxy, el modo local sigue funcionando mediante el proxy del servidor de desarrollo incluido.

## Desarrollo

```powershell
npm run dev:server
npm run test
npm run validate
npm run validate:prod
npm run build
```

Los archivos `.xlsm` de la raíz son referencias funcionales y no son dependencias del add-in.
