# CapIV para Excel

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
4. En **Cambio masivo**, aplicar métodos y parámetros Di/b a todas las concesiones, a un activo o a una sola concesión.
5. Opcionalmente, agrupar concesiones en activos. Cada concesión conserva su pronóstico individual y el activo agrega una hoja de resumen con gráficos propios.
6. Elegir **Actualizar** para conservar supuestos editados en Excel o **Regenerar** para reconstruir las hojas.
7. Revisar los resúmenes, los gráficos y `CapIV_Debug`.

### Actualizar un libro existente

El flujo **Datos** crea una única tabla plana. Por defecto, CapIV genera una hoja nueva (`CapIV_Datos`, `CapIV_Datos_2`, etc.) y comienza en A1; opcionalmente, el usuario puede elegir **Celda actual** para insertarla en una ubicación personalizada. Puede salir agregada por área y mes o detallada por pozo-mes. Antes de escribir sobre contenido existente, CapIV muestra el rango afectado y exige confirmación.

Al abrir un archivo generado anteriormente, usar **Actualizar datos**. CapIV recupera áreas, nivel y hoja de destino desde `_CapIV_State`, vuelve a consultar la serie oficial completa e incorpora el último mes o correcciones sin duplicar registros. El flujo **Pronósticos** es opcional e independiente.

Los pronósticos se calculan **por área/concesión**: si se seleccionan varias, CapIV genera un juego independiente de hojas para cada `cod_area`. `Resumen_Areas` las consolida visualmente, pero los cálculos no mezclan concesiones. La hoja de pozos proyecta actividad/cantidad de pozos del área; no genera una curva de producción individual por pozo.

Los **activos** son opcionales. Se pueden crear, por ejemplo, activos individuales para CLME y EFO, y un activo Mendoza con sus seis concesiones. Las curvas siguen calculándose por concesión; CapIV suma las fórmulas de sus hojas `_Prono` en una hoja `Activo_<nombre>` con gráficos de petróleo, gas, agua y producción bruta. Si no se crea ningún activo, no cambia el comportamiento anterior y todo queda separado.

## Salida en el libro

- `{AREA}_HDP`: histórico mensual.
- `{AREA}_Prono`: pronóstico de producción y supuestos editables.
- `{AREA}_Pozos`: pronóstico de actividad.
- `{AREA}_Graficos`: 11 gráficos técnicos separados: petróleo, gas, bruta, agua, corte de agua, RGP, RAP/WOR, inyección, acumuladas líquidas, pozos y RAP vs. Np.
- `{AREA}_Detalle`: detalle pozo-mes.
- `Resumen_Areas`: consolidado que sigue los cambios hechos en cada pronóstico.
- `Activo_<nombre>`: resumen y cuatro gráficos propios de las concesiones agrupadas; sólo se crea si el usuario define el activo.
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

También se admite una base sin plantilla, a la que CapIV agrega `?url=`. El proxy debe limitarse por allowlist a `datos.gob.ar` y `datos.energia.gob.ar`, conservar streaming y no registrar datos sensibles. En GitHub Pages, crear la variable de repositorio `CAPIV_PROXY_BASE_URL`; el workflow ya la incorpora al build.

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
