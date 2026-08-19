# Instalación de CapIV en Excel

El usuario sólo necesita Excel de escritorio, conexión a internet y `capiv-installer.zip`.

1. Descargar `capiv-installer.zip` desde el enlace oficial.
2. Hacer clic derecho y elegir **Extraer todo**.
3. Cerrar Excel completamente.
4. Ejecutar `instalar.bat` y esperar el mensaje de instalación completada.
5. Abrir Excel con un libro nuevo.
6. Ir a **Inicio > Complementos > Más complementos > Complementos de desarrollador**.
7. Seleccionar **CapIV**.

Para quitarlo, cerrar Excel y ejecutar `desinstalar.bat` desde la misma carpeta extraída.

## Descargar y actualizar una base

En **Datos**, elegir áreas, año inicial y nivel Área o Pozo. La opción predeterminada **Nueva hoja** crea `CapIV_Datos` desde A1 sin seleccionar ninguna celda. **Celda actual** queda como alternativa para una ubicación personalizada. Para incorporar el último mes disponible, abrir el mismo Excel y pulsar **Actualizar datos**.

El flujo **Pronósticos** es opcional e independiente. Genera un pronóstico separado por cada área o concesión seleccionada; `Resumen_Areas` las reúne para comparar.

Para configurar varias concesiones rápido, usar **Cambio masivo** y elegir el alcance: todas, un activo o una concesión. Los campos vacíos no se modifican. También se pueden crear activos opcionales —por ejemplo Mendoza—: cada concesión sigue pronosticándose por separado y se agrega una hoja `Activo_<nombre>` con la suma y gráficos propios. Sin activos creados, todo queda separado.

La guía completa está en `output/pdf/CapIV_Tutorial_Lanzamiento.pdf`.
