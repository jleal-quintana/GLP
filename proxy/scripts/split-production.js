const fs = require('node:fs');
const path = require('node:path');
const { parse } = require('csv-parse');

const [, , csvPath, outputDirectory] = process.argv;
if (!csvPath || !outputDirectory) {
  throw new Error('Uso: node split-production.js <archivo.csv> <directorio-salida>');
}

fs.mkdirSync(outputDirectory, { recursive: true });
const handles = new Map();
let scannedRows = 0;

function textValue(...values) {
  for (const value of values) {
    if (value !== undefined && value !== null && String(value).trim()) return String(value).trim();
  }
  return '';
}

function safeAreaFile(areaId) {
  return `${encodeURIComponent(areaId)}.ndjson`;
}

function appendRecord(areaId, record) {
  let handle = handles.get(areaId);
  if (!handle) {
    handle = fs.openSync(path.join(outputDirectory, safeAreaFile(areaId)), 'a');
    handles.set(areaId, handle);
  }
  fs.writeSync(handle, `${JSON.stringify(record)}\n`);
}

async function run() {
  const parser = fs.createReadStream(csvPath).pipe(parse({
    bom: true,
    columns: (headers) => headers.map((header) => String(header).trim().toLowerCase()),
    skip_empty_lines: true,
    relax_column_count: true,
  }));

  for await (const row of parser) {
    scannedRows++;
    const areaId = textValue(row.idareapermisoconcesion, row.cod_area);
    if (!areaId) continue;
    appendRecord(areaId, {
      idareapermisoconcesion: areaId,
      idpozo: textValue(row.idpozo),
      sigla: textValue(row.sigla),
      anio: textValue(row.anio),
      mes: textValue(row.mes),
      prod_pet: textValue(row.prod_pet, row.prod_petroleo, row.petroleo),
      prod_gas: textValue(row.prod_gas, row.gas),
      prod_agua: textValue(row.prod_agua, row.agua),
      iny_agua: textValue(row.iny_agua, row.agua_iny, row.inyeccion_agua),
    });
  }

  for (const handle of handles.values()) fs.closeSync(handle);
  fs.writeFileSync(
    path.join(outputDirectory, '_complete.json'),
    JSON.stringify({ scannedRows, areas: handles.size, createdAt: new Date().toISOString() }),
  );
  process.stdout.write(`scannedRows=${scannedRows} areas=${handles.size}\n`);
}

run().catch((error) => {
  for (const handle of handles.values()) {
    try { fs.closeSync(handle); } catch {}
  }
  console.error(error);
  process.exitCode = 1;
});
