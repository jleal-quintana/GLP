import type { AreaWorkbookPlan, CapituloIvDownloadEvent, ProductionRecord } from '../models/types';
import { getOrAddSheet, writeMatrix, writeRangeValues, writeTable, writeTitle } from './sheetLayout';

export const DOWNLOAD_LEDGER_SHEET = 'CapIV_Descarga';

export async function ensureDownloadLedger(plan: AreaWorkbookPlan): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = await getOrAddSheet(context, DOWNLOAD_LEDGER_SHEET);
    const used = sheet.getUsedRangeOrNullObject();
    used.load('rowCount');
    await context.sync();

    if (used.isNullObject) {
      writeTitle(sheet, 'Descarga Capitulo IV', 'Resumen incremental por recurso descargado');
      const headers = ['Timestamp', 'Area', 'Nombre', 'Evento', 'Fuente', 'Anio recurso', 'Filas leidas', 'Filas area', 'Petroleo', 'Gas', 'Agua', 'Agua iny.'];
      writeTable(sheet, 'A4:L4', headers, [], 'Descarga Capitulo IV');
      sheet.freezePanes.freezeRows(4);
    }

    writeMatrix(sheet, 'A2', [[`Ultima area iniciada: ${plan.selection.areaId} - ${plan.selection.areaName}`]], 'Estado descarga');
    await context.sync();
  });
}

export async function appendDownloadLedgerEvent(plan: AreaWorkbookPlan, event: CapituloIvDownloadEvent): Promise<void> {
  if (event.type !== 'resource_completed') return;

  await Excel.run(async (context) => {
    const sheet = await getOrAddSheet(context, DOWNLOAD_LEDGER_SHEET);
    const used = sheet.getUsedRangeOrNullObject();
    used.load('rowCount');
    await context.sync();

    const nextRow = used.isNullObject ? 4 : used.rowCount;
    const totals = summarizeRecords(event.records);
    const row: (string | number)[] = [
      new Date().toISOString(),
      plan.selection.areaId,
      plan.selection.areaName,
      'Recurso descargado',
      event.source,
      event.year ?? '',
      event.scannedRows,
      event.matchedRows,
      totals.oil,
      totals.gas,
      totals.water,
      totals.waterInjection,
    ];

    writeRangeValues(sheet.getRangeByIndexes(nextRow, 0, 1, row.length), [row], 'Fila descarga');
    writeMatrix(sheet, 'A2', [[`Ultimo recurso: ${plan.selection.areaId} ${event.source} ${event.year ?? 'NC'} - ${event.matchedRows} coincidencias de ${event.scannedRows} filas`]], 'Estado descarga');
    await context.sync();
  });
}

function summarizeRecords(records: ProductionRecord[]): Pick<ProductionRecord, 'oil' | 'gas' | 'water' | 'waterInjection'> {
  return records.reduce(
    (total, record) => ({
      oil: total.oil + record.oil,
      gas: total.gas + record.gas,
      water: total.water + record.water,
      waterInjection: total.waterInjection + record.waterInjection,
    }),
    { oil: 0, gas: 0, water: 0, waterInjection: 0 },
  );
}
