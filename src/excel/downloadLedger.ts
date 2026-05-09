import { brand } from '../branding/tokens';
import type { AreaWorkbookPlan, CapituloIvDownloadEvent, ProductionRecord } from '../models/types';

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
      const header = sheet.getRange('A4:L4');
      header.values = [headers];
      header.format.fill.color = brand.olive;
      header.format.font.color = '#FFFFFF';
      header.format.font.bold = true;
      sheet.freezePanes.freezeRows(4);
    }

    sheet.getRange('A2').values = [[`Ultima area iniciada: ${plan.selection.areaId} - ${plan.selection.areaName}`]];
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

    sheet.getRangeByIndexes(nextRow, 0, 1, row.length).values = [row];
    sheet.getRange('A2').values = [[`Ultimo recurso: ${plan.selection.areaId} ${event.source} ${event.year ?? 'NC'} - ${event.matchedRows} coincidencias de ${event.scannedRows} filas`]];
    await context.sync();
  });
}

async function getOrAddSheet(context: Excel.RequestContext, name: string): Promise<Excel.Worksheet> {
  let sheet = context.workbook.worksheets.getItemOrNullObject(name);
  sheet.load('name');
  await context.sync();
  if (sheet.isNullObject) sheet = context.workbook.worksheets.add(name);
  sheet.visibility = Excel.SheetVisibility.visible;
  return sheet;
}

function writeTitle(sheet: Excel.Worksheet, title: string, subtitle: string): void {
  sheet.getRange('A1:H1').merge(false);
  sheet.getRange('A1').values = [[title]];
  sheet.getRange('A1').format.font.bold = true;
  sheet.getRange('A1').format.font.size = 16;
  sheet.getRange('A1').format.font.color = brand.olive;
  sheet.getRange('A2:H2').merge(false);
  sheet.getRange('A2').values = [[subtitle]];
  sheet.getRange('A2').format.font.color = brand.muted;
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
