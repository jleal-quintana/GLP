import type { DebugEntry } from '../models/types';
import { brand } from '../branding/tokens';
import { writeMatrix, writeRangeValues } from './sheetLayout';

export const DEBUG_SHEET_NAME = 'CapIV_Debug';

export function createDebugEntry(step: string, status: DebugEntry['status'], detail: string): DebugEntry {
  return {
    timestamp: new Date().toISOString(),
    step,
    status,
    detail,
  };
}

export async function appendDebug(entry: DebugEntry): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = await ensureDebugSheet(context);
    const usedRange = sheet.getUsedRangeOrNullObject();
    usedRange.load('rowCount');
    await context.sync();

    const nextRow = usedRange.isNullObject ? 1 : usedRange.rowCount;
    const range = sheet.getRangeByIndexes(nextRow, 0, 1, 4);
    writeRangeValues(range, [[entry.timestamp, entry.step, entry.status, entry.detail]], 'Fila debug');
    formatStatusRow(range, entry.status);
    await context.sync();
  });
}

export async function ensureDebugSheet(context: Excel.RequestContext): Promise<Excel.Worksheet> {
  const sheets = context.workbook.worksheets;
  let sheet = sheets.getItemOrNullObject(DEBUG_SHEET_NAME);
  sheet.load('name');
  await context.sync();

  if (sheet.isNullObject) {
    sheet = sheets.add(DEBUG_SHEET_NAME);
    const header = sheet.getRange('A1:D1');
    writeMatrix(sheet, 'A1', [['Timestamp', 'Paso', 'Estado', 'Detalle']], 'Encabezado debug');
    header.format.fill.color = brand.forest;
    header.format.font.color = '#FFFFFF';
    header.format.font.bold = true;
    sheet.getRange('A:D').format.autofitColumns();
    sheet.freezePanes.freezeRows(1);
  }

  sheet.visibility = Excel.SheetVisibility.visible;
  return sheet;
}

function formatStatusRow(range: Excel.Range, status: DebugEntry['status']): void {
  const colorByStatus: Record<DebugEntry['status'], string> = {
    info: '#FFFFFF',
    ok: '#EEF6E7',
    warning: '#FFF4D8',
    error: '#F8E1E1',
  };
  range.format.fill.color = colorByStatus[status];
}
