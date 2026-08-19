import type {
  AreaWorkbookPlan,
  DataOutputTarget,
  MonthlyAggregate,
  OverwriteDecisionHandler,
  ProductionRecord,
  WorkbookAreaData,
} from '../models/types';

export interface DownloadedArea {
  plan: AreaWorkbookPlan;
  data: WorkbookAreaData;
  records: ProductionRecord[];
}

export async function captureSelectedCell(granularity: DataOutputTarget['granularity']): Promise<DataOutputTarget> {
  return Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    range.load(['rowIndex', 'columnIndex']);
    sheet.load('name');
    await context.sync();
    const startAddress = `${excelColumn(range.columnIndex)}${range.rowIndex + 1}`;
    return {
      sheetName: sheet.name,
      startAddress,
      granularity,
      tableName: createTableName(sheet.name, startAddress),
    };
  });
}

export async function writeDatabaseTable(
  target: DataOutputTarget,
  downloads: DownloadedArea[],
  onOverwrite?: OverwriteDecisionHandler,
): Promise<{ rangeAddress: string; rowCount: number }> {
  const matrix = buildDatabaseMatrix(target.granularity, downloads);
  const inspection = await inspectDestination(target, matrix.length, matrix[0].length);
  if (inspection.occupiedCells > 0 || inspection.overlappingTables.length > 0) {
    const accepted = onOverwrite ? await onOverwrite(inspection) : false;
    if (!accepted) throw new Error('La escritura fue cancelada para no sobrescribir datos existentes.');
  }

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(target.sheetName);
    const start = sheet.getRange(target.startAddress);
    start.load(['rowIndex', 'columnIndex']);
    const tables = sheet.tables;
    tables.load('items/name');
    await context.sync();

    const tableRanges = tables.items.map((table) => {
      const range = table.getRange();
      range.load(['rowIndex', 'columnIndex', 'rowCount', 'columnCount']);
      return { table, range };
    });
    await context.sync();

    const rowIndex = start.rowIndex;
    const columnIndex = start.columnIndex;
    for (const item of tableRanges) {
      if (rangesOverlap(rowIndex, columnIndex, matrix.length, matrix[0].length, item.range)) {
        item.range.clear(Excel.ClearApplyTo.all);
        item.table.delete();
      }
    }
    const output = sheet.getRangeByIndexes(rowIndex, columnIndex, matrix.length, matrix[0].length);
    output.clear(Excel.ClearApplyTo.all);
    output.values = matrix;
    const table = sheet.tables.add(output, true);
    table.name = target.tableName;
    table.showBandedRows = true;
    table.style = 'TableStyleMedium4';
    const header = table.getHeaderRowRange();
    header.format.fill.color = '#33492D';
    header.format.font.color = '#FFFFFF';
    header.format.font.bold = true;
    output.format.autofitColumns();
    output.format.autofitRows();
    await context.sync();
  });

  return {
    rangeAddress: inspection.rangeAddress,
    rowCount: matrix.length - 1,
  };
}

async function inspectDestination(target: DataOutputTarget, rowCount: number, columnCount: number) {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(target.sheetName);
    const start = sheet.getRange(target.startAddress);
    start.load(['rowIndex', 'columnIndex']);
    const tables = sheet.tables;
    tables.load('items/name');
    await context.sync();
    const destination = sheet.getRangeByIndexes(start.rowIndex, start.columnIndex, rowCount, columnCount);
    destination.load(['address', 'values']);
    const tableRanges = tables.items.map((table) => {
      const range = table.getRange();
      range.load(['rowIndex', 'columnIndex', 'rowCount', 'columnCount']);
      return { name: table.name, range };
    });
    await context.sync();
    const occupiedCells = destination.values.flat().filter((value) => value !== '' && value !== null).length;
    return {
      sheetName: target.sheetName,
      rangeAddress: destination.address,
      occupiedCells,
      overlappingTables: tableRanges
        .filter((item) => rangesOverlap(start.rowIndex, start.columnIndex, rowCount, columnCount, item.range))
        .map((item) => item.name),
    };
  });
}

export function buildDatabaseMatrix(granularity: DataOutputTarget['granularity'], downloads: DownloadedArea[]): (string | number)[][] {
  return granularity === 'area' ? areaMatrix(downloads) : wellMatrix(downloads);
}

function areaMatrix(downloads: DownloadedArea[]): (string | number)[][] {
  const headers = ['Fecha', 'Año', 'Mes', 'Código área', 'Área', 'Provincia', 'Petróleo', 'Gas', 'Agua', 'Bruta', 'Agua inyectada', 'Pozos petróleo', 'Pozos gas', 'Inyectores'];
  const rows = downloads.flatMap(({ plan, data }) => data.monthly.map((month) => areaRow(plan, month)));
  return [headers, ...rows];
}

function areaRow(plan: AreaWorkbookPlan, month: MonthlyAggregate): (string | number)[] {
  return [
    month.date,
    month.year,
    month.month,
    plan.selection.areaId,
    plan.selection.areaName,
    plan.selection.province,
    month.oil,
    month.gas,
    month.water,
    month.gross,
    month.waterInjection,
    month.oilWells,
    month.gasWells,
    month.injectorWells,
  ];
}

function wellMatrix(downloads: DownloadedArea[]): (string | number)[][] {
  const headers = ['Año', 'Mes', 'Código área', 'Área', 'Provincia', 'ID pozo', 'Pozo', 'Petróleo', 'Gas', 'Agua', 'Agua inyectada'];
  const rows = downloads.flatMap(({ plan, records }) => records.map((record) => [
    record.year,
    record.month,
    record.areaId,
    record.areaName,
    plan.selection.province,
    record.wellId,
    record.wellName,
    record.oil,
    record.gas,
    record.water,
    record.waterInjection,
  ]));
  return [headers, ...rows];
}

function createTableName(sheetName: string, address: string): string {
  const suffix = `${sheetName}_${address}`.replace(/[^A-Za-z0-9_]/g, '_').slice(0, 180);
  return `CapIV_Datos_${suffix}`;
}

export function rangesOverlap(
  rowIndex: number,
  columnIndex: number,
  rowCount: number,
  columnCount: number,
  other: { rowIndex: number; columnIndex: number; rowCount: number; columnCount: number },
): boolean {
  return !(
    rowIndex + rowCount <= other.rowIndex ||
    other.rowIndex + other.rowCount <= rowIndex ||
    columnIndex + columnCount <= other.columnIndex ||
    other.columnIndex + other.columnCount <= columnIndex
  );
}

function excelColumn(index: number): string {
  let value = index + 1;
  let output = '';
  while (value > 0) {
    value--;
    output = String.fromCharCode(65 + (value % 26)) + output;
    value = Math.floor(value / 26);
  }
  return output;
}
