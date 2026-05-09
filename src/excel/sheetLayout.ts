import { brand } from '../branding/tokens';

const TABLE_WRITE_BATCH_ROWS = 1000;

export async function getOrAddSheet(context: Excel.RequestContext, name: string): Promise<Excel.Worksheet> {
  let sheet = context.workbook.worksheets.getItemOrNullObject(name);
  sheet.load('name');
  await context.sync();
  if (sheet.isNullObject) sheet = context.workbook.worksheets.add(name);
  sheet.visibility = Excel.SheetVisibility.visible;
  return sheet;
}

export function writeTitle(sheet: Excel.Worksheet, title: string, subtitle: string): void {
  sheet.getRange('A1:H2').unmerge();
  writeMatrix(sheet, 'A1', [[title]], 'Titulo');
  sheet.getRange('A1').format.font.bold = true;
  sheet.getRange('A1').format.font.size = 16;
  sheet.getRange('A1').format.font.color = brand.olive;
  writeMatrix(sheet, 'A2', [[subtitle]], 'Subtitulo');
  sheet.getRange('A2').format.font.color = brand.muted;
}

export function writeMatrix(sheet: Excel.Worksheet, startAddress: string, values: (string | number)[][], label: string): Excel.Range {
  validateMatrix(label, values);
  const width = values[0]?.length ?? 0;
  const height = values.length;
  if (width === 0 || height === 0) {
    throw new Error(`${label} ${startAddress}: matriz vacia`);
  }
  const range = sheet.getRange(startAddress).getCell(0, 0).getResizedRange(height - 1, width - 1);
  range.values = values;
  return range;
}

export function writeTable(sheet: Excel.Worksheet, headerAddress: string, headers: string[], rows: (string | number)[][], label: string): void {
  validateRows(label, headers, rows);
  const header = writeMatrix(sheet, headerAddress, [headers], `${label} encabezado`);
  header.format.fill.color = brand.olive;
  header.format.font.color = '#FFFFFF';
  header.format.font.bold = true;

  if (rows.length > 0) {
    for (let offset = 0; offset < rows.length; offset += TABLE_WRITE_BATCH_ROWS) {
      const batch = rows.slice(offset, offset + TABLE_WRITE_BATCH_ROWS);
      const body = header.getOffsetRange(1 + offset, 0).getResizedRange(batch.length - 1, headers.length - 1);
      writeRangeValues(body, batch, `${label} filas ${offset + 1}-${offset + batch.length}`);
    }
  }
  header.worksheet.getUsedRangeOrNullObject().format.autofitColumns();
}

export function writeRangeValues(range: Excel.Range, values: (string | number)[][], label: string): void {
  validateMatrix(label, values);
  range.values = values;
}

export function addLineChart(sheet: Excel.Worksheet, source: Excel.Range, title: string, topLeft: string, bottomRight: string): void {
  const chart = sheet.charts.add(Excel.ChartType.line, source, Excel.ChartSeriesBy.columns);
  chart.title.text = title;
  chart.legend.position = Excel.ChartLegendPosition.bottom;
  chart.setPosition(topLeft, bottomRight);
}

function validateRows(label: string, headers: string[], rows: (string | number)[][]): void {
  for (let index = 0; index < rows.length; index++) {
    if (rows[index].length !== headers.length) {
      throw new Error(
        `${label}: fila ${index + 1} tiene ${rows[index].length} columnas; se esperaban ${headers.length}`,
      );
    }
  }
}

function validateMatrix(label: string, values: (string | number)[][]): void {
  if (values.length === 0) return;
  const width = values[0].length;
  for (let index = 0; index < values.length; index++) {
    if (values[index].length !== width) {
      throw new Error(`${label}: fila ${index + 1} tiene ${values[index].length} columnas; se esperaban ${width}`);
    }
  }
}
