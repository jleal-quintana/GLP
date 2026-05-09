import { brand } from '../branding/tokens';

export async function getOrAddSheet(context: Excel.RequestContext, name: string): Promise<Excel.Worksheet> {
  let sheet = context.workbook.worksheets.getItemOrNullObject(name);
  sheet.load('name');
  await context.sync();
  if (sheet.isNullObject) sheet = context.workbook.worksheets.add(name);
  sheet.visibility = Excel.SheetVisibility.visible;
  return sheet;
}

export function writeTitle(sheet: Excel.Worksheet, title: string, subtitle: string): void {
  sheet.getRange('A1:H1').merge(false);
  sheet.getRange('A1').values = [[title]];
  sheet.getRange('A1').format.font.bold = true;
  sheet.getRange('A1').format.font.size = 16;
  sheet.getRange('A1').format.font.color = brand.olive;
  sheet.getRange('A2:H2').merge(false);
  sheet.getRange('A2').values = [[subtitle]];
  sheet.getRange('A2').format.font.color = brand.muted;
}

export function writeTable(sheet: Excel.Worksheet, headerAddress: string, headers: string[], rows: (string | number)[][], label: string): void {
  validateRows(label, headerAddress, headers, rows);
  const header = sheet.getRange(headerAddress).getCell(0, 0).getResizedRange(0, headers.length - 1);
  header.values = [headers];
  header.format.fill.color = brand.olive;
  header.format.font.color = '#FFFFFF';
  header.format.font.bold = true;

  if (rows.length > 0) {
    const body = header.getOffsetRange(1, 0).getResizedRange(rows.length - 1, headers.length - 1);
    body.values = rows;
  }
  header.worksheet.getUsedRangeOrNullObject().format.autofitColumns();
}

export function addLineChart(sheet: Excel.Worksheet, source: Excel.Range, title: string, topLeft: string, bottomRight: string): void {
  const chart = sheet.charts.add(Excel.ChartType.line, source, Excel.ChartSeriesBy.columns);
  chart.title.text = title;
  chart.legend.position = Excel.ChartLegendPosition.bottom;
  chart.setPosition(topLeft, bottomRight);
}

function validateRows(label: string, headerAddress: string, headers: string[], rows: (string | number)[][]): void {
  for (let index = 0; index < rows.length; index++) {
    if (rows[index].length !== headers.length) {
      throw new Error(
        `${label} ${headerAddress}: fila ${index + 1} tiene ${rows[index].length} columnas; se esperaban ${headers.length}`,
      );
    }
  }
}
