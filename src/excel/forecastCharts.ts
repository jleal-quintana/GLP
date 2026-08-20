import { brand } from '../branding/tokens';
import { writeMatrix } from './sheetLayout';

type ChartKind = 'line' | 'scatter';

interface SeriesStyle {
  color: string;
  lineStyle?: 'Continuous' | 'Dash';
}

interface ChartTable {
  headers: string[];
  formulas: string[][];
}

interface ForecastChartDefinition {
  id: string;
  title: string;
  section: 'Producción mensual' | 'Indicadores de comportamiento' | 'Acumuladas y actividad' | 'Diagnóstico de secundaria';
  topLeft: string;
  bottomRight: string;
  kind: ChartKind;
  xAxisTitle: string;
  yAxisTitle: string;
  yNumberFormat: string;
  series: SeriesStyle[];
  buildTable: (sources: ChartSources, totalMonths: number, historyMonths: number) => ChartTable;
}

interface ChartSources {
  prono: string;
  pozos: string;
  hdp: string;
}

export interface ForecastChartCatalogItem {
  id: string;
  title: string;
  section: ForecastChartDefinition['section'];
}

const HISTORY = { color: brand.marine, lineStyle: 'Continuous' as const };
const FORECAST = { color: brand.olive, lineStyle: 'Dash' as const };

const DEFINITIONS: ForecastChartDefinition[] = [
  phaseChart('oil', 'Petróleo', 'B', 'C', 'Producción mensual', 'A6', 'H24', 'm³/mes'),
  phaseChart('gas', 'Gas', 'D', 'E', 'Producción mensual', 'I6', 'P24', 'miles de m³/mes'),
  phaseChart('gross', 'Producción bruta', 'F', 'G', 'Producción mensual', 'A26', 'H44', 'm³/mes'),
  phaseChart('water', 'Agua producida', 'H', 'I', 'Producción mensual', 'I26', 'P44', 'm³/mes'),
  ratioChart('water-cut', 'Corte de agua', 'J', 'Indicadores de comportamiento', 'A48', 'H66', 'fracción', '0%'),
  ratioChart('rgp', 'RGP', 'K', 'Indicadores de comportamiento', 'I48', 'P66', 'm³ gas / m³ petróleo', '#,##0'),
  ratioChart('rap', 'RAP / WOR', 'L', 'Indicadores de comportamiento', 'A68', 'H86', 'm³ agua / m³ petróleo', '0.0'),
  {
    id: 'water-injection',
    title: 'Agua inyectada',
    section: 'Indicadores de comportamiento',
    topLeft: 'I68',
    bottomRight: 'P86',
    kind: 'line',
    xAxisTitle: 'Mes',
    yAxisTitle: 'm³/mes',
    yNumberFormat: '#,##0',
    series: [{ color: brand.olive }],
    buildTable: (sources, _totalMonths, historyMonths) => directTable(sources.hdp, 10, historyMonths, [
      ['Fecha', 'A'],
      ['Agua inyectada', 'F'],
    ]),
  },
  {
    id: 'liquid-cumulatives',
    title: 'Acumuladas líquidas',
    section: 'Acumuladas y actividad',
    topLeft: 'A90',
    bottomRight: 'H108',
    kind: 'line',
    xAxisTitle: 'Mes',
    yAxisTitle: 'miles de m³',
    yNumberFormat: '#,##0',
    series: [
      { color: brand.olive },
      { color: brand.marine },
      { color: '#AD977D' },
    ],
    buildTable: cumulativeLiquidsTable,
  },
  {
    id: 'wells',
    title: 'Pozos activos',
    section: 'Acumuladas y actividad',
    topLeft: 'I90',
    bottomRight: 'P108',
    kind: 'line',
    xAxisTitle: 'Mes',
    yAxisTitle: 'cantidad de pozos',
    yNumberFormat: '0',
    series: [
      { color: brand.olive },
      { color: brand.marine },
      { color: '#AD977D' },
    ],
    buildTable: (sources, totalMonths) => directTable(sources.pozos, 10, totalMonths, [
      ['Fecha', 'A'],
      ['Pozos petróleo', 'B'],
      ['Pozos gas', 'C'],
      ['Inyectores', 'D'],
    ]),
  },
  {
    id: 'rap-vs-np',
    title: 'RAP vs. Np',
    section: 'Diagnóstico de secundaria',
    topLeft: 'A112',
    bottomRight: 'P132',
    kind: 'scatter',
    xAxisTitle: 'Np acumulado (miles de m³)',
    yAxisTitle: 'RAP (m³/m³)',
    yNumberFormat: '0.0',
    series: [{ color: brand.marine }],
    buildTable: rapVsNpTable,
  },
];

export function forecastChartCatalog(): ForecastChartCatalogItem[] {
  return DEFINITIONS.map(({ id, title, section }) => ({ id, title, section }));
}

export async function writeForecastCharts(
  sheet: Excel.Worksheet,
  sourceSheetNames: ChartSources,
  totalMonths: number,
  historyMonths: number,
  context?: Excel.RequestContext,
): Promise<void> {
  const sources: ChartSources = {
    prono: quoteSheetName(sourceSheetNames.prono),
    pozos: quoteSheetName(sourceSheetNames.pozos),
    hdp: quoteSheetName(sourceSheetNames.hdp),
  };
  // Con contexto, sincroniza tras cada bloque para que un fallo quede atribuido
  // al gráfico exacto en la hoja Debug en vez de al paso "contenido" completo.
  const checkpoint = async (label: string) => {
    if (!context) return;
    try {
      await context.sync();
    } catch (error) {
      const detail = error instanceof Error ? error.message : String(error);
      throw new Error(`${label}: ${detail}`);
    }
  };

  writeSectionHeader(sheet, 'A4:P4', 'PRODUCCIÓN MENSUAL');
  writeSectionHeader(sheet, 'A46:P46', 'INDICADORES DE COMPORTAMIENTO');
  writeSectionHeader(sheet, 'A88:P88', 'ACUMULADAS Y ACTIVIDAD');
  writeSectionHeader(sheet, 'A110:P110', 'DIAGNÓSTICO DE SECUNDARIA');
  await checkpoint('Encabezados de sección');

  let helperColumn = 25;
  for (const definition of DEFINITIONS) {
    const table = definition.buildTable(sources, totalMonths, historyMonths);
    const source = writeHelperTable(sheet, helperColumn, table, definition.id);
    addBrandedChart(sheet, source, definition);
    helperColumn += table.headers.length + 1;
    await checkpoint(`Gráfico ${definition.id} (${definition.title})`);
  }

  sheet.getRangeByIndexes(0, 25, Math.max(totalMonths + 1, historyMonths + 1), helperColumn - 25).columnHidden = true;
  sheet.getRange('A:P').format.columnWidth = 11;
  await checkpoint('Ocultar columnas auxiliares');
}

function phaseChart(
  id: string,
  title: string,
  historyColumn: string,
  forecastColumn: string,
  section: ForecastChartDefinition['section'],
  topLeft: string,
  bottomRight: string,
  yAxisTitle: string,
): ForecastChartDefinition {
  return {
    id,
    title,
    section,
    topLeft,
    bottomRight,
    kind: 'line',
    xAxisTitle: 'Mes',
    yAxisTitle,
    yNumberFormat: '#,##0',
    series: [HISTORY, FORECAST],
    buildTable: (sources, totalMonths) => directTable(sources.prono, 13, totalMonths, [
      ['Fecha', 'A'],
      ['Histórico', historyColumn],
      ['Pronóstico', forecastColumn],
    ]),
  };
}

function ratioChart(
  id: string,
  title: string,
  valueColumn: string,
  section: ForecastChartDefinition['section'],
  topLeft: string,
  bottomRight: string,
  yAxisTitle: string,
  yNumberFormat: string,
): ForecastChartDefinition {
  return {
    id,
    title,
    section,
    topLeft,
    bottomRight,
    kind: 'line',
    xAxisTitle: 'Mes',
    yAxisTitle,
    yNumberFormat,
    series: [HISTORY, FORECAST],
    buildTable: (sources, totalMonths) => splitRatioTable(sources.prono, totalMonths, valueColumn),
  };
}

function directTable(
  sheetName: string,
  firstRow: number,
  rowCount: number,
  columns: Array<[header: string, column: string]>,
): ChartTable {
  const formulas = Array.from({ length: rowCount }, (_, index) => columns.map(([header, column], columnIndex) => {
    const source = sourceCell(sheetName, column, firstRow + index);
    if (columnIndex === 0) return `=${source}`;
    return `=IF(${source}="",NA(),${source})`;
  }));
  return { headers: columns.map(([header]) => header), formulas };
}

function splitRatioTable(sheetName: string, rowCount: number, valueColumn: string): ChartTable {
  const formulas = Array.from({ length: rowCount }, (_, index) => {
    const row = 13 + index;
    const date = sourceCell(sheetName, 'A', row);
    const oilHistory = sourceCell(sheetName, 'B', row);
    const oilForecast = sourceCell(sheetName, 'C', row);
    const value = sourceCell(sheetName, valueColumn, row);
    return [
      `=${date}`,
      `=IF(OR(${oilHistory}="",${value}=""),NA(),${value})`,
      `=IF(OR(${oilForecast}="",${value}=""),NA(),${value})`,
    ];
  });
  return { headers: ['Fecha', 'Histórico', 'Pronóstico'], formulas };
}

function cumulativeLiquidsTable(sources: ChartSources, totalMonths: number, historyMonths: number): ChartTable {
  const formulas = Array.from({ length: totalMonths }, (_, index) => {
    const pronoRow = 13 + index;
    const hdpRow = 10 + index;
    const date = sourceCell(sources.prono, 'A', pronoRow);
    const oilHistory = sourceRange(sources.prono, 'B', 13, pronoRow);
    const oilForecast = sourceRange(sources.prono, 'C', 13, pronoRow);
    const waterHistory = sourceRange(sources.prono, 'H', 13, pronoRow);
    const waterForecast = sourceRange(sources.prono, 'I', 13, pronoRow);
    const injection = sourceRange(sources.hdp, 'F', 10, hdpRow);
    return [
      `=${date}`,
      `=SUM(${oilHistory},${oilForecast})/1000`,
      `=SUM(${waterHistory},${waterForecast})/1000`,
      index < historyMonths ? `=SUM(${injection})/1000` : '=NA()',
    ];
  });
  return { headers: ['Fecha', 'Np', 'Wp', 'Wi'], formulas };
}

function rapVsNpTable(sources: ChartSources, totalMonths: number): ChartTable {
  const formulas = Array.from({ length: totalMonths }, (_, index) => {
    const row = 13 + index;
    const oilHistory = sourceRange(sources.prono, 'B', 13, row);
    const oilForecast = sourceRange(sources.prono, 'C', 13, row);
    const rap = sourceCell(sources.prono, 'L', row);
    return [
      `=SUM(${oilHistory},${oilForecast})/1000`,
      `=IF(${rap}="",NA(),${rap})`,
    ];
  });
  return { headers: ['Np', 'RAP'], formulas };
}

function writeHelperTable(sheet: Excel.Worksheet, startColumn: number, table: ChartTable, id: string): Excel.Range {
  const header = writeMatrix(sheet, cellAddress(0, startColumn), [table.headers], `Fuente gráfico ${id}`);
  header.format.fill.color = brand.slate;
  header.format.font.bold = true;
  if (table.formulas.length > 0) {
    sheet.getRangeByIndexes(1, startColumn, table.formulas.length, table.headers.length).formulas = table.formulas;
  }
  return sheet.getRangeByIndexes(0, startColumn, table.formulas.length + 1, table.headers.length);
}

function addBrandedChart(sheet: Excel.Worksheet, source: Excel.Range, definition: ForecastChartDefinition): void {
  const type = definition.kind === 'scatter' ? 'XYScatterLinesNoMarkers' : 'Line';
  const chart = sheet.charts.add(type, source, 'Columns');
  chart.name = `CapIV_${definition.id}`;
  chart.title.text = definition.title;
  chart.title.format.font.name = 'Montserrat';
  chart.title.format.font.size = 12;
  chart.title.format.font.bold = true;
  chart.title.format.font.color = brand.forest;
  chart.legend.position = 'Bottom';
  chart.legend.visible = definition.series.length > 1;
  chart.legend.format.font.name = 'Inter';
  chart.legend.format.font.size = 8;
  chart.displayBlanksAs = 'NotPlotted';
  chart.plotVisibleOnly = false;
  chart.setPosition(definition.topLeft, definition.bottomRight);
  chart.format.fill.setSolidColor('#FFFFFF');
  chart.plotArea.format.fill.setSolidColor('#FFFFFF');

  const xAxis = chart.axes.categoryAxis;
  const yAxis = chart.axes.valueAxis;
  xAxis.title.text = definition.xAxisTitle;
  xAxis.title.visible = true;
  xAxis.title.format.font.name = 'Inter';
  xAxis.title.format.font.size = 8;
  yAxis.title.text = definition.yAxisTitle;
  yAxis.title.visible = true;
  yAxis.title.format.font.name = 'Inter';
  yAxis.title.format.font.size = 8;
  yAxis.minimum = 0;
  yAxis.numberFormat = definition.yNumberFormat;
  yAxis.majorGridlines.format.line.color = brand.slate;
  yAxis.majorGridlines.format.line.weight = 0.75;

  if (definition.id === 'water-cut') yAxis.maximum = 1;

  definition.series.forEach((style, index) => {
    const series = chart.series.getItemAt(index);
    series.markerStyle = 'None';
    series.format.line.color = style.color;
    series.format.line.weight = 2;
    series.format.line.lineStyle = style.lineStyle ?? 'Continuous';
  });
}

function writeSectionHeader(sheet: Excel.Worksheet, address: string, label: string): void {
  const range = sheet.getRange(address);
  range.unmerge();
  // El valor va en la celda superior izquierda ANTES del merge: asignar un array 1×1
  // a un rango mergeado de 1×16 dispara un error de dimensiones en Office.js.
  range.getCell(0, 0).values = [[label]];
  range.merge();
  range.format.fill.color = brand.forest;
  range.format.font.color = '#FFFFFF';
  range.format.font.bold = true;
  range.format.font.size = 10;
  range.format.verticalAlignment = 'Center';
  range.format.rowHeight = 23;
}

function sourceCell(sheetName: string, column: string, row: number, absoluteRow = true): string {
  return `${sheetName}!$${column}${absoluteRow ? '$' : ''}${row}`;
}

function sourceRange(sheetName: string, column: string, firstRow: number, lastRow: number): string {
  return `${sheetName}!$${column}$${firstRow}:$${column}${lastRow}`;
}

function quoteSheetName(name: string): string {
  return `'${name.replace(/'/g, "''")}'`;
}

function cellAddress(rowIndex: number, columnIndex: number): string {
  let column = '';
  let value = columnIndex + 1;
  while (value > 0) {
    value--;
    column = String.fromCharCode(65 + (value % 26)) + column;
    value = Math.floor(value / 26);
  }
  return `${column}${rowIndex + 1}`;
}
