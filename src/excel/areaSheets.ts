import { brand } from '../branding/tokens';
import { forecastFormula, lastNonMissing, nextMonth } from '../domain/forecast';
import type { AreaWorkbookPlan, MonthlyAggregate, ProductionRecord } from '../models/types';
import { ensureDebugSheet } from './debugSheet';
import { areaSheetNames } from './names';

export async function writeAreaSheets(
  plan: AreaWorkbookPlan,
  records: ProductionRecord[],
  monthly: MonthlyAggregate[],
  warnings: string[],
  middleMissingPolicy: 'blank' | 'zero',
): Promise<void> {
  await Excel.run(async (context) => {
    await ensureDebugSheet(context);
    const names = areaSheetNames(plan.selection.areaId);
    if (plan.mode === 'regenerate') {
      await deleteIfExists(context, Object.values(names));
    }

    const hdp = await getOrAddSheet(context, names.hdp);
    const prono = await getOrAddSheet(context, names.prono);
    const pozos = await getOrAddSheet(context, names.pozos);
    const graficos = await getOrAddSheet(context, names.graficos);
    const detalle = await getOrAddSheet(context, names.detalle);

    writeHdpSheet(hdp, plan, monthly, warnings, middleMissingPolicy);
    writePronoSheet(prono, plan, monthly, middleMissingPolicy);
    writePozosSheet(pozos, plan, monthly, middleMissingPolicy);
    writeDetailSheet(detalle, records);
    writeChartsSheet(context, graficos, plan, monthly.length);

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

async function deleteIfExists(context: Excel.RequestContext, names: string[]): Promise<void> {
  for (const name of names) {
    const sheet = context.workbook.worksheets.getItemOrNullObject(name);
    sheet.load('name');
    await context.sync();
    if (!sheet.isNullObject) sheet.delete();
  }
}

function writeHdpSheet(
  sheet: Excel.Worksheet,
  plan: AreaWorkbookPlan,
  monthly: MonthlyAggregate[],
  warnings: string[],
  middleMissingPolicy: 'blank' | 'zero',
): void {
  sheet.getRange().clear();
  writeTitle(sheet, `Historico de produccion - ${plan.selection.areaId}`, plan.selection.areaName);
  sheet.getRange('A4:B8').values = [
    ['Provincia', plan.selection.province],
    ['Area', `${plan.selection.areaId} - ${plan.selection.areaName}`],
    ['Ultimo mes publicado', monthly.at(-1)?.date ?? 'Sin datos'],
    ['Warnings intermedios', warnings.length ? warnings.join(', ') : 'Sin warnings'],
    ['Politica faltantes intermedios', middleMissingPolicy],
  ];

  const headers = ['Fecha', 'Petroleo', 'Gas', 'Agua', 'Bruta', 'Agua iny.', 'Pozos petroleo', 'Pozos gas', 'Inyectores', 'Faltante'];
  const rows = monthly.map((m) => [
    m.date,
    maybeBlank(m, m.oil, middleMissingPolicy),
    maybeBlank(m, m.gas, middleMissingPolicy),
    maybeBlank(m, m.water, middleMissingPolicy),
    maybeBlank(m, m.gross, middleMissingPolicy),
    maybeBlank(m, m.waterInjection, middleMissingPolicy),
    maybeBlank(m, m.oilWells, middleMissingPolicy),
    maybeBlank(m, m.gasWells, middleMissingPolicy),
    maybeBlank(m, m.injectorWells, middleMissingPolicy),
    m.missingKind === 'leading' ? 'Inicio' : m.missingKind === 'middle' ? 'Intermedio' : '',
  ]);
  writeTable(sheet, 'A9:J9', headers, rows, `HDP ${plan.selection.areaId}`);
}

function writePronoSheet(
  sheet: Excel.Worksheet,
  plan: AreaWorkbookPlan,
  monthly: MonthlyAggregate[],
  middleMissingPolicy: 'blank' | 'zero',
): void {
  sheet.getRange().clear();
  writeTitle(sheet, `Pronostico de produccion - ${plan.selection.areaId}`, plan.selection.areaName);
  const override = plan.override;
  const gross = override?.grossMethod ?? plan.defaults.grossMethod;
  const oil = override?.oilMethod ?? plan.defaults.oilMethod;
  const gas = override?.gasMethod ?? plan.defaults.gasMethod;
  const takeInitial = override?.takeInitialFromHistory ?? plan.defaults.takeInitialFromHistory;

  sheet.getRange('A4:B10').values = [
    ['Metodo bruta', gross],
    ['Metodo petroleo', oil],
    ['Metodo gas', gas],
    ['Tomar inicial de historia', takeInitial ? 'Si' : 'No'],
    ['Horizonte anos', plan.defaults.horizonYears],
    ['Inicio pronostico', nextMonth(monthly.at(-1)?.date)],
    ['Nota', 'Celdas de supuestos editables. Actualizar resumen desde el taskpane.'],
  ];
  sheet.getRange('B4').dataValidation.rule = { list: { inCellDropDown: true, source: 'Constante,HypMod,Declinación Hip.,Declinación Exp.' } };
  sheet.getRange('B5').dataValidation.rule = { list: { inCellDropDown: true, source: 'Constante,HypMod,Declinación Hip.,Declinación Exp.,Rap Np' } };
  sheet.getRange('B6').dataValidation.rule = { list: { inCellDropDown: true, source: 'Constante,HypMod,Declinación Hip.,Declinación Exp.,RGP' } };
  sheet.getRange('B7').dataValidation.rule = { list: { inCellDropDown: true, source: 'Si,No' } };

  sheet.getRange('D4:E10').values = [
    ['Di bruta anual', 0.12],
    ['b bruta', 0.7],
    ['Di petroleo anual', 0.12],
    ['b petroleo', 0.7],
    ['Di gas anual', 0.12],
    ['b gas', 0.7],
    ['RGP constante', (lastNonMissing(monthly, 'gas') / Math.max(lastNonMissing(monthly, 'oil'), 1)) * 1000],
  ];

  const headers = ['Fecha', 'Petroleo hist.', 'Petroleo prono', 'Gas hist.', 'Gas prono', 'Bruta hist.', 'Bruta prono', 'Agua hist.', 'Agua prono', 'Wcut', 'RGP', 'RAP'];
  const rows: (string | number)[][] = monthly.map((m) => [
    m.date,
    maybeBlank(m, m.oil, middleMissingPolicy),
    '',
    maybeBlank(m, m.gas, middleMissingPolicy),
    '',
    maybeBlank(m, m.gross, middleMissingPolicy),
    '',
    maybeBlank(m, m.water, middleMissingPolicy),
    '',
    m.gross ? m.water / m.gross : '',
    m.oil ? (m.gas / m.oil) * 1000 : '',
    m.oil ? m.water / m.oil : '',
  ]);
  const last = monthly.at(-1);
  if (last) {
    const start = new Date(`${last.date}T00:00:00`);
    const months = plan.defaults.horizonYears * 12;
    const lastOil = lastNonMissing(monthly, 'oil');
    const lastGas = lastNonMissing(monthly, 'gas');
    const lastGross = lastNonMissing(monthly, 'gross');
    for (let i = 1; i <= months; i++) {
      const date = new Date(start);
      date.setMonth(date.getMonth() + i);
      const tYears = i / 12;
      const rowNumber = 12 + rows.length + 1;
      rows.push([
        date.toISOString().slice(0, 10),
        '',
        forecastFormula('B5', lastOil, 'E6', 'E7', tYears),
        '',
        forecastFormula('B6', lastGas, 'E8', 'E9', tYears, 'E10', `C${rowNumber}`),
        '',
        forecastFormula('B4', lastGross, 'E4', 'E5', tYears),
        '',
        `=MAX(G${rowNumber}-C${rowNumber},0)`,
        `=IF(G${rowNumber}>0,I${rowNumber}/G${rowNumber},"")`,
        `=IF(C${rowNumber}>0,E${rowNumber}/C${rowNumber}*1000,"")`,
        `=IF(C${rowNumber}>0,I${rowNumber}/C${rowNumber},"")`,
      ]);
    }
  }
  writeTable(sheet, 'A12:L12', headers, rows, `Prono ${plan.selection.areaId}`);
}

function writePozosSheet(
  sheet: Excel.Worksheet,
  plan: AreaWorkbookPlan,
  monthly: MonthlyAggregate[],
  middleMissingPolicy: 'blank' | 'zero',
): void {
  sheet.getRange().clear();
  writeTitle(sheet, 'Pronostico de pozos', 'Cantidad historica y supuestos editables');
  sheet.getRange('A4:B7').values = [
    ['Tomar pozos petroleo de historia', 'Si'],
    ['Tomar pozos gas de historia', 'Si'],
    ['Tomar inyectores de historia', 'Si'],
    ['Metodo', 'Declinacion exponencial'],
  ];
  sheet.getRange('B4:B6').dataValidation.rule = { list: { inCellDropDown: true, source: 'Si,No' } };
  sheet.getRange('D4:E6').values = [
    ['Di pozos petroleo anual', 0.04],
    ['Di pozos gas anual', 0.04],
    ['Di inyectores anual', 0.03],
  ];

  const headers = ['Fecha', 'Pozos petroleo', 'Pozos gas', 'Inyectores'];
  const rows = monthly.map((m) => [
    m.date,
    maybeBlank(m, m.oilWells, middleMissingPolicy),
    maybeBlank(m, m.gasWells, middleMissingPolicy),
    maybeBlank(m, m.injectorWells, middleMissingPolicy),
  ]);
  const last = monthly.at(-1);
  if (last) {
    const start = new Date(`${last.date}T00:00:00`);
    const months = plan.defaults.horizonYears * 12;
    const lastOilWells = lastNonMissing(monthly, 'oilWells');
    const lastGasWells = lastNonMissing(monthly, 'gasWells');
    const lastInjectors = lastNonMissing(monthly, 'injectorWells');
    for (let i = 1; i <= months; i++) {
      const date = new Date(start);
      date.setMonth(date.getMonth() + i);
      const t = Number((i / 12).toFixed(6));
      rows.push([
        date.toISOString().slice(0, 10),
        `=IF($B$4="Si",${lastOilWells}*EXP(-$E$4*${t}),${lastOilWells})`,
        `=IF($B$5="Si",${lastGasWells}*EXP(-$E$5*${t}),${lastGasWells})`,
        `=IF($B$6="Si",${lastInjectors}*EXP(-$E$6*${t}),${lastInjectors})`,
      ]);
    }
  }
  writeTable(sheet, 'A9:D9', headers, rows, 'Prono Pozos');
}

function writeDetailSheet(sheet: Excel.Worksheet, records: ProductionRecord[]): void {
  sheet.getRange().clear();
  writeTitle(sheet, 'Detalle Capitulo IV', 'Pozo-mes descargado');
  const headers = ['Anio', 'Mes', 'Area', 'Pozo', 'Id pozo', 'Petroleo', 'Gas', 'Agua', 'Agua iny.'];
  const rows = records.map((r) => [r.year, r.month, r.areaId, r.wellName, r.wellId, r.oil, r.gas, r.water, r.waterInjection]);
  writeTable(sheet, 'A4:I4', headers, rows, 'Detalle Capitulo IV');
}

function writeChartsSheet(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  plan: AreaWorkbookPlan,
  historyMonths: number,
): void {
  sheet.getRange().clear();
  writeTitle(sheet, `Graficos - ${plan.selection.areaId}`, 'Graficos nativos de Excel');
  const names = areaSheetNames(plan.selection.areaId);
  const prono = context.workbook.worksheets.getItem(names.prono);
  const pozos = context.workbook.worksheets.getItem(names.pozos);
  const forecastMonths = plan.defaults.horizonYears * 12;
  const pronoRows = Math.max(2, historyMonths + forecastMonths + 1);
  const pozosRows = Math.max(2, historyMonths + forecastMonths + 1);
  addLineChart(sheet, prono.getRangeByIndexes(11, 0, pronoRows, 3), 'Petroleo', 'A4', 'H22');
  addLineChart(sheet, prono.getRangeByIndexes(11, 0, pronoRows, 5), 'Gas', 'I4', 'P22');
  addLineChart(sheet, prono.getRangeByIndexes(11, 0, pronoRows, 7), 'Bruta', 'A24', 'H42');
  addLineChart(sheet, prono.getRangeByIndexes(11, 0, pronoRows, 12), 'Ratios', 'I24', 'P42');
  addLineChart(sheet, pozos.getRangeByIndexes(8, 0, pozosRows, 4), 'Pozos', 'A44', 'H62');
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

function writeTable(sheet: Excel.Worksheet, headerAddress: string, headers: string[], rows: (string | number)[][], label: string): void {
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

function validateRows(label: string, headerAddress: string, headers: string[], rows: (string | number)[][]): void {
  for (let index = 0; index < rows.length; index++) {
    if (rows[index].length !== headers.length) {
      throw new Error(
        `${label} ${headerAddress}: fila ${index + 1} tiene ${rows[index].length} columnas; se esperaban ${headers.length}`,
      );
    }
  }
}

function addLineChart(sheet: Excel.Worksheet, source: Excel.Range, title: string, topLeft: string, bottomRight: string): void {
  const chart = sheet.charts.add(Excel.ChartType.line, source, Excel.ChartSeriesBy.columns);
  chart.title.text = title;
  chart.legend.position = Excel.ChartLegendPosition.bottom;
  chart.setPosition(topLeft, bottomRight);
}

function maybeBlank(month: MonthlyAggregate, value: number, middleMissingPolicy: 'blank' | 'zero'): string | number {
  if (month.missingKind === 'middle' && middleMissingPolicy === 'blank') return '';
  return value;
}
