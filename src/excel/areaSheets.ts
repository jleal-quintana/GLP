import { forecastFormula, lastNonMissing, nextMonth } from '../domain/forecast';
import type { AreaWorkbookPlan, MonthlyAggregate, ProductionRecord } from '../models/types';
import { appendDebug, createDebugEntry, ensureDebugSheet } from './debugSheet';
import { areaSheetNames } from './names';
import { addLineChart, getOrAddSheet, writeMatrix, writeTable, writeTitle } from './sheetLayout';

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
    await context.sync();
  });

  await writeAreaSheet(plan.selection.areaId, 'HDP', async (context) => {
    const hdp = await getOrAddSheet(context, areaSheetNames(plan.selection.areaId).hdp);
    writeHdpSheet(hdp, plan, monthly, warnings, middleMissingPolicy);
  });
  await writeAreaSheet(plan.selection.areaId, 'Prono', async (context) => {
    const prono = await getOrAddSheet(context, areaSheetNames(plan.selection.areaId).prono);
    writePronoSheet(prono, plan, monthly, middleMissingPolicy);
  });
  await writeAreaSheet(plan.selection.areaId, 'Pozos', async (context) => {
    const pozos = await getOrAddSheet(context, areaSheetNames(plan.selection.areaId).pozos);
    writePozosSheet(pozos, plan, monthly, middleMissingPolicy);
  });
  await writeAreaSheet(plan.selection.areaId, 'Detalle', async (context) => {
    const detalle = await getOrAddSheet(context, areaSheetNames(plan.selection.areaId).detalle);
    writeDetailSheet(detalle, records);
  });
  await writeAreaSheet(plan.selection.areaId, 'Graficos', async (context) => {
    const graficos = await getOrAddSheet(context, areaSheetNames(plan.selection.areaId).graficos);
    writeChartsSheet(context, graficos, plan, monthly.length);
  });
}

async function writeAreaSheet(areaId: string, sheetStep: string, writer: (context: Excel.RequestContext) => Promise<void> | void): Promise<void> {
  await appendDebug(createDebugEntry(areaId, 'info', `Escribiendo hoja ${sheetStep}`));
  try {
    await Excel.run(async (context) => {
      await writer(context);
      await context.sync();
    });
    await appendDebug(createDebugEntry(areaId, 'ok', `Hoja ${sheetStep} escrita`));
  } catch (error) {
    const detail = error instanceof Error ? error.message : String(error);
    await appendDebug(createDebugEntry(areaId, 'error', `Fallo hoja ${sheetStep}: ${detail}`));
    throw error;
  }
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
  writeMatrix(sheet, 'A4', [
    ['Provincia', plan.selection.province],
    ['Area', `${plan.selection.areaId} - ${plan.selection.areaName}`],
    ['Ultimo mes publicado', monthly.at(-1)?.date ?? 'Sin datos'],
    ['Warnings intermedios', warnings.length ? warnings.join(', ') : 'Sin warnings'],
    ['Politica faltantes intermedios', middleMissingPolicy],
  ], `Parametros HDP ${plan.selection.areaId}`);

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

  writeMatrix(sheet, 'A4', [
    ['Metodo bruta', gross],
    ['Metodo petroleo', oil],
    ['Metodo gas', gas],
    ['Tomar inicial de historia', takeInitial ? 'Si' : 'No'],
    ['Horizonte anos', plan.defaults.horizonYears],
    ['Inicio pronostico', nextMonth(monthly.at(-1)?.date)],
    ['Nota', 'Celdas de supuestos editables. Actualizar resumen desde el taskpane.'],
  ], `Parametros Prono ${plan.selection.areaId}`);
  sheet.getRange('B4').dataValidation.rule = { list: { inCellDropDown: true, source: 'Constante,HypMod,Declinación Hip.,Declinación Exp.' } };
  sheet.getRange('B5').dataValidation.rule = { list: { inCellDropDown: true, source: 'Constante,HypMod,Declinación Hip.,Declinación Exp.,Rap Np' } };
  sheet.getRange('B6').dataValidation.rule = { list: { inCellDropDown: true, source: 'Constante,HypMod,Declinación Hip.,Declinación Exp.,RGP' } };
  sheet.getRange('B7').dataValidation.rule = { list: { inCellDropDown: true, source: 'Si,No' } };

  writeMatrix(sheet, 'D4', [
    ['Di bruta anual', 0.12],
    ['b bruta', 0.7],
    ['Di petroleo anual', 0.12],
    ['b petroleo', 0.7],
    ['Di gas anual', 0.12],
    ['b gas', 0.7],
    ['RGP constante', (lastNonMissing(monthly, 'gas') / Math.max(lastNonMissing(monthly, 'oil'), 1)) * 1000],
  ], `Supuestos Prono ${plan.selection.areaId}`);

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
  writeMatrix(sheet, 'A4', [
    ['Tomar pozos petroleo de historia', 'Si'],
    ['Tomar pozos gas de historia', 'Si'],
    ['Tomar inyectores de historia', 'Si'],
    ['Metodo', 'Declinacion exponencial'],
  ], `Parametros Pozos ${plan.selection.areaId}`);
  sheet.getRange('B4:B6').dataValidation.rule = { list: { inCellDropDown: true, source: 'Si,No' } };
  writeMatrix(sheet, 'D4', [
    ['Di pozos petroleo anual', 0.04],
    ['Di pozos gas anual', 0.04],
    ['Di inyectores anual', 0.03],
  ], `Supuestos Pozos ${plan.selection.areaId}`);

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

function maybeBlank(month: MonthlyAggregate, value: number, middleMissingPolicy: 'blank' | 'zero'): string | number {
  if (month.missingKind === 'middle' && middleMissingPolicy === 'blank') return '';
  return value;
}
