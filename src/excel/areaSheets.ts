import { DEFAULT_DECLINE, forecastFormula, lastNonMissing, nextMonth } from '../domain/forecast';
import type { AreaForecastOverrideField, AreaWorkbookPlan, MonthlyAggregate, ProductionRecord } from '../models/types';
import { appendDebug, createDebugEntry, ensureDebugSheet } from './debugSheet';
import { writeForecastCharts } from './forecastCharts';
import { areaSheetNames } from './names';
import { getOrAddSheet, writeMatrix, writeTable, writeTitle } from './sheetLayout';

interface PreservedSettings {
  prono?: {
    methods: (string | number)[][];
    decline: (string | number)[][];
    initials: (string | number)[][];
  };
  pozos?: {
    switches: (string | number)[][];
    decline: (string | number)[][];
  };
}

export async function writeAreaSheets(
  plan: AreaWorkbookPlan,
  records: ProductionRecord[],
  monthly: MonthlyAggregate[],
  warnings: string[],
  middleMissingPolicy: 'blank' | 'zero',
): Promise<void> {
  await writeAreaDataSheets(plan, records, monthly, warnings, middleMissingPolicy);
  await writeAreaForecastSheets(plan, monthly, middleMissingPolicy);
}

export async function writeAreaDataSheets(
  plan: AreaWorkbookPlan,
  records: ProductionRecord[],
  monthly: MonthlyAggregate[],
  warnings: string[],
  middleMissingPolicy: 'blank' | 'zero',
): Promise<void> {
  const names = areaSheetNames(plan.selection.areaId);
  await Excel.run(async (context) => {
    await ensureDebugSheet(context);
    if (plan.mode === 'regenerate') {
      await deleteIfExists(context, [names.hdp, names.detalle]);
    }
    await context.sync();
  });
  await prepareAreaSheet(plan.selection.areaId, 'HDP', names.hdp);
  await writeAreaSheetStep(plan.selection.areaId, 'HDP titulo', names.hdp, (hdp) => {
    writeTitle(hdp, `Histórico de producción - ${plan.selection.areaId}`, plan.selection.areaName);
  });
  await writeAreaSheetStep(plan.selection.areaId, 'HDP parametros', names.hdp, (hdp) => {
    writeHdpParams(hdp, plan, monthly, warnings, middleMissingPolicy);
  });
  await writeAreaSheetStep(plan.selection.areaId, 'HDP tabla', names.hdp, (hdp) => {
    writeHdpTable(hdp, plan, monthly, middleMissingPolicy);
  });

  await writeAreaSheet(plan.selection.areaId, 'Detalle', names.detalle, async (detalle) => {
    writeDetailSheet(detalle, records);
  });
}

export async function writeAreaForecastSheets(
  plan: AreaWorkbookPlan,
  monthly: MonthlyAggregate[],
  middleMissingPolicy: 'blank' | 'zero',
  applyOverrideFields: AreaForecastOverrideField[] = [],
): Promise<void> {
  const names = areaSheetNames(plan.selection.areaId);
  const preserved = plan.mode === 'update' ? await readExistingSettings(names.prono, names.pozos) : undefined;
  await Excel.run(async (context) => {
    await ensureDebugSheet(context);
    if (plan.mode === 'regenerate') {
      await deleteIfExists(context, [names.prono, names.pozos, names.graficos]);
    }
    await context.sync();
  });
  await writeAreaSheet(plan.selection.areaId, 'Prono', names.prono, async (prono) => {
    writePronoSheet(prono, plan, monthly, middleMissingPolicy);
    if (preserved?.prono) restorePronoSettings(prono, preserved.prono);
    if (applyOverrideFields.length > 0) applyPlanSettings(prono, plan, applyOverrideFields);
  });
  await writeAreaSheet(plan.selection.areaId, 'Pozos', names.pozos, async (pozos) => {
    writePozosSheet(pozos, plan, monthly, middleMissingPolicy);
    if (preserved?.pozos) restorePozosSettings(pozos, preserved.pozos);
  });
  await writeAreaSheet(plan.selection.areaId, 'Graficos', names.graficos, async (graficos) => {
    writeChartsSheet(graficos, plan, monthly.length);
  });
}

async function readExistingSettings(pronoName: string, pozosName: string): Promise<PreservedSettings> {
  return Excel.run(async (context) => {
    const output: PreservedSettings = {};
    const prono = context.workbook.worksheets.getItemOrNullObject(pronoName);
    const pozos = context.workbook.worksheets.getItemOrNullObject(pozosName);
    prono.load('name');
    pozos.load('name');
    await context.sync();

    if (!prono.isNullObject) {
      const methods = prono.getRange('B4:B7');
      const decline = prono.getRange('E4:E10');
      const initials = prono.getRange('H4:H6');
      methods.load('values');
      decline.load('values');
      initials.load('values');
      await context.sync();
      output.prono = { methods: methods.values, decline: decline.values, initials: initials.values };
    }
    if (!pozos.isNullObject) {
      const switches = pozos.getRange('B4:B6');
      const decline = pozos.getRange('E4:E6');
      switches.load('values');
      decline.load('values');
      await context.sync();
      output.pozos = { switches: switches.values, decline: decline.values };
    }
    return output;
  });
}

function restorePronoSettings(sheet: Excel.Worksheet, settings: NonNullable<PreservedSettings['prono']>): void {
  sheet.getRange('B4:B7').values = settings.methods;
  sheet.getRange('E4:E10').values = settings.decline;
  sheet.getRange('H4:H6').values = settings.initials;
}

function restorePozosSettings(sheet: Excel.Worksheet, settings: NonNullable<PreservedSettings['pozos']>): void {
  sheet.getRange('B4:B6').values = settings.switches;
  sheet.getRange('E4:E6').values = settings.decline;
}

function applyPlanSettings(sheet: Excel.Worksheet, plan: AreaWorkbookPlan, fields: AreaForecastOverrideField[]): void {
  const override = plan.override;
  const values: Partial<Record<AreaForecastOverrideField, string | number>> = {
    grossMethod: override?.grossMethod ?? plan.defaults.grossMethod,
    oilMethod: override?.oilMethod ?? plan.defaults.oilMethod,
    gasMethod: override?.gasMethod ?? plan.defaults.gasMethod,
    takeInitialFromHistory: (override?.takeInitialFromHistory ?? plan.defaults.takeInitialFromHistory) ? 'Sí' : 'No',
    grossDi: override?.grossDi ?? DEFAULT_DECLINE.grossDi,
    grossB: override?.grossB ?? DEFAULT_DECLINE.grossB,
    oilDi: override?.oilDi ?? DEFAULT_DECLINE.oilDi,
    oilB: override?.oilB ?? DEFAULT_DECLINE.oilB,
    gasDi: override?.gasDi ?? DEFAULT_DECLINE.gasDi,
    gasB: override?.gasB ?? DEFAULT_DECLINE.gasB,
  };
  const cells: Partial<Record<AreaForecastOverrideField, string>> = {
    grossMethod: 'B4',
    oilMethod: 'B5',
    gasMethod: 'B6',
    takeInitialFromHistory: 'B7',
    grossDi: 'E4',
    grossB: 'E5',
    oilDi: 'E6',
    oilB: 'E7',
    gasDi: 'E8',
    gasB: 'E9',
  };
  for (const field of fields) {
    const cell = cells[field];
    const value = values[field];
    if (cell && value !== undefined) sheet.getRange(cell).values = [[value]];
  }
}

async function writeAreaSheet(
  areaId: string,
  sheetStep: string,
  sheetName: string,
  writer: (sheet: Excel.Worksheet, context: Excel.RequestContext) => Promise<void> | void,
): Promise<void> {
  await appendDebug(createDebugEntry(areaId, 'info', `Escribiendo hoja ${sheetStep}`));
  try {
    await prepareAreaSheet(areaId, sheetStep, sheetName);
    await writeAreaSheetStep(areaId, `${sheetStep} contenido`, sheetName, writer);
    await appendDebug(createDebugEntry(areaId, 'ok', `Hoja ${sheetStep} escrita`));
  } catch (error) {
    const detail = error instanceof Error ? error.message : String(error);
    await appendDebug(createDebugEntry(areaId, 'error', `Fallo hoja ${sheetStep}: ${detail}`));
    throw error;
  }
}

async function prepareAreaSheet(areaId: string, sheetStep: string, sheetName: string): Promise<void> {
  await writeAreaSheetStep(areaId, `${sheetStep} preparar`, sheetName, async (sheet, context) => {
    const charts = sheet.charts;
    charts.load('items');
    await context.sync();
    for (const chart of charts.items) chart.delete();
    sheet.getRange().clear();
    sheet.getRange('A1:H2').unmerge();
  });
}

async function writeAreaSheetStep(
  areaId: string,
  step: string,
  sheetName: string,
  writer: (sheet: Excel.Worksheet, context: Excel.RequestContext) => Promise<void> | void,
): Promise<void> {
  await appendDebug(createDebugEntry(areaId, 'info', `Paso hoja ${step}`));
  try {
    await Excel.run(async (context) => {
      const sheet = await getOrAddSheet(context, sheetName);
      await writer(sheet, context);
      await context.sync();
    });
    await appendDebug(createDebugEntry(areaId, 'ok', `Paso hoja ${step} ok`));
  } catch (error) {
    const detail = error instanceof Error ? error.message : String(error);
    await appendDebug(createDebugEntry(areaId, 'error', `Fallo paso hoja ${step}: ${detail}`));
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
  writeTitle(sheet, `Histórico de producción - ${plan.selection.areaId}`, plan.selection.areaName);
  writeHdpParams(sheet, plan, monthly, warnings, middleMissingPolicy);
  writeHdpTable(sheet, plan, monthly, middleMissingPolicy);
}

function writeHdpParams(
  sheet: Excel.Worksheet,
  plan: AreaWorkbookPlan,
  monthly: MonthlyAggregate[],
  warnings: string[],
  middleMissingPolicy: 'blank' | 'zero',
): void {
  writeMatrix(sheet, 'A4', [
    ['Provincia', plan.selection.province],
    ['Area', `${plan.selection.areaId} - ${plan.selection.areaName}`],
    ['Último mes publicado', monthly.at(-1)?.date ?? 'Sin datos'],
    ['Advertencias intermedias', warnings.length ? warnings.join(', ') : 'Sin advertencias'],
    ['Política de faltantes', warnings.length ? (middleMissingPolicy === 'blank' ? 'Vacío' : 'Cero') : 'No aplica'],
  ], `Parámetros HDP ${plan.selection.areaId}`);
}

function writeHdpTable(
  sheet: Excel.Worksheet,
  plan: AreaWorkbookPlan,
  monthly: MonthlyAggregate[],
  middleMissingPolicy: 'blank' | 'zero',
): void {
  const headers = ['Fecha', 'Petróleo', 'Gas', 'Agua', 'Bruta', 'Agua iny.', 'Pozos petróleo', 'Pozos gas', 'Inyectores', 'Faltante'];
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
  sheet.freezePanes.freezeRows(9);
}

function writePronoSheet(
  sheet: Excel.Worksheet,
  plan: AreaWorkbookPlan,
  monthly: MonthlyAggregate[],
  middleMissingPolicy: 'blank' | 'zero',
): void {
  writeTitle(sheet, `Pronóstico de producción - ${plan.selection.areaId}`, plan.selection.areaName);
  const override = plan.override;
  const gross = override?.grossMethod ?? plan.defaults.grossMethod;
  const oil = override?.oilMethod ?? plan.defaults.oilMethod;
  const gas = override?.gasMethod ?? plan.defaults.gasMethod;
  const takeInitial = override?.takeInitialFromHistory ?? plan.defaults.takeInitialFromHistory;

  writeMatrix(sheet, 'A4', [
    ['Método bruta', gross],
    ['Método petróleo', oil],
    ['Método gas', gas],
    ['Tomar inicial de historia', takeInitial ? 'Sí' : 'No'],
    ['Horizonte años', plan.defaults.horizonYears],
    ['Inicio pronóstico', nextMonth(monthly.at(-1)?.date)],
    ['Nota', 'Celdas de supuestos editables. Actualizar resumen desde el taskpane.'],
  ], `Parametros Prono ${plan.selection.areaId}`);
  sheet.getRange('B4').dataValidation.rule = { list: { inCellDropDown: true, source: 'Constante,HypMod,Declinación Hip.,Declinación Exp.' } };
  sheet.getRange('B5').dataValidation.rule = { list: { inCellDropDown: true, source: 'Constante,HypMod,Declinación Hip.,Declinación Exp.,Rap Np' } };
  sheet.getRange('B6').dataValidation.rule = { list: { inCellDropDown: true, source: 'Constante,HypMod,Declinación Hip.,Declinación Exp.,RGP' } };
  sheet.getRange('B7').dataValidation.rule = { list: { inCellDropDown: true, source: 'Sí,No' } };

  writeMatrix(sheet, 'D4', [
    ['Di bruta anual', override?.grossDi ?? DEFAULT_DECLINE.grossDi],
    ['b bruta', override?.grossB ?? DEFAULT_DECLINE.grossB],
    ['Di petróleo anual', override?.oilDi ?? DEFAULT_DECLINE.oilDi],
    ['b petróleo', override?.oilB ?? DEFAULT_DECLINE.oilB],
    ['Di gas anual', override?.gasDi ?? DEFAULT_DECLINE.gasDi],
    ['b gas', override?.gasB ?? DEFAULT_DECLINE.gasB],
    ['RGP constante', (lastNonMissing(monthly, 'gas') / Math.max(lastNonMissing(monthly, 'oil'), 1)) * 1000],
  ], `Supuestos Prono ${plan.selection.areaId}`);

  const lastOil = lastNonMissing(monthly, 'oil');
  const lastGas = lastNonMissing(monthly, 'gas');
  const lastGross = lastNonMissing(monthly, 'gross');
  writeMatrix(sheet, 'G4', [
    ['Inicial bruta manual', lastGross],
    ['Inicial petróleo manual', lastOil],
    ['Inicial gas manual', lastGas],
  ], `Iniciales Prono ${plan.selection.areaId}`);

  const headers = ['Fecha', 'Petróleo hist.', 'Petróleo prono', 'Gas hist.', 'Gas prono', 'Bruta hist.', 'Bruta prono', 'Agua hist.', 'Agua prono', 'Wcut', 'RGP', 'RAP'];
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
    const oilInitial = `IF($B$7="Sí",${lastOil},$H$5)`;
    const gasInitial = `IF($B$7="Sí",${lastGas},$H$6)`;
    const grossInitial = `IF($B$7="Sí",${lastGross},$H$4)`;
    for (let i = 1; i <= months; i++) {
      const date = new Date(start);
      date.setMonth(date.getMonth() + i);
      const tYears = i / 12;
      const rowNumber = 12 + rows.length + 1;
      rows.push([
        date.toISOString().slice(0, 10),
        '',
        forecastFormula('B5', oilInitial, 'E6', 'E7', tYears),
        '',
        forecastFormula('B6', gasInitial, 'E8', 'E9', tYears, 'E10', `C${rowNumber}`),
        '',
        forecastFormula('B4', grossInitial, 'E4', 'E5', tYears),
        '',
        `=MAX(G${rowNumber}-C${rowNumber},0)`,
        `=IF(G${rowNumber}>0,I${rowNumber}/G${rowNumber},"")`,
        `=IF(C${rowNumber}>0,E${rowNumber}/C${rowNumber}*1000,"")`,
        `=IF(C${rowNumber}>0,I${rowNumber}/C${rowNumber},"")`,
      ]);
    }
  }
  writeTable(sheet, 'A12:L12', headers, rows, `Prono ${plan.selection.areaId}`);
  sheet.freezePanes.freezeRows(12);
}

function writePozosSheet(
  sheet: Excel.Worksheet,
  plan: AreaWorkbookPlan,
  monthly: MonthlyAggregate[],
  middleMissingPolicy: 'blank' | 'zero',
): void {
  writeTitle(sheet, 'Pronóstico de pozos', 'Cantidad histórica y supuestos editables');
  writeMatrix(sheet, 'A4', [
    ['Tomar pozos petróleo de historia', 'Sí'],
    ['Tomar pozos gas de historia', 'Sí'],
    ['Tomar inyectores de historia', 'Sí'],
    ['Método', 'Declinación exponencial'],
  ], `Parametros Pozos ${plan.selection.areaId}`);
  sheet.getRange('B4:B6').dataValidation.rule = { list: { inCellDropDown: true, source: 'Sí,No' } };
  writeMatrix(sheet, 'D4', [
    ['Di pozos petróleo anual', 0.04],
    ['Di pozos gas anual', 0.04],
    ['Di inyectores anual', 0.03],
  ], `Supuestos Pozos ${plan.selection.areaId}`);

  const headers = ['Fecha', 'Pozos petróleo', 'Pozos gas', 'Inyectores'];
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
        `=IF($B$4="Sí",${lastOilWells}*EXP(-$E$4*${t}),${lastOilWells})`,
        `=IF($B$5="Sí",${lastGasWells}*EXP(-$E$5*${t}),${lastGasWells})`,
        `=IF($B$6="Sí",${lastInjectors}*EXP(-$E$6*${t}),${lastInjectors})`,
      ]);
    }
  }
  writeTable(sheet, 'A9:D9', headers, rows, 'Prono Pozos');
  sheet.freezePanes.freezeRows(9);
}

function writeDetailSheet(sheet: Excel.Worksheet, records: ProductionRecord[]): void {
  writeTitle(sheet, 'Detalle Capítulo IV', 'Pozo-mes descargado');
  const headers = ['Año', 'Mes', 'Área', 'Pozo', 'Id pozo', 'Petróleo', 'Gas', 'Agua', 'Agua iny.'];
  const rows = records.map((r) => [r.year, r.month, r.areaId, r.wellName, r.wellId, r.oil, r.gas, r.water, r.waterInjection]);
  writeTable(sheet, 'A4:I4', headers, rows, 'Detalle Capitulo IV');
  sheet.freezePanes.freezeRows(4);
}

function writeChartsSheet(
  sheet: Excel.Worksheet,
  plan: AreaWorkbookPlan,
  historyMonths: number,
): void {
  const names = areaSheetNames(plan.selection.areaId);
  const forecastMonths = plan.defaults.horizonYears * 12;
  const totalMonths = Math.max(1, historyMonths + forecastMonths);
  writeTitle(sheet, `Gráficos técnicos - ${plan.selection.areaId}`, 'Histórico y pronóstico separados por variable');
  writeForecastCharts(sheet, { prono: names.prono, pozos: names.pozos, hdp: names.hdp }, totalMonths, historyMonths);
}

function maybeBlank(month: MonthlyAggregate, value: number, middleMissingPolicy: 'blank' | 'zero'): string | number {
  if (month.missingKind === 'middle' && middleMissingPolicy === 'blank') return '';
  return value;
}
