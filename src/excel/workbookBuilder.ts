import { brand } from '../branding/tokens';
import type {
  AreaWorkbookPlan,
  BuildProgressHandler,
  ForecastDefaults,
  ForecastMethod,
  MissingMonthsDecisionHandler,
  MonthlyAggregate,
  ProductionRecord,
} from '../models/types';
import { aggregateMonthly, countProductionResources, fetchAreaProduction } from '../services/capiv';
import { appendDebug, createDebugEntry, ensureDebugSheet } from './debugSheet';
import { areaSheetNames, SUMMARY_SHEET, STATE_SHEET } from './names';

interface BuildResult {
  areaId: string;
  areaName: string;
  monthly: MonthlyAggregate[];
  summary: SummaryMonthly[];
  warnings: string[];
  middleMissingPolicy: 'blank' | 'zero';
}

interface SummaryMonthly {
  date: string;
  oil: number;
  gas: number;
  water: number;
  gross: number;
  kind: 'hist' | 'prono';
}

const DOWNLOAD_SHEET_NAME = 'CapIV_Descarga';

export async function buildWorkbook(
  plans: AreaWorkbookPlan[],
  onProgress?: BuildProgressHandler,
  onMissingMonths?: MissingMonthsDecisionHandler,
): Promise<void> {
  const total = estimateWorkUnits(plans);
  let completed = 0;
  const report = (message: string, plan?: AreaWorkbookPlan, areaIndex?: number, increment = 0) => {
    completed = Math.min(total, completed + increment);
    onProgress?.({
      areaId: plan?.selection.areaId,
      areaIndex,
      areaTotal: plans.length,
      completed,
      total,
      percent: total ? Math.round((completed / total) * 100) : 0,
      message,
    });
  };

  report('Iniciando workbook');
  await appendDebug(createDebugEntry('Inicio', 'info', `Areas seleccionadas: ${plans.map((p) => p.selection.areaId).join(', ')}`));

  const results: BuildResult[] = [];
  for (const [index, plan] of plans.entries()) {
    const areaIndex = index + 1;
    const startYear = plan.override?.startYear ?? plan.selection.startYearOverride ?? plan.defaults.startYear;
    try {
      report(`Area ${areaIndex}/${plans.length}: ${plan.selection.areaId} desde ${startYear}`, plan, areaIndex);
      await appendDebug(createDebugEntry(plan.selection.areaId, 'info', `Descarga desde ${startYear}`));
      await ensureDownloadSheet(plan);
      const records = await fetchAreaProduction(
        plan.selection,
        startYear,
        async (message) => {
          report(message, plan, areaIndex, message.startsWith('Descargado ') ? 1 : 0);
          await appendDebug(createDebugEntry(plan.selection.areaId, 'info', message));
        },
        async (resource) => {
          await appendDownloadRows(plan, resource.kind, resource.year, resource.totalRows, resource.matchedRows);
        },
      );
      report(`Filtrando registros ${plan.selection.areaId}: ${records.length} pozo-mes`, plan, areaIndex, 1);
      await appendDebug(createDebugEntry(plan.selection.areaId, 'ok', `Registros pozo-mes descargados: ${records.length}`));

      const monthly = aggregateMonthly(records, startYear);
      report(`Armando serie mensual ${plan.selection.areaId}: ${monthly.length} meses`, plan, areaIndex, 1);
      await appendDebug(createDebugEntry(plan.selection.areaId, 'ok', `Serie mensual armada: ${monthly.length} meses`));

      const middleWarnings = monthly
        .filter((m) => m.missingKind === 'middle')
        .map((m) => `${m.year}-${String(m.month).padStart(2, '0')}`);
      const leadingMissing = monthly.filter((m) => m.missingKind === 'leading').length;
      let middleMissingPolicy: 'blank' | 'zero' = 'blank';
      if (middleWarnings.length > 0) {
        await appendDebug(createDebugEntry(plan.selection.areaId, 'warning', `Se pide decision por faltantes intermedios: ${middleWarnings.join(', ')}`));
        report(`Esperando decision por faltantes ${plan.selection.areaId}`, plan, areaIndex);
        middleMissingPolicy = onMissingMonths
          ? await onMissingMonths({
              areaId: plan.selection.areaId,
              areaName: plan.selection.areaName,
              months: middleWarnings,
            })
          : 'blank';
        await appendDebug(createDebugEntry(plan.selection.areaId, 'info', `Decision faltantes intermedios: ${middleMissingPolicy}`));
      }
      const warnings = middleWarnings;
      if (leadingMissing > 0) {
        await appendDebug(createDebugEntry(plan.selection.areaId, 'info', `${leadingMissing} meses iniciales sin dato completados con 0`));
      }
      if (warnings.length > 0) {
        await appendDebug(createDebugEntry(plan.selection.areaId, 'warning', `Meses intermedios faltantes: ${warnings.join(', ')}. Politica: ${middleMissingPolicy}`));
      }
      const summary = buildAreaSummary(plan, monthly, middleMissingPolicy);
      report(`Preparando hojas ${plan.selection.areaId}`, plan, areaIndex, 1);
      await appendDebug(createDebugEntry(plan.selection.areaId, 'info', `Escritura de hojas iniciada`));
      await writeAreaSheets(plan, records, monthly, warnings, middleMissingPolicy);
      report(`Hojas escritas ${plan.selection.areaId}`, plan, areaIndex, 1);
      await appendDebug(createDebugEntry(plan.selection.areaId, 'ok', `Hojas escritas. Filas resumen area: ${summary.length}`));
      results.push({
        areaId: plan.selection.areaId,
        areaName: plan.selection.areaName,
        monthly,
        summary,
        warnings,
        middleMissingPolicy,
      });
    } catch (error) {
      const detail = error instanceof Error ? error.stack ?? error.message : String(error);
      await appendDebug(createDebugEntry(plan.selection.areaId, 'error', detail));
      throw error;
    }
  }

  report(`Escribiendo resumen consolidado`, undefined, undefined, 1);
  await appendDebug(createDebugEntry('Resumen_Areas', 'info', `Escritura resumen consolidado para ${results.length} areas`));
  await writeSummary(results);
  await appendDebug(createDebugEntry('Resumen_Areas', 'ok', 'Resumen consolidado escrito'));
  report(`Guardando estado del workbook`, undefined, undefined, 1);
  await appendDebug(createDebugEntry(STATE_SHEET, 'info', 'Guardando estado oculto'));
  await writeState(plans, results);
  await appendDebug(createDebugEntry(STATE_SHEET, 'ok', 'Estado guardado'));
  report(`Workbook actualizado`, undefined, undefined, total - completed);
  await appendDebug(createDebugEntry('Fin', 'ok', 'Workbook actualizado'));
}

async function ensureDownloadSheet(plan: AreaWorkbookPlan): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = await getOrAddSheet(context, DOWNLOAD_SHEET_NAME);
    const used = sheet.getUsedRangeOrNullObject();
    used.load('rowCount');
    await context.sync();
    if (used.isNullObject) {
      writeTitle(sheet, 'Descarga Capitulo IV', 'Datos pegados a medida que termina cada recurso');
      const headers = ['Timestamp', 'Area', 'Nombre', 'Fuente', 'Anio recurso', 'Anio', 'Mes', 'Pozo', 'Id pozo', 'Petroleo', 'Gas', 'Agua', 'Agua iny.'];
      const header = sheet.getRange('A4:M4');
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

async function appendDownloadRows(
  plan: AreaWorkbookPlan,
  kind: 'conv' | 'nc',
  year: number | undefined,
  totalRows: number,
  records: ProductionRecord[],
): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = await getOrAddSheet(context, DOWNLOAD_SHEET_NAME);
    const used = sheet.getUsedRangeOrNullObject();
    used.load('rowCount');
    await context.sync();

    const nextRow = used.isNullObject ? 4 : used.rowCount;
    const timestamp = new Date().toISOString();
    const rows: (string | number)[][] = records.length
      ? records.map((record) => [
          timestamp,
          plan.selection.areaId,
          plan.selection.areaName,
          kind === 'conv' ? 'Convencional' : 'No convencional',
          year ?? '',
          record.year,
          record.month,
          record.wellName,
          record.wellId,
          record.oil,
          record.gas,
          record.water,
          record.waterInjection,
        ])
      : [[timestamp, plan.selection.areaId, plan.selection.areaName, kind === 'conv' ? 'Convencional' : 'No convencional', year ?? '', '', '', '', '', 0, 0, 0, 0]];

    const range = sheet.getRangeByIndexes(nextRow, 0, rows.length, 13);
    range.values = rows;
    sheet.getRange('A2').values = [[`Ultimo recurso: ${plan.selection.areaId} ${kind === 'conv' ? year : 'NC'} - ${records.length} coincidencias de ${totalRows} filas`]];
    await context.sync();
  });
}

function estimateWorkUnits(plans: AreaWorkbookPlan[]): number {
  const perAreaFixedSteps = 4;
  const globalSteps = 2;
  return plans.reduce((total, plan) => {
    const startYear = plan.override?.startYear ?? plan.selection.startYearOverride ?? plan.defaults.startYear;
    return total + countProductionResources(startYear) + perAreaFixedSteps;
  }, globalSteps);
}

async function writeAreaSheets(
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
  writeTable(sheet, 'A9:J9', headers, rows);
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
    const lastWater = lastNonMissing(monthly, 'water');
    for (let i = 1; i <= months; i++) {
      const date = new Date(start);
      date.setMonth(date.getMonth() + i);
      const tYears = i / 12;
      rows.push([
        date.toISOString().slice(0, 10),
        '',
        forecastFormula('B5', lastOil, 'E6', 'E7', tYears),
        '',
        forecastFormula('B6', lastGas, 'E8', 'E9', tYears, 'E10', lastOil),
        '',
        forecastFormula('B4', lastGross, 'E4', 'E5', tYears),
        '',
        `=MAX(G${12 + rows.length + 1}-C${12 + rows.length + 1},0)`,
        `=IF(G${12 + rows.length + 1}>0,I${12 + rows.length + 1}/G${12 + rows.length + 1},"")`,
        `=IF(C${12 + rows.length + 1}>0,E${12 + rows.length + 1}/C${12 + rows.length + 1}*1000,"")`,
        `=IF(C${12 + rows.length + 1}>0,I${12 + rows.length + 1}/C${12 + rows.length + 1},"")`,
      ]);
    }
  }
  writeTable(sheet, 'A12:L12', headers, rows);
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
  writeTable(sheet, 'A9:D9', headers, rows);
}

function writeDetailSheet(sheet: Excel.Worksheet, records: ProductionRecord[]): void {
  sheet.getRange().clear();
  writeTitle(sheet, 'Detalle Capitulo IV', 'Pozo-mes descargado');
  const headers = ['Anio', 'Mes', 'Area', 'Pozo', 'Id pozo', 'Petroleo', 'Gas', 'Agua', 'Agua iny.'];
  const rows = records.map((r) => [r.year, r.month, r.areaId, r.wellName, r.wellId, r.oil, r.gas, r.water, r.waterInjection]);
  writeTable(sheet, 'A4:I4', headers, rows);
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

async function writeSummary(results: BuildResult[]): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = await getOrAddSheet(context, SUMMARY_SHEET);
    sheet.getRange().clear();
    writeTitle(sheet, 'Resumen consolidado de areas', 'Suma de historicos y proyecciones individuales');
    const headers = ['Area', 'Nombre', 'Ultimo mes', 'Petroleo ultimo', 'Gas ultimo', 'Bruta ultimo', 'Warnings'];
    const rows = results.map((result) => {
      const last = result.monthly.at(-1);
      return [
        result.areaId,
        result.areaName,
        last?.date ?? 'Sin datos',
        last?.oil ?? 0,
        last?.gas ?? 0,
        last?.gross ?? 0,
        result.warnings.join(', '),
      ];
    });
    writeTable(sheet, 'A4:G4', headers, rows);
    writeConsolidatedMonthly(sheet, results);
    await context.sync();
  });
}

async function writeState(plans: AreaWorkbookPlan[], results: BuildResult[]): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = await getOrAddSheet(context, STATE_SHEET);
    sheet.visibility = Excel.SheetVisibility.veryHidden;
    sheet.getRange().clear();
    sheet.getRange('A1').values = [[JSON.stringify({ schema: 1, plans, results, savedAt: new Date().toISOString() }, null, 2)]];
    await context.sync();
  });
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

function writeTable(sheet: Excel.Worksheet, headerAddress: string, headers: string[], rows: (string | number)[][]): void {
  validateRows(sheet.name, headerAddress, headers, rows);
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

function validateRows(sheetName: string, headerAddress: string, headers: string[], rows: (string | number)[][]): void {
  for (let index = 0; index < rows.length; index++) {
    if (rows[index].length !== headers.length) {
      throw new Error(
        `${sheetName} ${headerAddress}: fila ${index + 1} tiene ${rows[index].length} columnas; se esperaban ${headers.length}`,
      );
    }
  }
}

function writeConsolidatedMonthly(sheet: Excel.Worksheet, results: BuildResult[]): void {
  const byDate = new Map<string, { oil: number; gas: number; water: number; gross: number; kind: 'hist' | 'prono' }>();
  for (const result of results) {
    for (const month of result.summary) {
      const current = byDate.get(month.date) ?? { oil: 0, gas: 0, water: 0, gross: 0, kind: month.kind };
      current.oil += month.oil;
      current.gas += month.gas;
      current.water += month.water;
      current.gross += month.gross;
      if (month.kind === 'prono') current.kind = 'prono';
      byDate.set(month.date, current);
    }
  }
  const rows = [...byDate.entries()]
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([date, values]) => [date, values.kind, values.oil, values.gas, values.water, values.gross]);
  writeTable(sheet, 'A14:F14', ['Fecha', 'Tipo', 'Petroleo total', 'Gas total', 'Agua total', 'Bruta total'], rows);
  const rowCount = Math.max(2, rows.length + 1);
  addLineChart(sheet, sheet.getRangeByIndexes(13, 0, rowCount, 6), 'Consolidado total', 'I4', 'P22');
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

function lastNonMissing(
  monthly: MonthlyAggregate[],
  key: 'oil' | 'gas' | 'water' | 'gross' | 'oilWells' | 'gasWells' | 'injectorWells',
): number {
  for (let i = monthly.length - 1; i >= 0; i--) {
    const month = monthly[i];
    if (!month.missing && month[key] > 0) return month[key];
  }
  return 0;
}

function buildAreaSummary(
  plan: AreaWorkbookPlan,
  monthly: MonthlyAggregate[],
  middleMissingPolicy: 'blank' | 'zero',
): SummaryMonthly[] {
  const rows: SummaryMonthly[] = monthly
    .filter((m) => !(m.missingKind === 'middle' && middleMissingPolicy === 'blank'))
    .map((m) => ({
      date: m.date,
      oil: m.oil,
      gas: m.gas,
      water: m.water,
      gross: m.gross,
      kind: 'hist',
    }));
  const last = monthly.at(-1);
  if (!last) return rows;

  const methods = resolveMethods(plan);
  const start = new Date(`${last.date}T00:00:00`);
  const months = plan.defaults.horizonYears * 12;
  const lastOil = lastNonMissing(monthly, 'oil');
  const lastGas = lastNonMissing(monthly, 'gas');
  const lastGross = lastNonMissing(monthly, 'gross');
  const rgp = (lastGas / Math.max(lastOil, 1)) * 1000;
  for (let i = 1; i <= months; i++) {
    const date = new Date(start);
    date.setMonth(date.getMonth() + i);
    const t = i / 12;
    const oil = evaluateForecast(methods.oilMethod, lastOil, 0.12, 0.7, t);
    const gas = methods.gasMethod === 'RGP' ? (rgp * oil) / 1000 : evaluateForecast(methods.gasMethod, lastGas, 0.12, 0.7, t);
    const gross = evaluateForecast(methods.grossMethod, lastGross, 0.12, 0.7, t);
    rows.push({
      date: date.toISOString().slice(0, 10),
      oil,
      gas,
      water: Math.max(gross - oil, 0),
      gross,
      kind: 'prono',
    });
  }
  return rows;
}

function resolveMethods(plan: AreaWorkbookPlan): Pick<ForecastDefaults, 'grossMethod' | 'oilMethod' | 'gasMethod'> {
  return {
    grossMethod: plan.override?.grossMethod ?? plan.defaults.grossMethod,
    oilMethod: plan.override?.oilMethod ?? plan.defaults.oilMethod,
    gasMethod: plan.override?.gasMethod ?? plan.defaults.gasMethod,
  };
}

function evaluateForecast(method: ForecastMethod, initial: number, di: number, b: number, tYears: number): number {
  if (!Number.isFinite(initial) || initial <= 0) return 0;
  if (method === 'Constante') return initial;
  if (method === 'Declinación Exp.') return initial * Math.exp(-di * tYears);
  if (method === 'Declinación Hip.' || method === 'HypMod' || method === 'Rap Np') {
    return initial / Math.pow(1 + b * di * tYears, 1 / b);
  }
  return initial;
}

function forecastFormula(
  methodCell: string,
  initialValue: number,
  diCell: string,
  bCell: string,
  tYears: number,
  rgpCell?: string,
  oilInitial?: number,
): string {
  const initial = Number.isFinite(initialValue) ? initialValue : 0;
  const oil = Number.isFinite(oilInitial) ? oilInitial : initial;
  const t = Number(tYears.toFixed(6));
  return `=IF(${methodCell}="Constante",${initial},IF(${methodCell}="Declinación Exp.",${initial}*EXP(-${diCell}*${t}),IF(${methodCell}="Declinación Hip.",${initial}/POWER(1+${bCell}*${diCell}*${t},1/${bCell}),IF(${methodCell}="HypMod",${initial}/POWER(1+${bCell}*${diCell}*${t},1/${bCell}),IF(${methodCell}="RGP",${rgpCell ?? 0}*${oil}/1000,IF(${methodCell}="Rap Np",${initial}/POWER(1+${bCell}*${diCell}*${t},1/${bCell}),${initial}))))))`;
}

function nextMonth(date?: string): string {
  if (!date) return '';
  const parsed = new Date(`${date}T00:00:00`);
  parsed.setMonth(parsed.getMonth() + 1);
  return parsed.toISOString().slice(0, 10);
}
