import type {
  AreaForecastOverrideField,
  AreaWorkbookPlan,
  AssetGroup,
  BuildProgressHandler,
  DataOutputTarget,
  MissingMonthsDecisionHandler,
  MonthlyAggregate,
  WorkbookAreaData,
  OverwriteDecisionHandler,
} from '../models/types';
import { normalizeAssetGroups } from '../domain/assets';
import { evaluateForecast, lastNonMissing, resolveForecastMethods } from '../domain/forecast';
import { aggregateMonthly, countProductionResources, fetchAreaProduction } from '../services/capiv';
import { appendDebug, createDebugEntry } from './debugSheet';
import { writeAreaForecastSheets } from './areaSheets';
import { writeDatabaseTable } from './databaseSheet';
import { appendDownloadLedgerEvent, ensureDownloadLedger } from './downloadLedger';
import { areaSheetNames, assetSheetName, SUMMARY_SHEET, STATE_SHEET } from './names';
import { addLineChart, getOrAddSheet, writeMatrix, writeTable, writeTitle } from './sheetLayout';
import { readSavedWorkbookPlans } from './workbookState';

interface BuildResult extends WorkbookAreaData {
  summary: SummaryMonthly[];
}

interface SummaryMonthly {
  date: string;
  oil: number;
  gas: number;
  water: number;
  gross: number;
  kind: 'hist' | 'prono';
}

export interface BuildForecastOptions {
  assetGroups?: AssetGroup[];
  applyOverrideFields?: Record<string, AreaForecastOverrideField[]>;
}

export async function downloadWorkbookData(
  plans: AreaWorkbookPlan[],
  output: DataOutputTarget,
  onProgress?: BuildProgressHandler,
  onMissingMonths?: MissingMonthsDecisionHandler,
  onOverwrite?: OverwriteDecisionHandler,
): Promise<void> {
  const saved = await readSavedWorkbookPlans();
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

  const results: Array<{ plan: AreaWorkbookPlan; data: WorkbookAreaData; records: import('../models/types').ProductionRecord[] }> = [];
  for (const [index, plan] of plans.entries()) {
    const areaIndex = index + 1;
    const startYear = plan.override?.startYear ?? plan.selection.startYearOverride ?? plan.defaults.startYear;
    try {
      report(`Area ${areaIndex}/${plans.length}: ${plan.selection.areaId} desde ${startYear}`, plan, areaIndex);
      await appendDebug(createDebugEntry(plan.selection.areaId, 'info', `Descarga desde ${startYear}`));
      await ensureDownloadLedger(plan);
      const records = await fetchAreaProduction(
        plan.selection,
        startYear,
        async (message) => {
          report(message, plan, areaIndex, message.startsWith('Descargado ') ? 1 : 0);
          await appendDebug(createDebugEntry(plan.selection.areaId, 'info', message));
        },
        async (event) => {
          await appendDownloadLedgerEvent(plan, event);
        },
      );
      if (records.length === 0) {
        throw new Error(`No se encontraron registros para ${plan.selection.areaId} desde ${startYear}.`);
      }
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
      report(`Preparando base ${plan.selection.areaId}`, plan, areaIndex, 1);
      const data = {
        areaId: plan.selection.areaId,
        areaName: plan.selection.areaName,
        monthly,
        warnings,
        middleMissingPolicy,
      } satisfies WorkbookAreaData;
      results.push({ plan, data, records });
    } catch (error) {
      const detail = error instanceof Error ? error.stack ?? error.message : String(error);
      await appendDebug(createDebugEntry(plan.selection.areaId, 'error', detail));
      throw error;
    }
  }

  report('Escribiendo tabla en la celda elegida', undefined, undefined, 1);
  const written = await writeDatabaseTable(output, results, onOverwrite);
  await appendDebug(createDebugEntry('Base de datos', 'ok', `${written.rowCount} filas escritas en ${written.rangeAddress}`));
  report('Guardando datos del workbook', undefined, undefined, 1);
  await appendDebug(createDebugEntry(STATE_SHEET, 'info', 'Guardando estado oculto'));
  const mergedPlans = mergePlans(saved?.plans ?? [], plans);
  const mergedData = mergeData(saved?.data ?? [], results.map((result) => result.data));
  await writeState(mergedPlans, mergedData, {
    dataSavedAt: new Date().toISOString(),
    forecastSavedAt: saved?.forecastSavedAt,
  }, output, saved?.assetGroups);
  await appendDebug(createDebugEntry(STATE_SHEET, 'ok', 'Estado guardado'));
  report('Datos actualizados', undefined, undefined, total - completed);
  await appendDebug(createDebugEntry('Fin datos', 'ok', 'Descarga de datos completada'));
}

export async function buildForecastWorkbook(
  plans: AreaWorkbookPlan[],
  onProgress?: BuildProgressHandler,
  options: BuildForecastOptions = {},
): Promise<void> {
  const saved = await readSavedWorkbookPlans();
  if (!saved || saved.data.length === 0) {
    throw new Error('No hay datos descargados en este libro. Completá primero el flujo Datos.');
  }

  const total = plans.length * 3 + 2;
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

  report('Leyendo datos del libro');
  await appendDebug(createDebugEntry('Inicio pronósticos', 'info', `Áreas: ${plans.map((plan) => plan.selection.areaId).join(', ')}`));
  const results: BuildResult[] = [];
  for (const [index, plan] of plans.entries()) {
    const areaIndex = index + 1;
    const data = saved.data.find((item) => item.areaId === plan.selection.areaId);
    if (!data) throw new Error(`No hay datos descargados para ${plan.selection.areaId}.`);

    report(`Preparando pronóstico ${plan.selection.areaId}`, plan, areaIndex, 1);
    const summary = buildAreaSummary(plan, data.monthly, data.middleMissingPolicy);
    await writeAreaForecastSheets(plan, data.monthly, data.middleMissingPolicy, options.applyOverrideFields?.[plan.selection.areaId]);
    report(`Pronóstico escrito ${plan.selection.areaId}`, plan, areaIndex, 1);
    results.push({ ...data, summary });
    await appendDebug(createDebugEntry(plan.selection.areaId, 'ok', `Pronóstico generado: ${summary.length} meses`));
    report(`Área ${plan.selection.areaId} completada`, plan, areaIndex, 1);
  }

  report('Escribiendo resumen consolidado', undefined, undefined, 1);
  await loggedStep('Resumen consolidado', () => writeSummary(results));
  const assetGroups = normalizeAssetGroups(options.assetGroups ?? saved.assetGroups, results.map((result) => result.areaId));
  await loggedStep('Resumen de activos', () => writeAssetSummaries(results, saved.assetGroups, assetGroups));
  const forecastSavedAt = new Date().toISOString();
  await loggedStep('Estado del libro', () => writeState(mergePlans(saved.plans, plans), saved.data, { dataSavedAt: saved.dataSavedAt, forecastSavedAt }, saved.dataOutput, assetGroups));
  // Terminar mostrando el resumen consolidado (las hojas de activos se activan
  // al escribirse). Cosmético: si falla no aborta la generación.
  await Excel.run(async (context) => {
    context.workbook.worksheets.getItem(SUMMARY_SHEET).activate();
    await context.sync();
  }).catch(() => undefined);
  report('Pronósticos actualizados', undefined, undefined, total - completed);
  await appendDebug(createDebugEntry('Fin pronósticos', 'ok', 'Pronósticos y resumen actualizados'));
}

export async function buildWorkbook(
  plans: AreaWorkbookPlan[],
  output: DataOutputTarget,
  onProgress?: BuildProgressHandler,
  onMissingMonths?: MissingMonthsDecisionHandler,
): Promise<void> {
  await downloadWorkbookData(plans, output, onProgress, onMissingMonths);
  await buildForecastWorkbook(plans, onProgress);
}

function estimateWorkUnits(plans: AreaWorkbookPlan[]): number {
  const perAreaFixedSteps = 4;
  const globalSteps = 2;
  return plans.reduce((total, plan) => {
    const startYear = plan.override?.startYear ?? plan.selection.startYearOverride ?? plan.defaults.startYear;
    return total + countProductionResources(startYear) + perAreaFixedSteps;
  }, globalSteps);
}

async function loggedStep(step: string, action: () => Promise<void>): Promise<void> {
  await appendDebug(createDebugEntry(step, 'info', `Escribiendo ${step}`));
  try {
    await action();
    await appendDebug(createDebugEntry(step, 'ok', `${step} listo`));
  } catch (error) {
    const detail = error instanceof Error ? error.message : String(error);
    await appendDebug(createDebugEntry(step, 'error', `Fallo ${step}: ${detail}`));
    throw error;
  }
}

function mergePlans(existing: AreaWorkbookPlan[], replacements: AreaWorkbookPlan[]): AreaWorkbookPlan[] {
  const byArea = new Map(existing.map((plan) => [plan.selection.areaId, plan]));
  for (const plan of replacements) byArea.set(plan.selection.areaId, plan);
  return [...byArea.values()];
}

function mergeData(existing: WorkbookAreaData[], replacements: WorkbookAreaData[]): WorkbookAreaData[] {
  const byArea = new Map(existing.map((data) => [data.areaId, data]));
  for (const data of replacements) byArea.set(data.areaId, data);
  return [...byArea.values()];
}

async function writeSummary(results: BuildResult[]): Promise<void> {
  await writeSummarySheet(
    SUMMARY_SHEET,
    'Resumen consolidado de áreas',
    'Suma de históricos y proyecciones individuales',
    results,
    'Resumen Areas',
  );
}

async function writeAssetSummaries(
  results: BuildResult[],
  previousGroups: AssetGroup[],
  groups: AssetGroup[],
): Promise<void> {
  const activeNames = new Set(groups.map((group) => assetSheetName(group.name)));
  await Excel.run(async (context) => {
    for (const group of previousGroups) {
      const name = assetSheetName(group.name);
      if (activeNames.has(name)) continue;
      const sheet = context.workbook.worksheets.getItemOrNullObject(name);
      sheet.load('isNullObject');
      await context.sync();
      if (!sheet.isNullObject) sheet.delete();
    }
    await context.sync();
  });
  for (const group of groups) {
    const included = results.filter((result) => group.areaIds.includes(result.areaId));
    if (included.length === 0) continue;
    try {
      await writeSummarySheet(
        assetSheetName(group.name),
        `Resumen de activo - ${group.name}`,
        `Suma de ${included.length} ${included.length === 1 ? 'concesión' : 'concesiones'} pronosticadas por separado`,
        included,
        `Activo ${group.name}`,
        group.name,
      );
    } catch (error) {
      const detail = error instanceof Error ? error.message : String(error);
      throw new Error(`Activo ${group.name}: ${detail}`);
    }
  }
}

async function writeSummarySheet(
  sheetName: string,
  title: string,
  subtitle: string,
  results: BuildResult[],
  tableLabel: string,
  chartPrefix?: string,
): Promise<void> {
  const softWarnings: string[] = [];
  await Excel.run(async (context) => {
    // Syncs intermedios: atribuyen un fallo a la etapa exacta en la hoja Debug
    // y evitan acumular todo el resumen (tablas + fórmulas + gráficos) en un
    // solo batch, que con muchas áreas puede tirar un error interno de Excel.
    const checkpoint = async (label: string) => {
      try {
        await context.sync();
      } catch (error) {
        const detail = error instanceof Error ? error.message : String(error);
        throw new Error(`${label}: ${detail}`);
      }
    };
    // Variante para etapas cosméticas: un fallo queda como warning en Debug
    // sin abortar la generación.
    const softCheckpoint = async (label: string) => {
      try {
        await context.sync();
      } catch (error) {
        const detail = error instanceof Error ? error.message : String(error);
        softWarnings.push(`${label}: ${detail}`);
      }
    };
    const sheet = await getOrAddSheet(context, sheetName);
    const charts = sheet.charts;
    charts.load('items');
    await context.sync();
    for (const chart of charts.items) chart.delete();
    sheet.getRange().clear();
    await checkpoint('Preparación de hoja');
    writeTitle(sheet, title, subtitle);
    const headers = ['Área', 'Nombre', 'Último mes', 'Petróleo último', 'Gas último', 'Bruta última', 'Advertencias'];
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
    writeTable(sheet, 'A4:G4', headers, rows, tableLabel);
    await checkpoint('Tabla de áreas');
    await writeConsolidatedMonthly(sheet, results, checkpoint, softCheckpoint, chartPrefix);
    // freezeRows tira GeneralException si la hoja no está activa (bug conocido
    // de Office.js): las hojas de área se recrean y quedan activas, pero esta
    // persiste entre corridas, así que hay que activarla antes.
    sheet.activate();
    sheet.freezePanes.freezeRows(14);
    await softCheckpoint('Inmovilizar paneles');
  });
  for (const warning of softWarnings) {
    await appendDebug(createDebugEntry(sheetName, 'warning', warning));
  }
}

async function writeState(
  plans: AreaWorkbookPlan[],
  results: WorkbookAreaData[],
  timestamps: { dataSavedAt?: string; forecastSavedAt?: string },
  dataOutput?: DataOutputTarget,
  assetGroups: AssetGroup[] = [],
): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = await getOrAddSheet(context, STATE_SHEET);
    sheet.visibility = Excel.SheetVisibility.veryHidden;
    sheet.getRange().clear();
    writeWorkbookState(sheet, { schema: 4, plans, results, savedAt: new Date().toISOString(), ...timestamps, dataOutput, assetGroups });
    await context.sync();
  });
}

function writeWorkbookState(sheet: Excel.Worksheet, state: unknown): void {
  const serialized = JSON.stringify(state, null, 2);
  const chunkSize = 30000;
  const rows: string[][] = [];
  for (let index = 0; index < serialized.length; index += chunkSize) {
    rows.push([serialized.slice(index, index + chunkSize)]);
  }
  writeMatrix(sheet, 'A1', rows.length ? rows : [['']], 'Estado workbook');
}

async function writeConsolidatedMonthly(
  sheet: Excel.Worksheet,
  results: BuildResult[],
  checkpoint: (label: string) => Promise<void>,
  softCheckpoint: (label: string) => Promise<void>,
  chartPrefix?: string,
): Promise<void> {
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
    .map(([date, values], index) => {
      const summaryRow = 15 + index;
      return [
        date,
        values.kind,
        consolidatedFormula(results, summaryRow, [1, 2]),
        consolidatedFormula(results, summaryRow, [3, 4]),
        consolidatedFormula(results, summaryRow, [7, 8]),
        consolidatedFormula(results, summaryRow, [5, 6]),
      ];
    });
  writeTable(sheet, 'A14:F14', ['Fecha', 'Tipo', 'Petróleo total', 'Gas total', 'Agua total', 'Bruta total'], rows, 'Resumen mensual');
  await checkpoint('Tabla mensual consolidada');
  const oilSource = writeSummaryChartSource(sheet, 25, rows.length, 'C', 'Petróleo total');
  const gasSource = writeSummaryChartSource(sheet, 28, rows.length, 'D', 'Gas total');
  const waterSource = writeSummaryChartSource(sheet, 31, rows.length, 'E', 'Agua total');
  const grossSource = writeSummaryChartSource(sheet, 34, rows.length, 'F', 'Bruta total');
  await checkpoint('Fuentes de gráficos');
  const prefix = chartPrefix ? `${chartPrefix} - ` : '';
  addLineChart(sheet, oilSource, `${prefix}Petróleo`, 'I4', 'P22', 'm³/mes');
  await checkpoint('Gráfico Petróleo');
  addLineChart(sheet, gasSource, `${prefix}Gas`, 'Q4', 'X22', 'miles de m³/mes', '#1B4B6C');
  await checkpoint('Gráfico Gas');
  addLineChart(sheet, waterSource, `${prefix}Agua`, 'I24', 'P42', 'm³/mes', '#1B4B6C');
  await checkpoint('Gráfico Agua');
  addLineChart(sheet, grossSource, `${prefix}Producción bruta`, 'Q24', 'X42', 'm³/mes');
  await checkpoint('Gráfico Producción bruta');
  sheet.getRangeByIndexes(0, 25, Math.max(15 + rows.length, 16), 11).columnHidden = true;
  await softCheckpoint('Ocultar columnas auxiliares');
}

function writeSummaryChartSource(
  sheet: Excel.Worksheet,
  startColumn: number,
  rowCount: number,
  valueColumn: string,
  valueHeader: string,
): Excel.Range {
  const header = sheet.getRangeByIndexes(13, startColumn, 1, 2);
  header.values = [['Fecha', valueHeader]];
  header.format.fill.color = '#DAE0E5';
  header.format.font.bold = true;
  if (rowCount > 0) {
    const formulas = Array.from({ length: rowCount }, (_, index) => {
      const sourceRow = 15 + index;
      return [`=$A$${sourceRow}`, `=$${valueColumn}$${sourceRow}`];
    });
    sheet.getRangeByIndexes(14, startColumn, rowCount, 2).formulas = formulas;
  }
  return sheet.getRangeByIndexes(13, startColumn, rowCount + 1, 2);
}

function consolidatedFormula(results: BuildResult[], summaryRow: number, valueColumns: number[]): string {
  const terms: string[] = [];
  for (const result of results) {
    const pronoName = areaSheetNames(result.areaId).prono.replace(/'/g, "''");
    const lastRow = pronoDataLastRow(result.monthly.length, result.summary);
    for (const columnIndex of valueColumns) {
      const column = excelColumn(columnIndex);
      terms.push(`SUMIF('${pronoName}'!$A$13:$A$${lastRow},$A${summaryRow},'${pronoName}'!$${column}$13:$${column}$${lastRow})`);
    }
  }
  return terms.length ? `=${terms.join('+')}` : '=0';
}

export function pronoDataLastRow(
  historyMonthCount: number,
  summary: ReadonlyArray<{ kind: 'hist' | 'prono' }>,
): number {
  const forecastMonthCount = summary.filter((month) => month.kind === 'prono').length;
  return 12 + historyMonthCount + forecastMonthCount;
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

  const methods = resolveForecastMethods(plan);
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
