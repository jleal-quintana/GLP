import type {
  AreaWorkbookPlan,
  BuildProgressHandler,
  DataOutputTarget,
  MissingMonthsDecisionHandler,
  MonthlyAggregate,
  WorkbookAreaData,
  OverwriteDecisionHandler,
} from '../models/types';
import { evaluateForecast, lastNonMissing, resolveForecastMethods } from '../domain/forecast';
import { aggregateMonthly, countProductionResources, fetchAreaProduction } from '../services/capiv';
import { appendDebug, createDebugEntry } from './debugSheet';
import { writeAreaForecastSheets } from './areaSheets';
import { writeDatabaseTable } from './databaseSheet';
import { appendDownloadLedgerEvent, ensureDownloadLedger } from './downloadLedger';
import { areaSheetNames, SUMMARY_SHEET, STATE_SHEET } from './names';
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
  }, output);
  await appendDebug(createDebugEntry(STATE_SHEET, 'ok', 'Estado guardado'));
  report('Datos actualizados', undefined, undefined, total - completed);
  await appendDebug(createDebugEntry('Fin datos', 'ok', 'Descarga de datos completada'));
}

export async function buildForecastWorkbook(
  plans: AreaWorkbookPlan[],
  onProgress?: BuildProgressHandler,
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
    await writeAreaForecastSheets(plan, data.monthly, data.middleMissingPolicy);
    report(`Pronóstico escrito ${plan.selection.areaId}`, plan, areaIndex, 1);
    results.push({ ...data, summary });
    await appendDebug(createDebugEntry(plan.selection.areaId, 'ok', `Pronóstico generado: ${summary.length} meses`));
    report(`Área ${plan.selection.areaId} completada`, plan, areaIndex, 1);
  }

  report('Escribiendo resumen consolidado', undefined, undefined, 1);
  await writeSummary(results);
  const forecastSavedAt = new Date().toISOString();
  await writeState(mergePlans(saved.plans, plans), saved.data, { dataSavedAt: saved.dataSavedAt, forecastSavedAt }, saved.dataOutput);
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
  await Excel.run(async (context) => {
    const sheet = await getOrAddSheet(context, SUMMARY_SHEET);
    sheet.getRange().clear();
    writeTitle(sheet, 'Resumen consolidado de áreas', 'Suma de históricos y proyecciones individuales');
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
    writeTable(sheet, 'A4:G4', headers, rows, 'Resumen Areas');
    writeConsolidatedMonthly(sheet, results);
    sheet.freezePanes.freezeRows(14);
    await context.sync();
  });
}

async function writeState(
  plans: AreaWorkbookPlan[],
  results: WorkbookAreaData[],
  timestamps: { dataSavedAt?: string; forecastSavedAt?: string },
  dataOutput?: DataOutputTarget,
): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = await getOrAddSheet(context, STATE_SHEET);
    sheet.visibility = Excel.SheetVisibility.veryHidden;
    sheet.getRange().clear();
    writeWorkbookState(sheet, { schema: 3, plans, results, savedAt: new Date().toISOString(), ...timestamps, dataOutput });
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
  const rowCount = Math.max(2, rows.length + 1);
  addLineChart(sheet, sheet.getRangeByIndexes(13, 0, rowCount, 6), 'Consolidado total', 'I4', 'P22');
}

function consolidatedFormula(results: BuildResult[], summaryRow: number, valueColumns: number[]): string {
  const terms: string[] = [];
  for (const result of results) {
    const pronoName = areaSheetNames(result.areaId).prono.replace(/'/g, "''");
    const lastRow = 12 + result.summary.length;
    for (const columnIndex of valueColumns) {
      const column = excelColumn(columnIndex);
      terms.push(`SUMIF('${pronoName}'!$A$13:$A$${lastRow},$A${summaryRow},'${pronoName}'!$${column}$13:$${column}$${lastRow})`);
    }
  }
  return terms.length ? `=${terms.join('+')}` : '=0';
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
