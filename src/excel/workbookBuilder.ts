import type {
  AreaWorkbookPlan,
  BuildProgressHandler,
  MissingMonthsDecisionHandler,
  MonthlyAggregate,
} from '../models/types';
import { evaluateForecast, lastNonMissing, resolveForecastMethods } from '../domain/forecast';
import { aggregateMonthly, countProductionResources, fetchAreaProduction } from '../services/capiv';
import { appendDebug, createDebugEntry } from './debugSheet';
import { writeAreaSheets } from './areaSheets';
import { appendDownloadLedgerEvent, ensureDownloadLedger } from './downloadLedger';
import { SUMMARY_SHEET, STATE_SHEET } from './names';
import { addLineChart, getOrAddSheet, writeMatrix, writeTable, writeTitle } from './sheetLayout';

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

function estimateWorkUnits(plans: AreaWorkbookPlan[]): number {
  const perAreaFixedSteps = 4;
  const globalSteps = 2;
  return plans.reduce((total, plan) => {
    const startYear = plan.override?.startYear ?? plan.selection.startYearOverride ?? plan.defaults.startYear;
    return total + countProductionResources(startYear) + perAreaFixedSteps;
  }, globalSteps);
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
    writeTable(sheet, 'A4:G4', headers, rows, 'Resumen Areas');
    writeConsolidatedMonthly(sheet, results);
    await context.sync();
  });
}

async function writeState(plans: AreaWorkbookPlan[], results: BuildResult[]): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = await getOrAddSheet(context, STATE_SHEET);
    sheet.visibility = Excel.SheetVisibility.veryHidden;
    sheet.getRange().clear();
    writeWorkbookState(sheet, { schema: 1, plans, results, savedAt: new Date().toISOString() });
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
    .map(([date, values]) => [date, values.kind, values.oil, values.gas, values.water, values.gross]);
  writeTable(sheet, 'A14:F14', ['Fecha', 'Tipo', 'Petroleo total', 'Gas total', 'Agua total', 'Bruta total'], rows, 'Resumen mensual');
  const rowCount = Math.max(2, rows.length + 1);
  addLineChart(sheet, sheet.getRangeByIndexes(13, 0, rowCount, 6), 'Consolidado total', 'I4', 'P22');
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
