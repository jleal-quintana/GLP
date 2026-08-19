import { normalizeAssetGroups } from '../domain/assets';
import type { AreaWorkbookPlan, AssetGroup, DataOutputTarget, WorkbookAreaData } from '../models/types';
import { STATE_SHEET } from './names';

interface StoredWorkbookState {
  schema: number;
  plans: AreaWorkbookPlan[];
  results?: WorkbookAreaData[];
  savedAt?: string;
  dataSavedAt?: string;
  forecastSavedAt?: string;
  dataOutput?: DataOutputTarget;
  assetGroups?: AssetGroup[];
}

export interface SavedWorkbookPlans {
  plans: AreaWorkbookPlan[];
  data: WorkbookAreaData[];
  savedAt?: string;
  dataSavedAt?: string;
  forecastSavedAt?: string;
  dataOutput?: DataOutputTarget;
  assetGroups: AssetGroup[];
}

export function parseWorkbookState(serialized: string): SavedWorkbookPlans {
  let parsed: unknown;
  try {
    parsed = JSON.parse(serialized);
  } catch {
    throw new Error('El estado interno del libro no contiene JSON válido.');
  }

  if (!isStoredWorkbookState(parsed)) {
    throw new Error('El estado interno del libro no tiene un formato compatible con esta versión de CapIV.');
  }

  return {
    savedAt: parsed.savedAt,
    dataSavedAt: parsed.dataSavedAt ?? parsed.savedAt,
    forecastSavedAt: parsed.forecastSavedAt,
    dataOutput: isDataOutputTarget(parsed.dataOutput) ? parsed.dataOutput : undefined,
    assetGroups: normalizeAssetGroups(Array.isArray(parsed.assetGroups) ? parsed.assetGroups.filter(isAssetGroup) : []),
    plans: parsed.plans.map((plan) => ({ ...plan, mode: 'update' })),
    data: (parsed.results ?? []).filter(isWorkbookAreaData),
  };
}

export async function readSavedWorkbookPlans(): Promise<SavedWorkbookPlans | null> {
  if (typeof Excel === 'undefined') {
    throw new Error('Abrí CapIV desde Microsoft Excel para actualizar el libro.');
  }

  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItemOrNullObject(STATE_SHEET);
    sheet.load('isNullObject');
    await context.sync();
    if (sheet.isNullObject) return null;

    const usedRange = sheet.getUsedRangeOrNullObject(true);
    usedRange.load(['isNullObject', 'values']);
    await context.sync();
    if (usedRange.isNullObject) return null;

    const serialized = usedRange.values
      .map((row) => String(row[0] ?? ''))
      .join('');
    if (!serialized.trim()) return null;
    return parseWorkbookState(serialized);
  });
}

function isStoredWorkbookState(value: unknown): value is StoredWorkbookState {
  if (!value || typeof value !== 'object') return false;
  const candidate = value as Partial<StoredWorkbookState>;
  if (![1, 2, 3, 4].includes(candidate.schema ?? 0) || !Array.isArray(candidate.plans) || candidate.plans.length === 0) return false;
  return candidate.plans.every(isAreaWorkbookPlan);
}

function isAssetGroup(value: unknown): value is AssetGroup {
  if (!value || typeof value !== 'object') return false;
  const group = value as Partial<AssetGroup>;
  return Boolean(group.id && group.name && Array.isArray(group.areaIds) && group.areaIds.every((areaId) => typeof areaId === 'string'));
}

function isDataOutputTarget(value: unknown): value is DataOutputTarget {
  if (!value || typeof value !== 'object') return false;
  const target = value as Partial<DataOutputTarget>;
  return Boolean(
    target.sheetName &&
      target.startAddress &&
      target.tableName &&
      (target.granularity === 'area' || target.granularity === 'well'),
  );
}

function isWorkbookAreaData(value: unknown): value is WorkbookAreaData {
  if (!value || typeof value !== 'object') return false;
  const data = value as Partial<WorkbookAreaData>;
  return Boolean(
    data.areaId &&
      data.areaName &&
      Array.isArray(data.monthly) &&
      Array.isArray(data.warnings) &&
      (data.middleMissingPolicy === 'blank' || data.middleMissingPolicy === 'zero'),
  );
}

function isAreaWorkbookPlan(value: unknown): value is AreaWorkbookPlan {
  if (!value || typeof value !== 'object') return false;
  const plan = value as Partial<AreaWorkbookPlan>;
  return Boolean(
    plan.selection?.areaId &&
      plan.selection.areaName &&
      plan.defaults &&
      Number.isFinite(plan.defaults.startYear) &&
      Number.isFinite(plan.defaults.horizonYears),
  );
}
