import type { AreaWorkbookPlan } from '../models/types';
import { STATE_SHEET } from './names';

interface StoredWorkbookState {
  schema: number;
  plans: AreaWorkbookPlan[];
  savedAt?: string;
}

export interface SavedWorkbookPlans {
  plans: AreaWorkbookPlan[];
  savedAt?: string;
}

export function parseWorkbookState(serialized: string): SavedWorkbookPlans {
  let parsed: unknown;
  try {
    parsed = JSON.parse(serialized);
  } catch {
    throw new Error('El estado interno del libro no contiene JSON válido.');
  }

  if (!isStoredWorkbookState(parsed)) {
    throw new Error('El estado interno del libro no tiene un formato compatible con esta versión de GLP.');
  }

  return {
    savedAt: parsed.savedAt,
    plans: parsed.plans.map((plan) => ({ ...plan, mode: 'update' })),
  };
}

export async function readSavedWorkbookPlans(): Promise<SavedWorkbookPlans | null> {
  if (typeof Excel === 'undefined') {
    throw new Error('Abrí GLP desde Microsoft Excel para actualizar el libro.');
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
  if (candidate.schema !== 1 || !Array.isArray(candidate.plans) || candidate.plans.length === 0) return false;
  return candidate.plans.every(isAreaWorkbookPlan);
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
