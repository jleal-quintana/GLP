import type { AreaWorkbookPlan, ForecastDefaults, ForecastMethod, MonthlyAggregate } from '../models/types';

type MonthlyValueKey = 'oil' | 'gas' | 'water' | 'gross' | 'oilWells' | 'gasWells' | 'injectorWells';

export function resolveForecastMethods(plan: AreaWorkbookPlan): Pick<ForecastDefaults, 'grossMethod' | 'oilMethod' | 'gasMethod'> {
  return {
    grossMethod: plan.override?.grossMethod ?? plan.defaults.grossMethod,
    oilMethod: plan.override?.oilMethod ?? plan.defaults.oilMethod,
    gasMethod: plan.override?.gasMethod ?? plan.defaults.gasMethod,
  };
}

export function lastNonMissing(monthly: MonthlyAggregate[], key: MonthlyValueKey): number {
  for (let i = monthly.length - 1; i >= 0; i--) {
    const month = monthly[i];
    if (!month.missing && month[key] > 0) return month[key];
  }
  return 0;
}

export function evaluateForecast(method: ForecastMethod, initial: number, di: number, b: number, tYears: number): number {
  if (!Number.isFinite(initial) || initial <= 0) return 0;
  if (method === 'Constante') return initial;
  if (method === 'Declinación Exp.') return initial * Math.exp(-di * tYears);
  if (method === 'Declinación Hip.' || method === 'HypMod' || method === 'Rap Np') {
    return initial / Math.pow(1 + b * di * tYears, 1 / b);
  }
  return initial;
}

export function forecastFormula(
  methodCell: string,
  initialValue: number,
  diCell: string,
  bCell: string,
  tYears: number,
  rgpCell?: string,
  oilReference?: string | number,
): string {
  const initial = Number.isFinite(initialValue) ? initialValue : 0;
  const oil = oilReference ?? initial;
  const t = Number(tYears.toFixed(6));
  return `=IF(${methodCell}="Constante",${initial},IF(${methodCell}="Declinación Exp.",${initial}*EXP(-${diCell}*${t}),IF(${methodCell}="Declinación Hip.",${initial}/POWER(1+${bCell}*${diCell}*${t},1/${bCell}),IF(${methodCell}="HypMod",${initial}/POWER(1+${bCell}*${diCell}*${t},1/${bCell}),IF(${methodCell}="RGP",${rgpCell ?? 0}*${oil}/1000,IF(${methodCell}="Rap Np",${initial}/POWER(1+${bCell}*${diCell}*${t},1/${bCell}),${initial}))))))`;
}

export function nextMonth(date?: string): string {
  if (!date) return '';
  const parsed = new Date(`${date}T00:00:00`);
  parsed.setMonth(parsed.getMonth() + 1);
  return parsed.toISOString().slice(0, 10);
}
