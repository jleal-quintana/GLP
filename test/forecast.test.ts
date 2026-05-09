import { describe, expect, it } from 'vitest';
import { evaluateForecast, lastNonMissing, nextMonth } from '../src/domain/forecast';
import type { MonthlyAggregate } from '../src/models/types';

function month(date: string, oil: number, missing = false): MonthlyAggregate {
  const [year, monthNumber] = date.split('-').map(Number);
  return {
    date,
    year,
    month: monthNumber,
    oil,
    gas: oil * 10,
    water: oil * 2,
    gross: oil * 3,
    waterInjection: 0,
    oilWells: oil > 0 ? 1 : 0,
    gasWells: 0,
    injectorWells: 0,
    missing,
    missingKind: missing ? 'middle' : 'none',
  };
}

describe('forecast projection', () => {
  it('uses the last non-missing positive value as the projection initial', () => {
    const rows = [month('2025-01-01', 10), month('2025-02-01', 0, true), month('2025-03-01', 0)];

    expect(lastNonMissing(rows, 'oil')).toBe(10);
  });

  it('evaluates supported decline methods consistently', () => {
    expect(evaluateForecast('Constante', 100, 0.12, 0.7, 1)).toBe(100);
    expect(evaluateForecast('Declinación Exp.', 100, 0.12, 0.7, 1)).toBeCloseTo(88.692, 3);
    expect(evaluateForecast('Declinación Hip.', 100, 0.12, 0.7, 1)).toBeCloseTo(89.117, 3);
  });

  it('moves a published month to the next projection month', () => {
    expect(nextMonth('2025-12-01')).toBe('2026-01-01');
  });
});
