import { describe, expect, it } from 'vitest';
import { DEFAULT_DECLINE, evaluateForecast, forecastFormula, lastNonMissing, nextMonth, resolveAreaParams } from '../src/domain/forecast';
import type { ForecastDefaults, MonthlyAggregate } from '../src/models/types';

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

  it('can express RGP gas as a function of the current oil forecast row', () => {
    expect(forecastFormula('B6', 100, 'E8', 'E9', 1, 'E10', 'C13')).toContain('IF(B6="RGP",E10*C13/1000');
  });

  it('can switch between a historical and an editable manual initial value', () => {
    expect(forecastFormula('B5', 'IF($B$7="Sí",100,$H$5)', 'E6', 'E7', 1)).toContain(
      'IF($B$7="Sí",100,$H$5)',
    );
  });

  it('resolves complete area params from defaults and partial overrides', () => {
    const defaults: ForecastDefaults = {
      startYear: 2015,
      horizonYears: 10,
      grossMethod: 'Constante',
      oilMethod: 'Declinación Exp.',
      gasMethod: 'RGP',
      takeInitialFromHistory: true,
    };

    expect(resolveAreaParams(defaults)).toEqual({
      grossMethod: 'Constante',
      oilMethod: 'Declinación Exp.',
      gasMethod: 'RGP',
      takeInitialFromHistory: true,
      ...DEFAULT_DECLINE,
    });

    expect(resolveAreaParams(defaults, {
      areaId: 'A',
      oilMethod: 'Declinación Hip.',
      oilDi: 0.2,
      oilB: 1.1,
      takeInitialFromHistory: false,
    })).toMatchObject({
      grossMethod: 'Constante',
      oilMethod: 'Declinación Hip.',
      oilDi: 0.2,
      oilB: 1.1,
      takeInitialFromHistory: false,
      gasDi: DEFAULT_DECLINE.gasDi,
    });
  });
});
