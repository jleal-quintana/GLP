import { describe, expect, it } from 'vitest';
import { aggregateMonthly } from '../src/services/capiv';
import type { ProductionRecord } from '../src/models/types';

function record(year: number, month: number, oil: number): ProductionRecord {
  return {
    areaId: 'AAA',
    areaName: 'Area',
    wellId: `W${month}`,
    wellName: `W${month}`,
    year,
    month,
    oil,
    gas: oil * 10,
    water: oil * 2,
    waterInjection: oil > 0 ? 1 : 0,
    raw: {},
  };
}

describe('aggregateMonthly', () => {
  it('marks missing months before first publication as leading and middle gaps as middle', () => {
    const rows = aggregateMonthly([record(2024, 3, 10), record(2024, 5, 20)], 2024);

    expect(rows.map((row) => [row.month, row.missingKind])).toEqual([
      [1, 'leading'],
      [2, 'leading'],
      [3, 'none'],
      [4, 'middle'],
      [5, 'none'],
    ]);
  });
});
