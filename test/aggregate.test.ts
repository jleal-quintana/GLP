import { describe, expect, it } from 'vitest';
import { aggregateMonthly, normalizeProductionRecord } from '../src/services/capiv';
import type { ProductionRecord } from '../src/models/types';

function record(year: number, month: number, oil: number, extra: Partial<ProductionRecord> = {}): ProductionRecord {
  return {
    areaId: 'AAA',
    areaName: 'Area',
    wellId: extra.wellId ?? `W${month}`,
    wellName: extra.wellName ?? `W${month}`,
    year,
    month,
    oil,
    gas: extra.gas ?? oil * 10,
    water: extra.water ?? oil * 2,
    waterInjection: extra.waterInjection ?? (oil > 0 ? 1 : 0),
    gasInjection: extra.gasInjection ?? 0,
    raw: {},
    ...extra,
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

  it('keeps oil and water series intact while adding gas injection volumes and injector wells', () => {
    const rows = aggregateMonthly([
      record(2024, 1, 10, { wellId: 'OIL', wellName: 'OIL', waterInjection: 4 }),
      record(2024, 1, 0, { wellId: 'GI', wellName: 'GI', gas: 0, water: 0, waterInjection: 0, gasInjection: 12 }),
    ], 2024);

    expect(rows[0].oil).toBe(10);
    expect(rows[0].waterInjection).toBe(4);
    expect(rows[0].injectorWells).toBe(1);
    expect(rows[0].gasInjection).toBe(12);
    expect(rows[0].gasInjectorWells).toBe(1);
    expect(rows[0].oilWells).toBe(1);
    expect(rows[0].gasWells).toBe(0);
  });
});

describe('normalizeProductionRecord', () => {
  const source = { sourceAreaId: 'CMOE' };

  it('maps official iny_gas without dropping the well-month row', () => {
    const mapped = normalizeProductionRecord({
      idareapermisoconcesion: 'CMOE',
      idpozo: '99',
      sigla: 'CMOE-GI',
      anio: '2024',
      mes: '3',
      prod_pet: '0',
      prod_gas: '0',
      prod_agua: '0',
      iny_agua: '0',
      iny_gas: '18.5',
    }, 'CMOE', 'Cerro Mollar Oeste', source);

    expect(mapped).toMatchObject({
      wellName: 'CMOE-GI',
      oil: 0,
      gas: 0,
      waterInjection: 0,
      gasInjection: 18.5,
    });
  });

  it('still maps oil, gas production and water injection from the same official row', () => {
    const mapped = normalizeProductionRecord({
      idareapermisoconcesion: 'CMOE',
      idpozo: '1',
      sigla: 'CMOE-1',
      anio: '2024',
      mes: '3',
      prod_pet: '10',
      prod_gas: '20',
      prod_agua: '30',
      iny_agua: '5',
      iny_gas: '0',
    }, 'CMOE', 'Cerro Mollar Oeste', source);

    expect(mapped).toMatchObject({
      wellName: 'CMOE-1',
      oil: 10,
      gas: 20,
      water: 30,
      waterInjection: 5,
      gasInjection: 0,
    });
  });
});
