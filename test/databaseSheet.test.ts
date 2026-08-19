import { describe, expect, it } from 'vitest';
import { buildDatabaseMatrix, rangesOverlap, type DownloadedArea } from '../src/excel/databaseSheet';

const download: DownloadedArea = {
  plan: {
    selection: { areaId: 'CMOE', areaName: 'Cerro Mollar Oeste', province: 'Mendoza', companies: ['Quintana'] },
    defaults: {
      startYear: 2026,
      horizonYears: 10,
      grossMethod: 'Constante',
      oilMethod: 'Declinación Exp.',
      gasMethod: 'RGP',
      takeInitialFromHistory: true,
    },
    mode: 'update',
  },
  data: {
    areaId: 'CMOE',
    areaName: 'Cerro Mollar Oeste',
    warnings: [],
    middleMissingPolicy: 'blank',
    monthly: [{
      date: '2026-07-01', year: 2026, month: 7, oil: 10, gas: 20, water: 30, gross: 40,
      waterInjection: 5, oilWells: 2, gasWells: 1, injectorWells: 1, missing: false, missingKind: 'none',
    }],
  },
  records: [{
    areaId: 'CMOE', areaName: 'Cerro Mollar Oeste', wellId: '1', wellName: 'CMOE-1', year: 2026, month: 7,
    oil: 10, gas: 20, water: 30, waterInjection: 5, raw: {},
  }],
};

describe('base de datos de salida', () => {
  it('arma una fila mensual agregada por área', () => {
    const matrix = buildDatabaseMatrix('area', [download]);
    expect(matrix).toHaveLength(2);
    expect(matrix[0]).toContain('Código área');
    expect(matrix[1]).toEqual(['2026-07-01', 2026, 7, 'CMOE', 'Cerro Mollar Oeste', 'Mendoza', 10, 20, 30, 40, 5, 2, 1, 1]);
  });

  it('arma una fila detallada por pozo y mes', () => {
    const matrix = buildDatabaseMatrix('well', [download]);
    expect(matrix).toHaveLength(2);
    expect(matrix[0]).toContain('ID pozo');
    expect(matrix[1]).toContain('CMOE-1');
  });

  it('detecta rangos que se pisan y descarta los separados', () => {
    const other = { rowIndex: 5, columnIndex: 5, rowCount: 3, columnCount: 3 };
    expect(rangesOverlap(4, 4, 3, 3, other)).toBe(true);
    expect(rangesOverlap(0, 0, 2, 2, other)).toBe(false);
  });
});
