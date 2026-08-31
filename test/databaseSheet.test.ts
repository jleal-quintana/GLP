import { describe, expect, it } from 'vitest';
import { buildDatabaseMatrix, nextAvailableDataSheetName, rangesOverlap, type DownloadedArea } from '../src/excel/databaseSheet';

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
      waterInjection: 5, gasInjection: 8, oilWells: 2, gasWells: 1, injectorWells: 1, gasInjectorWells: 2,
      missing: false, missingKind: 'none',
    }],
  },
  records: [{
    areaId: 'CMOE', areaName: 'Cerro Mollar Oeste', wellId: '1', wellName: 'CMOE-1', year: 2026, month: 7,
    oil: 10, gas: 20, water: 30, waterInjection: 5, gasInjection: 8, raw: {},
  }],
};

describe('base de datos de salida', () => {
  it('arma una fila mensual agregada por área', () => {
    const matrix = buildDatabaseMatrix('area', [download]);
    expect(matrix).toHaveLength(2);
    expect(matrix[0]).toContain('Código área');
    expect(matrix[0]).toContain('Gas inyectado');
    expect(matrix[0]).toContain('Inyectores gas');
    expect(matrix[1]).toEqual(['2026-07-01', 2026, 7, 'CMOE', 'Cerro Mollar Oeste', 'Mendoza', 10, 20, 30, 40, 5, 8, 2, 1, 1, 2]);
  });

  it('arma una fila detallada por pozo y mes', () => {
    const matrix = buildDatabaseMatrix('well', [download]);
    expect(matrix).toHaveLength(2);
    expect(matrix[0]).toContain('ID pozo');
    expect(matrix[0]).toContain('Gas inyectado');
    expect(matrix[1]).toEqual([2026, 7, 'CMOE', 'Cerro Mollar Oeste', 'Mendoza', '1', 'CMOE-1', 10, 20, 30, 5, 8]);
  });

  it('detecta rangos que se pisan y descarta los separados', () => {
    const other = { rowIndex: 5, columnIndex: 5, rowCount: 3, columnCount: 3 };
    expect(rangesOverlap(4, 4, 3, 3, other)).toBe(true);
    expect(rangesOverlap(0, 0, 2, 2, other)).toBe(false);
  });

  it('elige un nombre de hoja nuevo sin pisar hojas existentes', () => {
    expect(nextAvailableDataSheetName(['Sheet1'])).toBe('CapIV_Datos');
    expect(nextAvailableDataSheetName(['capiv_datos', 'CapIV_Datos_2'])).toBe('CapIV_Datos_3');
  });

  it('nombra la hoja con las áreas filtradas y respeta el límite de 31 caracteres', () => {
    expect(nextAvailableDataSheetName([], [{ areaId: 'CMOE', areaName: 'Cerro Mollar Oeste' }])).toBe('CapIV_Datos_Cerro_Mollar_Oeste');
    expect(nextAvailableDataSheetName([], [
      { areaId: 'AAA', areaName: 'AreaX' },
      { areaId: 'BBB', areaName: 'AreaY' },
    ])).toBe('CapIV_Datos_AreaX_AreaY');
    expect(nextAvailableDataSheetName([], [
      { areaId: 'CMOE', areaName: 'Cerro Mollar Oeste' },
      { areaId: 'EFO', areaName: 'El Fortín' },
    ])).toBe('CapIV_Datos_CMOE_EFO');
    expect(nextAvailableDataSheetName(['CapIV_Datos_CMOE'], [{ areaId: 'CMOE', areaName: 'CMOE' }])).toBe('CapIV_Datos_CMOE_2');
    expect(nextAvailableDataSheetName([], [
      { areaId: 'CMOE', areaName: 'Cerro Mollar Oeste' },
      { areaId: 'EFO', areaName: 'El Fortín' },
    ]).length).toBeLessThanOrEqual(31);
  });
});
