import { describe, expect, it } from 'vitest';
import { parseWorkbookState } from '../src/excel/workbookState';

const plan = {
  selection: {
    province: 'Mendoza',
    areaId: 'CMOE',
    areaName: 'Cañadón Amarillo',
    companies: ['Quintana'],
  },
  defaults: {
    startYear: 2024,
    horizonYears: 10,
    grossMethod: 'Constante',
    oilMethod: 'Declinación Exp.',
    gasMethod: 'RGP',
    takeInitialFromHistory: true,
  },
  mode: 'regenerate',
};

const result = {
  areaId: 'CMOE',
  areaName: 'Cañadón Amarillo',
  monthly: [],
  warnings: [],
  middleMissingPolicy: 'blank',
};

describe('parseWorkbookState', () => {
  it('recupera los planes y fuerza el modo de actualización', () => {
    const state = parseWorkbookState(JSON.stringify({ schema: 1, savedAt: '2026-08-19T12:00:00Z', plans: [plan], results: [result] }));
    expect(state.savedAt).toBe('2026-08-19T12:00:00Z');
    expect(state.plans).toHaveLength(1);
    expect(state.plans[0].selection.areaId).toBe('CMOE');
    expect(state.plans[0].mode).toBe('update');
    expect(state.data).toEqual([result]);
    expect(state.dataSavedAt).toBe('2026-08-19T12:00:00Z');
  });

  it('recupera por separado las fechas de datos y pronósticos del esquema 2', () => {
    const state = parseWorkbookState(JSON.stringify({
      schema: 2,
      savedAt: '2026-08-19T12:00:00Z',
      dataSavedAt: '2026-08-19T12:01:00Z',
      forecastSavedAt: '2026-08-19T12:02:00Z',
      plans: [plan],
      results: [result],
    }));
    expect(state.dataSavedAt).toBe('2026-08-19T12:01:00Z');
    expect(state.forecastSavedAt).toBe('2026-08-19T12:02:00Z');
    expect(state.data[0].areaId).toBe('CMOE');
  });

  it('recupera el destino de la tabla del esquema 3', () => {
    const dataOutput = { sheetName: 'Base', startAddress: 'B3', granularity: 'area', tableName: 'CapIV_Datos_Base_B3' };
    const state = parseWorkbookState(JSON.stringify({ schema: 3, plans: [plan], results: [result], dataOutput }));
    expect(state.dataOutput).toEqual(dataOutput);
  });

  it('recupera los activos del esquema 4', () => {
    const assetGroups = [{ id: 'mendoza', name: 'Mendoza', areaIds: ['CMOE'] }];
    const state = parseWorkbookState(JSON.stringify({ schema: 4, plans: [plan], results: [result], assetGroups }));
    expect(state.assetGroups).toEqual(assetGroups);
  });

  it('rechaza un estado incompatible', () => {
    expect(() => parseWorkbookState(JSON.stringify({ schema: 5, plans: [plan] }))).toThrow('formato compatible');
  });
});
