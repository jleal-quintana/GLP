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

describe('parseWorkbookState', () => {
  it('recupera los planes y fuerza el modo de actualización', () => {
    const state = parseWorkbookState(JSON.stringify({ schema: 1, savedAt: '2026-08-19T12:00:00Z', plans: [plan] }));
    expect(state.savedAt).toBe('2026-08-19T12:00:00Z');
    expect(state.plans).toHaveLength(1);
    expect(state.plans[0].selection.areaId).toBe('CMOE');
    expect(state.plans[0].mode).toBe('update');
  });

  it('rechaza un estado incompatible', () => {
    expect(() => parseWorkbookState(JSON.stringify({ schema: 2, plans: [plan] }))).toThrow('formato compatible');
  });
});
