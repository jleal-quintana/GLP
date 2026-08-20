import { describe, expect, it } from 'vitest';
import { createAssetId, normalizeAssetGroups } from '../src/domain/assets';
import { assetSheetName } from '../src/excel/names';

describe('activos', () => {
  it('crea identificadores y nombres de hoja seguros', () => {
    expect(createAssetId('Mendoza Norte', [])).toBe('mendoza-norte');
    expect(createAssetId('Mendoza Norte', [{ id: 'mendoza-norte', name: 'Mendoza Norte', areaIds: ['A'] }])).toBe('mendoza-norte-2');
    expect(assetSheetName('Mendoza Norte / Secundaria')).toBe('Activo_Mendoza_Norte_Secundari');
    expect(assetSheetName('Mendoza Norte / Secundaria').length).toBeLessThanOrEqual(31);
  });

  it('evita que una concesión pertenezca a dos activos', () => {
    const groups = normalizeAssetGroups([
      { id: 'mza', name: 'Mendoza', areaIds: ['A', 'B'] },
      { id: 'otro', name: 'Otro', areaIds: ['B', 'C'] },
    ], ['A', 'B', 'C']);
    expect(groups).toEqual([
      { id: 'mza', name: 'Mendoza', areaIds: ['A', 'B'] },
      { id: 'otro', name: 'Otro', areaIds: ['C'] },
    ]);
  });
});
