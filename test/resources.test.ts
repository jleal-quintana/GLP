import { describe, expect, it } from 'vitest';
import { discoverCapivResources, productionSourceForArea } from '../src/services/capiv';

describe('discoverCapivResources', () => {
  it('finds the live catalog and selects the newest annual publication', () => {
    const discovered = discoverCapivResources([
      {
        id: 'wells',
        name: 'Capítulo IV - Pozos',
        format: 'CSV',
        url: 'http://example.test/wells.csv',
        last_modified: '2026-08-01',
      },
      {
        id: 'old-2025',
        name: 'Producción de Pozos de Gas y Petróleo – 2025',
        format: 'CSV',
        url: 'http://example.test/old.csv',
        last_modified: '2026-06-01',
      },
      {
        id: 'new-2025',
        name: 'Producción de Pozos de Gas y Petróleo - 2025',
        format: 'CSV',
        url: 'http://example.test/new.csv',
        last_modified: '2026-08-18',
      },
      {
        id: 'ddjj-2025',
        name: 'Producción de Pozos de Gas y Petróleo - 2025 (DDJJ abiertas y cerradas)',
        format: 'CSV',
        url: 'http://example.test/ddjj.csv',
        last_modified: '2026-08-19',
      },
      {
        id: 'nc',
        name: 'Producción de Pozos de Gas y Petróleo No Convencional',
        format: 'CSV',
        url: 'http://example.test/nc.csv',
        last_modified: '2026-08-19',
      },
    ]);

    expect(discovered.wells.id).toBe('wells');
    expect(discovered.productionByYear[2025].id).toBe('new-2025');
    expect(Object.values(discovered.productionByYear)).not.toContainEqual(expect.objectContaining({ id: 'ddjj-2025' }));
    expect(Object.values(discovered.productionByYear)).not.toContainEqual(expect.objectContaining({ id: 'nc' }));
  });

  it('fails with an actionable message when the wells catalog is absent', () => {
    expect(() => discoverCapivResources([])).toThrow('Capítulo IV - Pozos');
  });

  it('recovers the legacy EPN code used for both El Portón areas in the official 2021 CSV', () => {
    const mendoza = productionSourceForArea('EPMD', 2021);
    const neuquen = productionSourceForArea('EPNQ', 2021);

    expect(mendoza.sourceAreaId).toBe('EPN');
    expect(mendoza.siglaPattern?.test('YPF.Md.NEPnN-1029')).toBe(true);
    expect(mendoza.siglaPattern?.test('YPF.MdN.EPnN-1047(h)')).toBe(true);
    expect(mendoza.siglaPattern?.test('YPF.Nq.EPnN-1019h')).toBe(false);
    expect(neuquen.sourceAreaId).toBe('EPN');
    expect(neuquen.siglaPattern?.test('YPF.Nq.EPnN-1019h')).toBe(true);
    expect(productionSourceForArea('EPMD', 2020).sourceAreaId).toBe('EPMD');
    expect(productionSourceForArea('EPNQ', 2022).sourceAreaId).toBe('EPNQ');
  });
});
