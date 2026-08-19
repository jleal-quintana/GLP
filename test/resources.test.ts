import { describe, expect, it } from 'vitest';
import { discoverCapivResources } from '../src/services/capiv';

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
});
