import { describe, expect, it } from 'vitest';
import { forecastChartCatalog } from '../src/excel/forecastCharts';

describe('forecast chart catalog', () => {
  it('contains only analyses supported by Capítulo IV and separates each objective', () => {
    const catalog = forecastChartCatalog();
    const ids = catalog.map((item) => item.id);

    expect(ids).toEqual([
      'oil',
      'gas',
      'gross',
      'water',
      'water-cut',
      'rgp',
      'rap',
      'water-injection',
      'liquid-cumulatives',
      'wells',
      'rap-vs-np',
    ]);
    expect(ids).not.toContain('vrr');
    expect(ids).not.toContain('ipv');
    expect(ids).not.toContain('recovery-factor');
    expect(new Set(ids).size).toBe(ids.length);
  });

  it('keeps the secondary diagnostic in its own section', () => {
    const catalog = forecastChartCatalog();
    expect(catalog.find((item) => item.id === 'rap-vs-np')?.section).toBe('Diagnóstico de secundaria');
  });
});
