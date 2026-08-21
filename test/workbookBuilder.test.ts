import { describe, expect, it } from 'vitest';
import { pronoDataLastRow } from '../src/excel/workbookBuilder';

describe('resumen consolidado', () => {
  it('incluye toda la historia física aunque el resumen omita meses faltantes', () => {
    const summary = [
      { kind: 'hist' as const },
      { kind: 'hist' as const },
      ...Array.from({ length: 12 }, () => ({ kind: 'prono' as const })),
    ];

    expect(pronoDataLastRow(3, summary)).toBe(27);
  });
});
