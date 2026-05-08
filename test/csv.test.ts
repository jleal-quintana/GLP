import { describe, expect, it } from 'vitest';
import { parseCsvLine } from '../src/services/csv';

describe('parseCsvLine', () => {
  it('parses quoted commas and escaped quotes', () => {
    expect(parseCsvLine('a,"b,c","d""e"')).toEqual(['a', 'b,c', 'd"e']);
  });
});
