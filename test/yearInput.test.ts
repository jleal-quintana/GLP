import { describe, expect, it } from 'vitest';
import { isValidStartYear, isYearDraft } from '../src/taskpane/App';

describe('start year editing', () => {
  it('allows an empty or partial numeric draft while the user is typing', () => {
    expect(isYearDraft('')).toBe(true);
    expect(isYearDraft('2')).toBe(true);
    expect(isYearDraft('202')).toBe(true);
    expect(isYearDraft('2020')).toBe(true);
    expect(isYearDraft('20200')).toBe(false);
    expect(isYearDraft('20a0')).toBe(false);
  });

  it('accepts only complete four-digit years in the available range', () => {
    expect(isValidStartYear('2006')).toBe(true);
    expect(isValidStartYear('2020')).toBe(true);
    expect(isValidStartYear('202')).toBe(false);
    expect(isValidStartYear('')).toBe(false);
    expect(isValidStartYear('2005')).toBe(false);
    expect(isValidStartYear(String(new Date().getFullYear() + 1))).toBe(false);
  });
});
