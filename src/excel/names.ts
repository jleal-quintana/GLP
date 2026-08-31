const EXCEL_SHEET_NAME_MAX = 31;
const DATA_SHEET_BASE = 'CapIV_Datos';

export function safeSheetPrefix(areaId: string): string {
  return areaId.replace(/[^A-Za-z0-9]/g, '').slice(0, 12) || 'AREA';
}

export function sanitizeSheetToken(value: string): string {
  return value
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^A-Za-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
}

export function areaSheetNames(areaId: string) {
  const prefix = safeSheetPrefix(areaId);
  return {
    hdp: `${prefix}_HDP`,
    prono: `${prefix}_Prono`,
    pozos: `${prefix}_Pozos`,
    graficos: `${prefix}_Graficos`,
    detalle: `${prefix}_Detalle`,
  };
}

export function assetSheetName(name: string): string {
  const prefix = sanitizeSheetToken(name).slice(0, 23) || 'Activo';
  return `Activo_${prefix}`.slice(0, EXCEL_SHEET_NAME_MAX);
}

export function dataSheetName(
  areas: Array<{ areaId: string; areaName: string }> = [],
  existingNames: string[] = [],
): string {
  const occupied = new Set(existingNames.map((name) => name.toLocaleLowerCase('es-AR')));
  const nameTokens = uniqueTokens(areas.map((area) => sanitizeSheetToken(area.areaName)));
  const idTokens = uniqueTokens(areas.map((area) => sanitizeSheetToken(area.areaId) || safeSheetPrefix(area.areaId)));
  const suffix = fitDataSheetSuffix(nameTokens)
    ?? fitDataSheetSuffix(idTokens, true)
    ?? '';
  const base = suffix ? `${DATA_SHEET_BASE}_${suffix}` : DATA_SHEET_BASE;
  return nextAvailableSheetName(base, occupied);
}

function uniqueTokens(tokens: string[]): string[] {
  const seen = new Set<string>();
  const output: string[] = [];
  for (const token of tokens) {
    const key = token.toLocaleLowerCase('es-AR');
    if (!token || seen.has(key)) continue;
    seen.add(key);
    output.push(token);
  }
  return output;
}

function fitDataSheetSuffix(tokens: string[], allowTruncate = false): string | undefined {
  if (tokens.length === 0) return undefined;
  const joined = tokens.join('_');
  if (`${DATA_SHEET_BASE}_${joined}`.length <= EXCEL_SHEET_NAME_MAX) return joined;
  if (!allowTruncate) return undefined;
  const truncated = joined.slice(0, EXCEL_SHEET_NAME_MAX - DATA_SHEET_BASE.length - 1).replace(/_+$/g, '');
  return truncated || undefined;
}

function nextAvailableSheetName(base: string, occupied: Set<string>): string {
  const fitted = fitExcelSheetName(base);
  if (!occupied.has(fitted.toLocaleLowerCase('es-AR'))) return fitted;
  let suffix = 2;
  while (true) {
    const extra = `_${suffix}`;
    const candidate = fitExcelSheetName(`${base.slice(0, EXCEL_SHEET_NAME_MAX - extra.length).replace(/_+$/g, '')}${extra}`);
    if (!occupied.has(candidate.toLocaleLowerCase('es-AR'))) return candidate;
    suffix++;
  }
}

function fitExcelSheetName(name: string): string {
  return name.slice(0, EXCEL_SHEET_NAME_MAX).replace(/_+$/g, '') || DATA_SHEET_BASE;
}

export const SUMMARY_SHEET = 'Resumen_Areas';
export const STATE_SHEET = '_CapIV_State';
