import type { AreaCatalogItem, CapituloIvDownloadEventHandler, MonthlyAggregate, ProductionRecord } from '../models/types';
import { parseCsv, parseCsvLine } from './csv';
import { fetchWithRetry } from './http';

const API_ACTION_URL = 'https://datos.gob.ar/api/3/action';
const WELLS_CAP_IV_RESOURCE = 'energia_cb5c0f04-7835-45cd-b982-3e25ca7d7751';

const PROD_RESOURCE_BY_YEAR: Record<number, string> = {
  2026: 'energia_fb7a47a0-cba9-4667-a004-6f6c1c346c23',
  2025: 'energia_d774b5d7-0756-48fe-88f2-8729b57b22da',
  2024: 'energia_43a09dce-1742-44d0-bc13-f193deaab563',
  2023: 'energia_231c39b3-e81e-4398-af8d-b115807f2c25',
  2022: 'energia_876b3746-85e2-4039-adeb-b1354436159f',
  2021: 'energia_465be754-a372-4c31-b855-81dc5fe3309f',
  2020: 'energia_c4a4a6a0-e75a-4e12-ae5c-54d53a70348c',
  2019: 'energia_8bc0d61c-0408-43d4-a7bc-7178fcb5d37e',
  2018: 'energia_333fd72a-9b83-4bc1-bc94-0f5940b52331',
  2017: 'energia_df4857e1-7c3f-4980-b5b5-184fe78bfcf0',
  2016: 'energia_d8539ae8-0a71-4339-a16c-139b21bd2cd0',
  2015: 'energia_e375aa35-fd8d-41d6-aa0c-e6879ca567a1',
  2014: 'energia_cd9813a7-1e19-4f60-a02a-7903dd81aff7',
  2013: 'energia_bc7ac8fe-2cec-4dab-acdd-322ea1ccc887',
  2012: 'energia_0dce0e75-1556-47ee-8615-1955fbd54ade',
  2011: 'energia_4817272c-7365-4bdd-b02d-75b118218b10',
  2010: 'energia_364ca28e-d069-4bd6-8771-925f0db152a8',
  2009: 'energia_48585038-055a-4437-bb1d-4fe36073f453',
  2008: 'energia_3430b5a8-a516-42ca-a47d-2e1ce45925fb',
  2007: 'energia_be663a63-f020-4e28-8f31-c5f81d47554d',
  2006: 'energia_4e1c55e5-1f1b-4fc8-aa37-2080d9795f29',
};

const NC_PRODUCTION_RESOURCE = 'energia_b5b58cdc-9e07-41f9-b392-fb9ec68b0725';
const urlCache = new Map<string, Promise<string>>();
const DOWNLOAD_PAUSE_MS = 250;

async function getResourceDownloadUrl(resourceId: string): Promise<string> {
  const cached = urlCache.get(resourceId);
  if (cached) return cached;
  const request = getResourceDownloadUrlUncached(resourceId);
  urlCache.set(resourceId, request);
  return request;
}

async function getResourceDownloadUrlUncached(resourceId: string): Promise<string> {
  const response = await fetchWithRetry(`${API_ACTION_URL}/resource_show?id=${resourceId}`);
  if (!response.ok) throw new Error(`resource_show failed ${response.status} for ${resourceId}`);
  const data = await response.json();
  if (!data.success || !data.result?.url) throw new Error(`resource_show returned no url for ${resourceId}`);
  return String(data.result.url);
}

async function downloadResourceCsv(resourceId: string): Promise<Record<string, string>[]> {
  const url = await getResourceDownloadUrl(resourceId);
  const response = await fetchWithRetry(url);
  if (!response.ok) throw new Error(`CSV download failed ${response.status} for ${resourceId}`);
  return parseCsv(await response.text());
}

async function streamResourceCsv(
  resourceId: string,
  onRecord: (record: Record<string, string>) => void,
): Promise<number> {
  const url = await getResourceDownloadUrl(resourceId);
  const response = await fetchWithRetry(url);
  if (!response.ok) throw new Error(`CSV download failed ${response.status} for ${resourceId}`);

  if (!response.body) {
    const records = parseCsv(await response.text());
    records.forEach(onRecord);
    return records.length;
  }

  const reader = response.body.getReader();
  const decoder = new TextDecoder('utf-8');
  let pending = '';
  let headers: string[] | null = null;
  let rowCount = 0;

  while (true) {
    const { done, value } = await reader.read();
    pending += decoder.decode(value, { stream: !done });
    const lines = pending.split(/\r?\n/);
    pending = lines.pop() ?? '';

    for (const line of lines) {
      if (!line) continue;
      if (!headers) {
        headers = parseCsvLine(line).map((h, index) => (index === 0 ? h.replace(/^\uFEFF/, '') : h));
        continue;
      }
      onRecord(recordFromLine(headers, line));
      rowCount++;
    }

    if (done) break;
  }

  if (pending.trim()) {
    if (!headers) {
      headers = parseCsvLine(pending).map((h, index) => (index === 0 ? h.replace(/^\uFEFF/, '') : h));
    } else {
      onRecord(recordFromLine(headers, pending));
      rowCount++;
    }
  }

  return rowCount;
}

function recordFromLine(headers: string[], line: string): Record<string, string> {
  const cols = parseCsvLine(line);
  const record: Record<string, string> = {};
  headers.forEach((header, index) => {
    record[header] = cols[index] ?? '';
  });
  return record;
}

function text(record: Record<string, string>, ...keys: string[]): string {
  for (const key of keys) {
    const value = record[key];
    if (value !== undefined && value !== null && String(value).trim() !== '') {
      return String(value).trim();
    }
  }
  return '';
}

function numberValue(record: Record<string, string>, ...keys: string[]): number {
  const raw = text(record, ...keys).replace(',', '.');
  const parsed = Number(raw);
  return Number.isFinite(parsed) ? parsed : 0;
}

export async function fetchAreaCatalog(): Promise<AreaCatalogItem[]> {
  const records = await downloadResourceCsv(WELLS_CAP_IV_RESOURCE);
  const byArea = new Map<string, AreaCatalogItem>();

  for (const record of records) {
    const areaId = text(record, 'cod_area', 'idareapermisoconcesion');
    const areaName = text(record, 'area', 'areapermisoconcesion');
    if (!areaId || !areaName) continue;

    const existing = byArea.get(areaId);
    const company = text(record, 'empresa');
    if (existing) {
      if (company && !existing.companies.includes(company)) existing.companies.push(company);
      continue;
    }

    byArea.set(areaId, {
      province: text(record, 'provincia') || 'Sin provincia',
      areaId,
      areaName,
      basin: text(record, 'cuenca'),
      companies: company ? [company] : [],
    });
  }

  return [...byArea.values()].sort((a, b) =>
    `${a.province}|${a.areaName}`.localeCompare(`${b.province}|${b.areaName}`, 'es'),
  );
}

function normalizeProductionRecord(record: Record<string, string>, areaId: string, areaName: string): ProductionRecord | null {
  const recordAreaId = text(record, 'idareapermisoconcesion', 'cod_area');
  if (recordAreaId !== areaId) return null;

  const year = numberValue(record, 'anio');
  const month = numberValue(record, 'mes');
  if (!Number.isInteger(year) || !Number.isInteger(month)) return null;

  return {
    areaId,
    areaName,
    wellId: text(record, 'idpozo'),
    wellName: text(record, 'sigla'),
    year,
    month,
    oil: numberValue(record, 'prod_pet', 'prod_petroleo', 'petroleo'),
    gas: numberValue(record, 'prod_gas', 'gas'),
    water: numberValue(record, 'prod_agua', 'agua'),
    waterInjection: numberValue(record, 'iny_agua', 'agua_iny', 'inyeccion_agua'),
    raw: {},
  };
}

export async function fetchAreaProduction(
  area: AreaCatalogItem,
  startYear: number,
  onStep?: (message: string) => void,
  onEvent?: CapituloIvDownloadEventHandler,
): Promise<ProductionRecord[]> {
  const currentYear = new Date().getFullYear();
  const records: ProductionRecord[] = [];
  const seen = new Set<string>();

  for (let year = startYear; year <= currentYear; year++) {
    const resourceId = PROD_RESOURCE_BY_YEAR[year];
    if (!resourceId) continue;
    await delay(DOWNLOAD_PAUSE_MS);
    onStep?.(`Descargando produccion convencional ${area.areaId} ${year}`);
    await onEvent?.({ type: 'resource_started', areaId: area.areaId, source: 'convencional', year });
    let matched = 0;
    const matchedRows: ProductionRecord[] = [];
    const rows = await streamResourceCsv(resourceId, (row) => {
      const normalized = normalizeProductionRecord(row, area.areaId, area.areaName);
      if (!normalized) return;
      const key = `${normalized.wellName}|${normalized.wellId}|${normalized.year}|${normalized.month}|conv`;
      if (seen.has(key)) return;
      seen.add(key);
      records.push(normalized);
      matchedRows.push(normalized);
      matched++;
    });
    onStep?.(`Descargado produccion convencional ${area.areaId} ${year}: ${matched} de ${rows} filas`);
    await onEvent?.({
      type: 'resource_completed',
      areaId: area.areaId,
      source: 'convencional',
      year,
      scannedRows: rows,
      matchedRows: matched,
      records: matchedRows,
    });
  }

  await delay(DOWNLOAD_PAUSE_MS);
  onStep?.(`Descargando produccion no convencional ${area.areaId}`);
  await onEvent?.({ type: 'resource_started', areaId: area.areaId, source: 'no convencional' });
  let ncMatched = 0;
  const ncMatchedRows: ProductionRecord[] = [];
  const ncRows = await streamResourceCsv(NC_PRODUCTION_RESOURCE, (row) => {
    const normalized = normalizeProductionRecord(row, area.areaId, area.areaName);
    if (!normalized || normalized.year < startYear) return;
    const key = `${normalized.wellName}|${normalized.wellId}|${normalized.year}|${normalized.month}|nc`;
    if (seen.has(key)) return;
    seen.add(key);
    records.push(normalized);
    ncMatchedRows.push(normalized);
    ncMatched++;
  });
  onStep?.(`Descargado produccion no convencional ${area.areaId}: ${ncMatched} de ${ncRows} filas`);
  await onEvent?.({
    type: 'resource_completed',
    areaId: area.areaId,
    source: 'no convencional',
    scannedRows: ncRows,
    matchedRows: ncMatched,
    records: ncMatchedRows,
  });

  const sorted = records.sort((a, b) => a.year - b.year || a.month - b.month || a.wellName.localeCompare(b.wellName));
  await onEvent?.({ type: 'completed', areaId: area.areaId, records: sorted });
  return sorted;
}

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

export function countProductionResources(startYear: number): number {
  const currentYear = new Date().getFullYear();
  let count = 1;
  for (let year = startYear; year <= currentYear; year++) {
    if (PROD_RESOURCE_BY_YEAR[year]) count++;
  }
  return count;
}

export function aggregateMonthly(records: ProductionRecord[], startYear: number): MonthlyAggregate[] {
  const byMonth = new Map<string, ProductionRecord[]>();
  for (const record of records) {
    const key = `${record.year}-${String(record.month).padStart(2, '0')}`;
    const monthRecords = byMonth.get(key);
    if (monthRecords) monthRecords.push(record);
    else byMonth.set(key, [record]);
  }

  const publishedKeys = [...byMonth.keys()].sort();
  const lastKey = publishedKeys[publishedKeys.length - 1];
  if (!lastKey) return [];

  const [lastYear, lastMonth] = lastKey.split('-').map(Number);
  const firstKey = publishedKeys[0];
  const output: MonthlyAggregate[] = [];

  let year = startYear;
  let month = 1;
  while (year < lastYear || (year === lastYear && month <= lastMonth)) {
    const key = `${year}-${String(month).padStart(2, '0')}`;
    const rows = byMonth.get(key) ?? [];
    const missing = rows.length === 0;
    const oilWells = new Set(rows.filter((r) => r.oil > 0).map((r) => r.wellName || r.wellId));
    const gasWells = new Set(rows.filter((r) => r.gas > 0).map((r) => r.wellName || r.wellId));
    const injectorWells = new Set(rows.filter((r) => r.waterInjection > 0).map((r) => r.wellName || r.wellId));

    const oil = sum(rows, 'oil');
    const water = sum(rows, 'water');
    output.push({
      date: `${year}-${String(month).padStart(2, '0')}-01`,
      year,
      month,
      oil,
      gas: sum(rows, 'gas'),
      water,
      gross: oil + water,
      waterInjection: sum(rows, 'waterInjection'),
      oilWells: oilWells.size,
      gasWells: gasWells.size,
      injectorWells: injectorWells.size,
      missing,
      missingKind: !missing ? 'none' : key < firstKey ? 'leading' : 'middle',
    });

    if (month === 12) {
      year++;
      month = 1;
    } else {
      month++;
    }
  }

  return output;
}

function sum(rows: ProductionRecord[], key: 'oil' | 'gas' | 'water' | 'waterInjection'): number {
  return rows.reduce((total, row) => total + row[key], 0);
}
