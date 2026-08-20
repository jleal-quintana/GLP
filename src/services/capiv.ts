import type { AreaCatalogItem, CapituloIvDownloadEventHandler, MonthlyAggregate, ProductionRecord } from '../models/types';
import { parseCsv, parseCsvLine } from './csv';
import { catalogProxyUrl, fetchWithRetry, filteredProductionProxyUrl, resourceCatalogProxyUrl } from './http';

const API_ACTION_URL = 'https://datos.gob.ar/api/3/action';
const DATASET_ID = 'produccion-de-petroleo-y-gas-por-pozo';
const DOWNLOAD_PAUSE_MS = 250;

interface CkanResource {
  id: string;
  name: string;
  format?: string;
  url: string;
  last_modified?: string;
}

export interface CapivResourceCatalog {
  wells: CkanResource;
  productionByYear: Record<number, CkanResource>;
}

export interface CatalogLoadProgress {
  step: 1 | 2 | 3;
  message: string;
}

export interface AreaProductionSource {
  sourceAreaId: string;
  siglaPattern?: RegExp;
  note?: string;
}

let resourceCatalogPromise: Promise<CapivResourceCatalog> | undefined;

export function discoverCapivResources(resources: CkanResource[]): CapivResourceCatalog {
  const csvResources = resources.filter((resource) => resource.format?.toUpperCase() === 'CSV');
  const wells = newestResource(csvResources.filter((resource) => normalize(resource.name) === 'capitulo iv pozos'));
  if (!wells) throw new Error('El catálogo oficial no publicó el recurso CSV "Capítulo IV - Pozos".');

  const candidatesByYear = new Map<number, CkanResource[]>();
  for (const resource of csvResources) {
    const normalizedName = normalize(resource.name);
    if (!normalizedName.includes('produccion de pozos de gas y petroleo')) continue;
    if (normalizedName.includes('ddjj abiertas y cerradas') || normalizedName.includes('no convencional')) continue;
    const match = normalizedName.match(/\b(20\d{2}|19\d{2})\b/);
    if (!match) continue;
    const year = Number(match[1]);
    const current = candidatesByYear.get(year) ?? [];
    current.push(resource);
    candidatesByYear.set(year, current);
  }

  const productionByYear: Record<number, CkanResource> = {};
  for (const [year, candidates] of candidatesByYear) {
    const selected = newestResource(candidates);
    if (selected) productionByYear[year] = selected;
  }
  return { wells, productionByYear };
}

async function loadResourceCatalog(): Promise<CapivResourceCatalog> {
  if (!resourceCatalogPromise) {
    resourceCatalogPromise = (async () => {
      const response = await fetchWithRetry(
        resourceCatalogProxyUrl() ?? `${API_ACTION_URL}/package_show?id=${DATASET_ID}`,
      );
      if (!response.ok) throw new Error(`No se pudo consultar el catálogo oficial (HTTP ${response.status}).`);
      const data = await response.json();
      if (!data.success || !Array.isArray(data.result?.resources)) {
        throw new Error('El catálogo oficial devolvió una respuesta inválida.');
      }
      return discoverCapivResources(data.result.resources as CkanResource[]);
    })().catch((error) => {
      resourceCatalogPromise = undefined;
      throw error;
    });
  }
  return resourceCatalogPromise;
}

async function streamResourceCsv(
  resource: CkanResource,
  onRecord: (record: Record<string, string>) => void,
): Promise<number> {
  const response = await fetchWithRetry(resource.url);
  if (!response.ok) throw new Error(`No se pudo descargar ${resource.name} (HTTP ${response.status}).`);

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

async function loadAreaProductionResource(
  resource: CkanResource,
  areaId: string,
  onRecord: (record: Record<string, string>) => void,
): Promise<number> {
  const filteredUrl = filteredProductionProxyUrl(resource.url, areaId);
  if (!filteredUrl) return streamResourceCsv(resource, onRecord);

  const response = await fetchWithRetry(filteredUrl);
  if (!response.ok) {
    const detail = (await response.text()).trim();
    throw new Error(
      `No se pudo descargar ${resource.name} (HTTP ${response.status})${detail ? `: ${detail}` : '.'}`,
    );
  }
  const payload = await response.json();
  if (!payload || !Number.isFinite(payload.scannedRows) || !Array.isArray(payload.records)) {
    throw new Error(`El servicio de datos devolvió una respuesta inválida para ${resource.name}.`);
  }
  for (const record of payload.records) {
    if (record && typeof record === 'object') onRecord(record as Record<string, string>);
  }
  return payload.scannedRows as number;
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

export async function fetchAreaCatalog(onProgress?: (progress: CatalogLoadProgress) => void): Promise<AreaCatalogItem[]> {
  onProgress?.({ step: 1, message: 'Conectando con el servicio seguro de CapIV' });
  const optimizedCatalogUrl = catalogProxyUrl();
  if (optimizedCatalogUrl) {
    const response = await fetchWithRetry(optimizedCatalogUrl);
    if (!response.ok) throw new Error(`No se pudo consultar el catálogo optimizado (HTTP ${response.status}).`);
    onProgress?.({ step: 2, message: 'Respuesta recibida; leyendo la lista de áreas' });
    const catalog = await response.json();
    if (!Array.isArray(catalog)) throw new Error('El catálogo optimizado devolvió una respuesta inválida.');
    onProgress?.({ step: 3, message: 'Validando nombres, provincias y empresas' });
    return catalog as AreaCatalogItem[];
  }
  onProgress?.({ step: 1, message: 'Consultando los recursos oficiales disponibles' });
  const resources = await loadResourceCatalog();
  const byArea = new Map<string, AreaCatalogItem>();
  onProgress?.({ step: 2, message: 'Leyendo el padrón oficial de pozos y áreas' });
  await streamResourceCsv(resources.wells, (record) => {
    const areaId = text(record, 'cod_area', 'idareapermisoconcesion');
    const areaName = text(record, 'area', 'areapermisoconcesion');
    if (!areaId || !areaName) return;

    const existing = byArea.get(areaId);
    const company = text(record, 'empresa');
    if (existing) {
      if (company && !existing.companies.includes(company)) existing.companies.push(company);
      return;
    }

    byArea.set(areaId, {
      province: text(record, 'provincia') || 'Sin provincia',
      areaId,
      areaName,
      basin: text(record, 'cuenca'),
      companies: company ? [company] : [],
    });
  });

  onProgress?.({ step: 3, message: 'Ordenando áreas, provincias y empresas' });
  return [...byArea.values()].sort((a, b) =>
    `${a.province}|${a.areaName}`.localeCompare(`${b.province}|${b.areaName}`, 'es'),
  );
}

export function productionSourceForArea(areaId: string, year: number): AreaProductionSource {
  if (year === 2021 && areaId === 'EPMD') {
    return {
      sourceAreaId: 'EPN',
      siglaPattern: /\.Md/i,
      note: 'El recurso oficial 2021 usa el código legado EPN; se recupera la porción Mendoza por sigla de pozo.',
    };
  }
  if (year === 2021 && areaId === 'EPNQ') {
    return {
      sourceAreaId: 'EPN',
      siglaPattern: /\.Nq\./i,
      note: 'El recurso oficial 2021 usa el código legado EPN; se recupera la porción Neuquén por sigla de pozo.',
    };
  }
  return { sourceAreaId: areaId };
}

function normalizeProductionRecord(
  record: Record<string, string>,
  areaId: string,
  areaName: string,
  source: AreaProductionSource,
): ProductionRecord | null {
  const recordAreaId = text(record, 'idareapermisoconcesion', 'cod_area');
  if (recordAreaId !== source.sourceAreaId) return null;
  const wellName = text(record, 'sigla');
  if (source.siglaPattern && !source.siglaPattern.test(wellName)) return null;

  const year = numberValue(record, 'anio');
  const month = numberValue(record, 'mes');
  if (!Number.isInteger(year) || !Number.isInteger(month)) return null;

  return {
    areaId,
    areaName,
    wellId: text(record, 'idpozo'),
    wellName,
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
  const resources = await loadResourceCatalog();

  for (let year = startYear; year <= currentYear; year++) {
    const resource = resources.productionByYear[year];
    if (!resource) {
      onStep?.(`Sin recurso anual publicado para ${year}; se continúa con el siguiente año`);
      continue;
    }
    const source = productionSourceForArea(area.areaId, year);
    await delay(DOWNLOAD_PAUSE_MS);
    if (source.note) onStep?.(`${area.areaId} ${year}: ${source.note}`);
    onStep?.(`Descargando Capítulo IV ${area.areaId} ${year}`);
    await onEvent?.({ type: 'resource_started', areaId: area.areaId, source: 'capitulo-iv', year });
    let matched = 0;
    const matchedRows: ProductionRecord[] = [];
    const rows = await loadAreaProductionResource(resource, source.sourceAreaId, (row) => {
      const normalized = normalizeProductionRecord(row, area.areaId, area.areaName, source);
      if (!normalized) return;
      const key = `${normalized.wellName}|${normalized.wellId}|${normalized.year}|${normalized.month}`;
      if (seen.has(key)) return;
      seen.add(key);
      records.push(normalized);
      matchedRows.push(normalized);
      matched++;
    });
    onStep?.(`Descargado Capítulo IV ${area.areaId} ${year}: ${matched} de ${rows} filas`);
    await onEvent?.({
      type: 'resource_completed',
      areaId: area.areaId,
      source: 'capitulo-iv',
      year,
      scannedRows: rows,
      matchedRows: matched,
      records: matchedRows,
    });
  }

  const sorted = records.sort((a, b) => a.year - b.year || a.month - b.month || a.wellName.localeCompare(b.wellName));
  await onEvent?.({ type: 'completed', areaId: area.areaId, records: sorted });
  return sorted;
}

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

export function countProductionResources(startYear: number): number {
  const currentYear = new Date().getFullYear();
  return Math.max(0, currentYear - Math.max(2006, startYear) + 1);
}

function newestResource(resources: CkanResource[]): CkanResource | undefined {
  return [...resources].sort((a, b) => (b.last_modified ?? '').localeCompare(a.last_modified ?? ''))[0];
}

function normalize(value: string): string {
  return value
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[–—]/g, '-')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
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
