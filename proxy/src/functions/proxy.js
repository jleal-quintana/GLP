const { app } = require('@azure/functions');
const { BlobServiceClient } = require('@azure/storage-blob');
const { parse } = require('csv-parse');
const { Readable } = require('node:stream');

const ALLOWED_HOSTS = new Set(['datos.gob.ar', 'datos.energia.gob.ar']);
const ALLOWED_ORIGINS = new Set([
  'https://jleal-quintana.github.io',
  'https://localhost:3002',
]);
const DATASET_URL = 'https://datos.gob.ar/api/3/action/package_show?id=produccion-de-petroleo-y-gas-por-pozo';
const PRODUCTION_CONTAINER = 'production-cache';
const CURRENT_YEAR_CACHE_MS = 6 * 60 * 60 * 1000;
let catalogCache;
let datasetCache;
const productionCache = new Map();
let productionContainerPromise;

// Annual production files can exceed 200 MB. Azure Functions must explicitly
// enable HTTP streaming or the runtime buffers/cuts large upstream responses.
app.setup({ enableHttpStream: true });

app.http('proxy', {
  methods: ['GET', 'OPTIONS'],
  authLevel: 'anonymous',
  route: 'proxy',
  handler: async (request) => {
    const origin = request.headers.get('origin') || '';
    const corsOrigin = ALLOWED_ORIGINS.has(origin) ? origin : 'https://jleal-quintana.github.io';
    const corsHeaders = {
      'access-control-allow-origin': corsOrigin,
      'access-control-allow-methods': 'GET, OPTIONS',
      'access-control-allow-headers': 'Content-Type',
      'vary': 'Origin',
    };

    if (request.method === 'OPTIONS') return { status: 204, headers: corsHeaders };

    if (request.query.get('catalog') === '1') {
      try {
        const catalog = await buildAreaCatalog();
        return {
          status: 200,
          headers: { ...corsHeaders, 'content-type': 'application/json; charset=utf-8', 'cache-control': 'public, max-age=900' },
          jsonBody: catalog,
        };
      } catch (error) {
        return { status: 502, headers: corsHeaders, body: `No se pudo preparar el catálogo: ${error.message}` };
      }
    }

    if (request.query.get('resources') === '1') {
      try {
        const dataset = await loadDataset();
        return {
          status: 200,
          headers: { ...corsHeaders, 'content-type': 'application/json; charset=utf-8', 'cache-control': 'public, max-age=900' },
          jsonBody: dataset,
        };
      } catch (error) {
        return { status: 502, headers: corsHeaders, body: `No se pudo consultar el índice oficial: ${error.message}` };
      }
    }

    const rawUrl = request.query.get('url');
    if (!rawUrl) return { status: 400, headers: corsHeaders, body: 'Falta el parámetro url.' };

    let target;
    try {
      target = validateTarget(rawUrl);
    } catch (error) {
      return { status: 400, headers: corsHeaders, body: error.message };
    }

    try {
      const areaId = request.query.get('areaId');
      if (areaId) {
        const filtered = await filterProductionByArea(target, validateAreaId(areaId));
        return {
          status: 200,
          headers: { ...corsHeaders, 'content-type': 'application/json; charset=utf-8', 'cache-control': 'public, max-age=900' },
          jsonBody: filtered,
        };
      }

      const upstream = await fetchAllowed(target);
      const headers = {
        ...corsHeaders,
        'content-type': upstream.headers.get('content-type') || 'application/octet-stream',
        'cache-control': upstream.headers.get('cache-control') || 'public, max-age=900',
      };
      const disposition = upstream.headers.get('content-disposition');
      const lastModified = upstream.headers.get('last-modified');
      const etag = upstream.headers.get('etag');
      if (disposition) headers['content-disposition'] = disposition;
      if (lastModified) headers['last-modified'] = lastModified;
      if (etag) headers.etag = etag;
      return { status: upstream.status, headers, body: upstream.body };
    } catch (error) {
      return { status: 502, headers: corsHeaders, body: `No se pudo consultar la fuente oficial: ${error.message}` };
    }
  },
});

async function filterProductionByArea(target, areaId) {
  const cacheKey = `${target.toString()}|${areaId}`;
  const cached = productionCache.get(cacheKey);
  if (cached?.expiresAt > Date.now()) return cached.value;
  if (cached) productionCache.delete(cacheKey);

  const persisted = await loadPreFilteredProduction(target, areaId);
  if (persisted) {
    productionCache.set(cacheKey, { value: persisted, expiresAt: Date.now() + 15 * 60 * 1000 });
    return persisted;
  }

  const source = await getProductionSource(target);

  let scannedRows = 0;
  const records = [];
  const parser = source.pipe(parse({
    bom: true,
    columns: (headers) => headers.map((header) => String(header).trim().toLowerCase()),
    skip_empty_lines: true,
    relax_column_count: true,
  }));

  for await (const row of parser) {
    scannedRows++;
    if (textValue(row.idareapermisoconcesion, row.cod_area) !== areaId) continue;
    records.push({
      idareapermisoconcesion: areaId,
      idpozo: textValue(row.idpozo),
      sigla: textValue(row.sigla),
      anio: textValue(row.anio),
      mes: textValue(row.mes),
      prod_pet: textValue(row.prod_pet, row.prod_petroleo, row.petroleo),
      prod_gas: textValue(row.prod_gas, row.gas),
      prod_agua: textValue(row.prod_agua, row.agua),
      iny_agua: textValue(row.iny_agua, row.agua_iny, row.inyeccion_agua),
    });
  }

  const value = { scannedRows, records };
  await persistFilteredProduction(target, areaId, records);
  productionCache.set(cacheKey, { value, expiresAt: Date.now() + 15 * 60 * 1000 });
  if (productionCache.size > 40) {
    const oldestKey = productionCache.keys().next().value;
    if (oldestKey) productionCache.delete(oldestKey);
  }
  return value;
}

async function loadPreFilteredProduction(target, areaId) {
  const year = productionYear(target);
  if (!year || !process.env.AzureWebJobsStorage) return undefined;
  try {
    const container = await getProductionContainer();
    const recordsBlob = container.getBlobClient(`filtered/${year}/${encodeURIComponent(areaId)}.ndjson`);
    const properties = await recordsBlob.getProperties();
    if (year === new Date().getUTCFullYear()
      && Date.now() - properties.lastModified.getTime() >= CURRENT_YEAR_CACHE_MS) return undefined;
    const recordsBuffer = await recordsBlob.downloadToBuffer();
    let scannedRows;
    try {
      const markerBuffer = await container.getBlobClient(`filtered/${year}/_complete.json`).downloadToBuffer();
      scannedRows = Number(JSON.parse(markerBuffer.toString('utf8')).scannedRows);
    } catch (error) {
      if (error.statusCode !== 404) throw error;
    }
    const records = recordsBuffer.toString('utf8')
      .split(/\r?\n/)
      .filter(Boolean)
      .map((line) => JSON.parse(line));
    return { scannedRows: scannedRows || records.length, records };
  } catch (error) {
    if (error.statusCode !== 404) console.warn(`No se pudo leer el recorte ${year}/${areaId}: ${error.message}`);
    return undefined;
  }
}

async function persistFilteredProduction(target, areaId, records) {
  const year = productionYear(target);
  if (!year || !process.env.AzureWebJobsStorage) return;
  try {
    const container = await getProductionContainer();
    const body = records.map((record) => JSON.stringify(record)).join('\n') + (records.length ? '\n' : '');
    await container
      .getBlockBlobClient(`filtered/${year}/${encodeURIComponent(areaId)}.ndjson`)
      .uploadData(Buffer.from(body), { blobHTTPHeaders: { blobContentType: 'application/x-ndjson; charset=utf-8' } });
  } catch (error) {
    console.warn(`No se pudo guardar el recorte ${year}/${areaId}: ${error.message}`);
  }
}

async function getProductionSource(target) {
  const year = productionYear(target);
  if (year && process.env.AzureWebJobsStorage) {
    try {
      const container = await getProductionContainer();
      const blob = container.getBlobClient(`${year}.csv`);
      let properties;
      try {
        properties = await blob.getProperties();
      } catch (error) {
        if (error.statusCode !== 404) throw error;
      }

      const currentYear = new Date().getUTCFullYear();
      const isFresh = properties?.contentLength > 0 && properties.copyStatus !== 'pending'
        && (year < currentYear || Date.now() - properties.lastModified.getTime() < CURRENT_YEAR_CACHE_MS);
      if (!isFresh) {
        const poller = await blob.beginCopyFromURL(target.toString(), { intervalInMs: 2000 });
        await poller.pollUntilDone();
      }

      const download = await blob.download();
      if (download.readableStreamBody) return download.readableStreamBody;
    } catch (error) {
      // Keep the official source as a fallback while the cache is being seeded.
      console.warn(`No se pudo usar el cache ${year}: ${error.message}`);
    }
  }

  const response = await fetchAllowed(target);
  if (!response.ok || !response.body) throw new Error(`recurso anual HTTP ${response.status}`);
  return Readable.fromWeb(response.body);
}

async function getProductionContainer() {
  if (!productionContainerPromise) {
    productionContainerPromise = (async () => {
      const service = BlobServiceClient.fromConnectionString(process.env.AzureWebJobsStorage);
      const container = service.getContainerClient(PRODUCTION_CONTAINER);
      await container.createIfNotExists();
      return container;
    })().catch((error) => {
      productionContainerPromise = undefined;
      throw error;
    });
  }
  return productionContainerPromise;
}

function productionYear(target) {
  const match = decodeURIComponent(target.pathname).match(/(?:19|20)\d{2}/g);
  return match ? Number(match[match.length - 1]) : undefined;
}

async function buildAreaCatalog() {
  if (catalogCache && catalogCache.expiresAt > Date.now()) return catalogCache.value;
  const dataset = await loadDataset();
  const wells = (dataset.result?.resources || [])
    .filter((resource) => String(resource.format).toUpperCase() === 'CSV' && normalize(resource.name) === 'capitulo iv pozos')
    .sort((a, b) => String(b.last_modified || '').localeCompare(String(a.last_modified || '')))[0];
  if (!wells?.url) throw new Error('no se encontró el recurso Capítulo IV - Pozos');

  const csvResponse = await fetchAllowed(validateTarget(wells.url));
  if (!csvResponse.ok || !csvResponse.body) throw new Error(`recurso de pozos HTTP ${csvResponse.status}`);
  const byArea = new Map();
  const parser = Readable.fromWeb(csvResponse.body).pipe(parse({
    bom: true,
    columns: (headers) => headers.map((header) => String(header).trim().toLowerCase()),
    skip_empty_lines: true,
    relax_column_count: true,
  }));
  for await (const row of parser) {
    const areaId = textValue(row.cod_area, row.idareapermisoconcesion);
    const areaName = textValue(row.area, row.areapermisoconcesion);
    if (!areaId || !areaName) continue;
    const company = textValue(row.empresa);
    const existing = byArea.get(areaId);
    if (existing) {
      if (company) existing.companies.add(company);
    } else {
      byArea.set(areaId, {
        province: textValue(row.provincia) || 'Sin provincia',
        areaId,
        areaName,
        basin: textValue(row.cuenca),
        companies: new Set(company ? [company] : []),
      });
    }
  }
  const value = [...byArea.values()]
    .map((area) => ({ ...area, companies: [...area.companies] }))
    .sort((a, b) => `${a.province}|${a.areaName}`.localeCompare(`${b.province}|${b.areaName}`, 'es'));
  catalogCache = { value, expiresAt: Date.now() + 15 * 60 * 1000 };
  return value;
}

async function loadDataset() {
  if (datasetCache?.expiresAt > Date.now()) return datasetCache.value;
  const response = await fetchAllowed(validateTarget(DATASET_URL));
  if (!response.ok) throw new Error(`catálogo oficial HTTP ${response.status}`);
  const value = await response.json();
  if (!value?.success || !Array.isArray(value.result?.resources)) {
    throw new Error('el catálogo oficial devolvió una respuesta inválida');
  }
  datasetCache = { value, expiresAt: Date.now() + 15 * 60 * 1000 };
  return value;
}

function textValue(...values) {
  for (const value of values) if (value !== undefined && value !== null && String(value).trim()) return String(value).trim();
  return '';
}

function normalize(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function validateAreaId(areaId) {
  const value = String(areaId).trim();
  if (!value || value.length > 100 || /[\u0000-\u001f]/.test(value)) {
    throw new Error('El identificador de área no es válido.');
  }
  return value;
}

function validateTarget(rawUrl) {
  const target = new URL(rawUrl);
  if (!['http:', 'https:'].includes(target.protocol) || !ALLOWED_HOSTS.has(target.hostname)) {
    throw new Error('El destino solicitado no está permitido.');
  }
  target.username = '';
  target.password = '';
  return target;
}

async function fetchAllowed(initialTarget) {
  let target = initialTarget;
  for (let redirect = 0; redirect <= 5; redirect++) {
    const response = await fetch(target, {
      redirect: 'manual',
      headers: { 'user-agent': 'Quintana-CapIV-Addin/1.0' },
    });
    if (![301, 302, 303, 307, 308].includes(response.status)) return response;
    const location = response.headers.get('location');
    if (!location) return response;
    target = validateTarget(new URL(location, target).toString());
  }
  throw new Error('La fuente oficial devolvió demasiadas redirecciones.');
}
