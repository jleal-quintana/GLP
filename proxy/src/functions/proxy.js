const { app } = require('@azure/functions');
const { parse } = require('csv-parse');
const { Readable } = require('node:stream');

const ALLOWED_HOSTS = new Set(['datos.gob.ar', 'datos.energia.gob.ar']);
const ALLOWED_ORIGINS = new Set([
  'https://jleal-quintana.github.io',
  'https://localhost:3002',
]);
const DATASET_URL = 'https://datos.gob.ar/api/3/action/package_show?id=produccion-de-petroleo-y-gas-por-pozo';
let catalogCache;

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

    const rawUrl = request.query.get('url');
    if (!rawUrl) return { status: 400, headers: corsHeaders, body: 'Falta el parámetro url.' };

    let target;
    try {
      target = validateTarget(rawUrl);
    } catch (error) {
      return { status: 400, headers: corsHeaders, body: error.message };
    }

    try {
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

async function buildAreaCatalog() {
  if (catalogCache && catalogCache.expiresAt > Date.now()) return catalogCache.value;
  const datasetResponse = await fetchAllowed(validateTarget(DATASET_URL));
  if (!datasetResponse.ok) throw new Error(`catálogo oficial HTTP ${datasetResponse.status}`);
  const dataset = await datasetResponse.json();
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
