const { app } = require('@azure/functions');

const ALLOWED_HOSTS = new Set(['datos.gob.ar', 'datos.energia.gob.ar']);
const ALLOWED_ORIGINS = new Set([
  'https://jleal-quintana.github.io',
  'https://localhost:3002',
]);

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
