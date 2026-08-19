export const USER_AGENT =
  'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36';

declare const CAPIV_PROXY_BASE_URL: string;

export async function fetchWithRetry(url: string, attempts = 4): Promise<Response> {
  const requestUrl = proxiedUrl(url);
  let lastError: unknown;
  for (let attempt = 1; attempt <= attempts; attempt++) {
    try {
      const response = await fetch(requestUrl);
      if (response.ok || (response.status < 500 && response.status !== 429)) {
        return response;
      }
      lastError = new Error(`HTTP ${response.status}`);
    } catch (error) {
      lastError = error;
    }
    await new Promise((resolve) => setTimeout(resolve, 500 * attempt));
  }
  throw lastError instanceof Error ? lastError : new Error(String(lastError));
}

function proxiedUrl(url: string): string {
  if (typeof window === 'undefined') return url;
  const host = window.location.hostname;
  const target = new URL(url);
  if (target.hostname !== 'datos.gob.ar' && target.hostname !== 'datos.energia.gob.ar') return url;
  if (host === 'localhost' || host === '127.0.0.1') {
    return `/capiv-proxy?url=${encodeURIComponent(url)}`;
  }
  if (target.hostname === 'datos.gob.ar') return url;
  const proxyBase = typeof CAPIV_PROXY_BASE_URL === 'string' ? CAPIV_PROXY_BASE_URL.trim() : '';
  if (proxyBase) {
    return proxyBase.includes('{url}')
      ? proxyBase.replace('{url}', encodeURIComponent(url))
      : `${proxyBase}${proxyBase.includes('?') ? '&' : '?'}url=${encodeURIComponent(url)}`;
  }
  throw new Error(
    'La descarga de producción requiere el servicio HTTPS de CapIV. Contactá al equipo de soporte.',
  );
}
