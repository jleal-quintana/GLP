export const USER_AGENT =
  'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36';

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
  if (host !== 'localhost' && host !== '127.0.0.1') return url;
  const target = new URL(url);
  if (target.hostname !== 'datos.gob.ar' && target.hostname !== 'datos.energia.gob.ar') return url;
  return `/capiv-proxy?url=${encodeURIComponent(url)}`;
}
