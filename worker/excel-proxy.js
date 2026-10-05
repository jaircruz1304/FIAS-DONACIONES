/**
 * Cloudflare Worker opcional para el Dashboard de Donaciones FIAS.
 * Su única función es descargar el Excel público de SharePoint desde servidor
 * y entregarlo al navegador con CORS habilitado. No almacena datos ni secretos.
 */
const SOURCE_URL = 'https://fiasec-my.sharepoint.com/personal/jcruzg_fias_org_ec/_layouts/15/download.aspx?share=IQAUjBvkE6TPT7eb5vhBTZxcAaAltunzyoWrEzM5q5mVINo';

export default {
  async fetch(request) {
    const cors = {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type',
      'Access-Control-Expose-Headers': 'ETag, Last-Modified, Content-Length, Content-Type',
      'Cache-Control': 'no-store, no-cache, must-revalidate, max-age=0'
    };

    if (request.method === 'OPTIONS') return new Response(null, { status: 204, headers: cors });
    if (request.method !== 'GET') return new Response('Método no permitido', { status: 405, headers: cors });

    const upstream = await fetch(`${SOURCE_URL}&_=${Date.now()}`, {
      redirect: 'follow',
      headers: {
        'Accept': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream,*/*',
        'User-Agent': 'FIAS-Donaciones-Excel-Proxy'
      },
      cf: { cacheTtl: 0, cacheEverything: false }
    });

    if (!upstream.ok) {
      return new Response(`No se pudo obtener el Excel. HTTP ${upstream.status}`, { status: 502, headers: cors });
    }

    const body = await upstream.arrayBuffer();
    const bytes = new Uint8Array(body);
    if (bytes.length < 4 || bytes[0] !== 0x50 || bytes[1] !== 0x4b) {
      return new Response('SharePoint no devolvió un XLSX válido.', { status: 502, headers: cors });
    }

    const headers = new Headers(cors);
    headers.set('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    headers.set('Content-Length', String(body.byteLength));
    const etag = upstream.headers.get('etag');
    const lastModified = upstream.headers.get('last-modified');
    if (etag) headers.set('ETag', etag);
    if (lastModified) headers.set('Last-Modified', lastModified);

    return new Response(body, { status: 200, headers });
  }
};
