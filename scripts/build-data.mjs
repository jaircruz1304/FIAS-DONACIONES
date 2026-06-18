import fs from 'node:fs';
import path from 'node:path';
import * as XLSX from 'xlsx';

const EXCEL_URL = 'https://fiasec-my.sharepoint.com/:x:/g/personal/jcruzg_fias_org_ec/IQAWjBvkE6TPTG6b5vhBTZxcAaAldunzyoWrEzM5q5mVINo?e=YRRr8R&download=1';
const OUT_DIR = path.resolve('data');
const OUT_JSON = path.join(OUT_DIR, 'control-donaciones.json');
const OUT_META = path.join(OUT_DIR, 'control-donaciones-meta.json');

function normalizeKey(value) {
  return String(value ?? '')
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}

function getFlexible(row, ...names) {
  const indexed = {};
  Object.keys(row || {}).forEach(key => {
    indexed[normalizeKey(key)] = row[key];
  });
  for (const name of names) {
    const key = normalizeKey(name);
    if (Object.prototype.hasOwnProperty.call(indexed, key)) return indexed[key];
  }
  return '';
}

function excelDateToISO(value) {
  if (!value) return '';
  if (typeof value === 'number') {
    const d = new Date(Date.UTC(1899, 11, 30));
    d.setUTCDate(d.getUTCDate() + value);
    return d.toISOString().slice(0, 10);
  }
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  return String(value).slice(0, 10);
}

function toNumber(value) {
  if (value === null || value === undefined || value === '') return 0;
  if (typeof value === 'number') return value;
  const cleaned = String(value).replace(/\$/g, '').replace(/\s/g, '').replace(/\./g, '').replace(',', '.');
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : 0;
}

function normalizeImported(row) {
  const avanceRaw = getFlexible(row, 'avance', 'Avance %', 'Avance');
  let avance = 0;
  if (typeof avanceRaw === 'string' && avanceRaw.includes('%')) {
    avance = parseFloat(avanceRaw.replace(',', '.')) / 100;
  } else {
    avance = Number(avanceRaw || 0);
    if (avance > 1) avance = avance / 100;
  }

  const respuestaRaw = getFlexible(row, 'respuesta', 'Respuesta del beneficiario', 'Respuesta');
  const respuesta = typeof respuestaRaw === 'number' ? excelDateToISO(respuestaRaw) : respuestaRaw;

  const diasRaw = getFlexible(row, 'diasGestion', 'Días en gestión', 'Dias en gestión', 'Días', 'Dias');
  const diasGestion = diasRaw === null || diasRaw === undefined || diasRaw === '' ? null : Number(diasRaw);

  return {
    nro: getFlexible(row, 'nro', 'Nro.', 'Nro', 'No', 'Número', 'Numero'),
    programa: getFlexible(row, 'programa', 'Fondo / Programa / Proyecto', 'Programa', 'Proyecto'),
    beneficiario: getFlexible(row, 'beneficiario', 'Beneficiario'),
    descripcion: getFlexible(row, 'descripcion', 'Descripción de bienes a donar', 'Descripcion de bienes a donar', 'Descripcion', 'Descripción'),
    valor: toNumber(getFlexible(row, 'valor', 'Valor estimado (USD)', 'Valor', 'Monto')),
    fechaIntencion: excelDateToISO(getFlexible(row, 'fechaIntencion', 'Fecha de intención', 'Fecha intención', 'Fecha intencion')),
    respuesta: respuesta || '',
    fechaActa: excelDateToISO(getFlexible(row, 'fechaActa', 'Fecha de acta entrega-recepción', 'Fecha de acta entrega recepcion', 'Fecha de acta', 'Fecha acta')),
    estado: getFlexible(row, 'estado', 'Estado actual', 'Estado'),
    avance: Number.isFinite(avance) ? avance : 0,
    proximaAccion: getFlexible(row, 'proximaAccion', 'Próxima acción requerida', 'Proxima accion requerida', 'Proxima acción', 'Próxima acción'),
    responsable: getFlexible(row, 'responsable', 'Responsable / área', 'Responsable / area', 'Responsable', 'Área', 'Area'),
    expediente: getFlexible(row, 'expediente', 'Expediente digital', 'Expediente'),
    detalle: getFlexible(row, 'detalle', 'Detalle del trámite', 'Detalle del tramite', 'Detalle'),
    observaciones: getFlexible(row, 'observaciones', 'Observaciones / alertas', 'Observaciones', 'Alertas'),
    diasGestion: Number.isFinite(diasGestion) ? diasGestion : null,
    semaforo: getFlexible(row, 'semaforo', 'Semáforo', 'Semaforo')
  };
}

function extractRowsFromWorkbook(workbook) {
  let selected = null;

  for (const sheetName of workbook.SheetNames) {
    const matrix = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
      header: 1,
      raw: true,
      defval: null
    });

    const headerIndex = matrix.findIndex(row =>
      Array.isArray(row) && row.some(cell => ['nro', 'nro.'].includes(String(cell || '').trim().toLowerCase()))
    );

    if (headerIndex >= 0) {
      selected = { sheetName, matrix, headerIndex };
      break;
    }
  }

  if (!selected) {
    throw new Error('No se encontró la fila de encabezados con Nro.');
  }

  const headers = selected.matrix[selected.headerIndex].map(h => String(h || '').trim());
  const rows = selected.matrix
    .slice(selected.headerIndex + 1)
    .filter(row => Array.isArray(row) && row.some(Boolean))
    .filter(row => row[0] !== null && String(row[0]).toLowerCase() !== 'total general')
    .map(row => {
      const obj = {};
      headers.forEach((header, index) => obj[header] = row[index]);
      return normalizeImported(obj);
    })
    .filter(row => row.nro);

  if (!rows.length) {
    throw new Error('El Excel fue leído, pero no se encontraron registros válidos.');
  }

  return rows;
}

async function main() {
  console.log('Descargando Excel público desde OneDrive...');
  const response = await fetch(EXCEL_URL, { redirect: 'follow' });

  if (!response.ok) {
    throw new Error(`No se pudo descargar el Excel. HTTP ${response.status} ${response.statusText}`);
  }

  const contentType = response.headers.get('content-type') || '';
  const buffer = Buffer.from(await response.arrayBuffer());

  if (buffer.byteLength < 100) {
    throw new Error(`Archivo descargado demasiado pequeño: ${buffer.byteLength} bytes. Content-Type: ${contentType}`);
  }

  const workbook = XLSX.read(buffer, { type: 'buffer' });
  const rows = extractRowsFromWorkbook(workbook);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  fs.writeFileSync(OUT_JSON, JSON.stringify(rows, null, 2), 'utf8');
  fs.writeFileSync(OUT_META, JSON.stringify({
    fuente: EXCEL_URL,
    registros: rows.length,
    generado: new Date().toISOString(),
    contentType,
    bytes: buffer.byteLength
  }, null, 2), 'utf8');

  console.log(`JSON generado correctamente: ${OUT_JSON} (${rows.length} registros).`);
}

main().catch(error => {
  console.error('Error generando data/control-donaciones.json');
  console.error(error);
  process.exit(1);
});
