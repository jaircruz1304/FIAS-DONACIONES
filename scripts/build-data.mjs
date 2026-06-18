import fs from 'node:fs/promises';
import path from 'node:path';
import process from 'node:process';
import * as XLSX from 'xlsx';

const OUT_DIR = 'data';
const OUT_JSON = path.join(OUT_DIR, 'control-donaciones.json');
const OUT_META = path.join(OUT_DIR, 'control-donaciones-meta.json');

const SHARE_ID = 'IQAWjBvkE6TPTG6b5vhBTZxcAaAldunzyoWrEzM5q5mVINo';

const DOWNLOAD_CANDIDATES = [
  {
    name: 'SharePoint download.aspx',
    url: `https://fiasec-my.sharepoint.com/personal/jcruzg_fias_org_ec/_layouts/15/download.aspx?share=${SHARE_ID}`
  },
  {
    name: 'Share link con download=1',
    url: 'https://fiasec-my.sharepoint.com/:x:/g/personal/jcruzg_fias_org_ec/IQAWjBvkE6TPTG6b5vhBTZxcAaAldunzyoWrEzM5q5mVINo?e=YRRr8R&download=1'
  },
  {
    name: 'Share link con download=1 y web=0',
    url: 'https://fiasec-my.sharepoint.com/:x:/g/personal/jcruzg_fias_org_ec/IQAWjBvkE6TPTG6b5vhBTZxcAaAldunzyoWrEzM5q5mVINo?e=YRRr8R&download=1&web=0'
  }
];

const EXPECTED_HEADERS = [
  'Nro.',
  'Fondo / Programa / Proyecto',
  'Beneficiario',
  'Descripción de bienes a donar',
  'Valor estimado (USD)',
  'Fecha de intención',
  'Respuesta del beneficiario',
  'Fecha de acta entrega-recepción',
  'Estado actual',
  'Avance %',
  'Próxima acción requerida',
  'Responsable / área',
  'Expediente digital',
  'Detalle del trámite',
  'Observaciones / alertas',
  'Días en gestión',
  'Semáforo'
];

function normalizeText(value) {
  return String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toLowerCase();
}

function normalizeKey(value) {
  return normalizeText(value)
    .replace(/[%()$]/g, '')
    .replace(/[^a-z0-9]+/g, '');
}

function parseNumber(value) {
  if (value === null || value === undefined || value === '') return 0;

  if (typeof value === 'number') return value;

  let s = String(value).trim();

  s = s.replace(/[^\d,.-]/g, '');

  if (s.includes(',') && s.includes('.')) {
    s = s.replace(/\./g, '').replace(',', '.');
  } else if (s.includes(',')) {
    s = s.replace(',', '.');
  }

  const n = Number(s);

  return Number.isFinite(n) ? n : 0;
}

function parseProgress(value) {
  if (value === null || value === undefined || value === '') return 0;

  if (typeof value === 'number') {
    return value > 1 ? value / 100 : value;
  }

  const raw = String(value).trim();
  const n = parseNumber(raw);

  if (raw.includes('%')) return n / 100;

  return n > 1 ? n / 100 : n;
}

function excelDateToISO(value) {
  if (!value) return '';

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toISOString().slice(0, 10);
  }

  if (typeof value === 'number') {
    const parsed = XLSX.SSF.parse_date_code(value);

    if (!parsed) return '';

    const yyyy = String(parsed.y).padStart(4, '0');
    const mm = String(parsed.m).padStart(2, '0');
    const dd = String(parsed.d).padStart(2, '0');

    return `${yyyy}-${mm}-${dd}`;
  }

  const text = String(value).trim();

  if (!text) return '';

  const iso = text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);

  if (iso) {
    const [, y, m, d] = iso;
    return `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
  }

  const latam = text.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);

  if (latam) {
    const [, d, m, y] = latam;
    return `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
  }

  return text;
}

const FIELD_ALIASES = {
  nro: [
    'nro',
    'nro.',
    'no',
    'numero',
    'número',
    'id'
  ],

  programa: [
    'fondo / programa / proyecto',
    'fondo programa proyecto',
    'programa',
    'proyecto',
    'fondo'
  ],

  beneficiario: [
    'beneficiario',
    'beneficiario final'
  ],

  descripcion: [
    'descripción de bienes a donar',
    'descripcion de bienes a donar',
    'descripcion',
    'descripción',
    'bienes',
    'detalle de bienes'
  ],

  valor: [
    'valor estimado (usd)',
    'valor estimado usd',
    'valor',
    'monto',
    'monto usd'
  ],

  fechaIntencion: [
    'fecha de intención',
    'fecha de intencion',
    'fecha intención',
    'fecha intencion'
  ],

  respuesta: [
    'respuesta del beneficiario',
    'respuesta',
    'fecha respuesta',
    'respuesta beneficiario'
  ],

  fechaActa: [
    'fecha de acta entrega-recepción',
    'fecha de acta entrega recepcion',
    'fecha acta',
    'fecha de acta',
    'fecha acta entrega recepción',
    'fecha acta entrega recepcion'
  ],

  estado: [
    'estado actual',
    'estado',
    'situacion',
    'situación'
  ],

  avance: [
    'avance %',
    'avance',
    'porcentaje de avance',
    'avance porcentaje'
  ],

  proximaAccion: [
    'próxima acción requerida',
    'proxima accion requerida',
    'próxima acción',
    'proxima accion',
    'accion requerida',
    'acción requerida'
  ],

  responsable: [
    'responsable / área',
    'responsable area',
    'responsable / area',
    'responsable',
    'área',
    'area'
  ],

  expediente: [
    'expediente digital',
    'expediente',
    'link',
    'enlace',
    'url'
  ],

  detalle: [
    'detalle del trámite',
    'detalle del tramite',
    'detalle',
    'trámite',
    'tramite'
  ],

  observaciones: [
    'observaciones / alertas',
    'observaciones alertas',
    'observaciones',
    'alertas'
  ],

  diasGestion: [
    'días en gestión',
    'dias en gestion',
    'días',
    'dias',
    'dias gestion'
  ],

  semaforo: [
    'semáforo',
    'semaforo',
    'alerta',
    'semaforo alerta'
  ]
};

const NORMALIZED_ALIASES = Object.fromEntries(
  Object.entries(FIELD_ALIASES).map(([field, aliases]) => [
    field,
    aliases.map(normalizeKey)
  ])
);

function valueByAliases(row, field) {
  const aliases = NORMALIZED_ALIASES[field] || [];

  for (const [key, value] of Object.entries(row)) {
    if (aliases.includes(normalizeKey(key))) {
      return value;
    }
  }

  return '';
}

function normalizeImported(row) {
  return {
    nro: valueByAliases(row, 'nro'),

    programa: String(valueByAliases(row, 'programa') ?? '').trim(),

    beneficiario: String(valueByAliases(row, 'beneficiario') ?? '').trim(),

    descripcion: String(valueByAliases(row, 'descripcion') ?? '').trim(),

    valor: parseNumber(valueByAliases(row, 'valor')),

    fechaIntencion: excelDateToISO(valueByAliases(row, 'fechaIntencion')),

    respuesta: String(valueByAliases(row, 'respuesta') ?? '').trim(),

    fechaActa: excelDateToISO(valueByAliases(row, 'fechaActa')),

    estado: String(valueByAliases(row, 'estado') ?? '').trim(),

    avance: parseProgress(valueByAliases(row, 'avance')),

    proximaAccion: String(valueByAliases(row, 'proximaAccion') ?? '').trim(),

    responsable: String(valueByAliases(row, 'responsable') ?? '').trim(),

    expediente: String(valueByAliases(row, 'expediente') ?? '').trim(),

    detalle: String(valueByAliases(row, 'detalle') ?? '').trim(),

    observaciones: String(valueByAliases(row, 'observaciones') ?? '').trim(),

    diasGestion: parseNumber(valueByAliases(row, 'diasGestion')) || null,

    semaforo: String(valueByAliases(row, 'semaforo') ?? '').trim()
  };
}

function isProbablyHtml(buffer) {
  const sample = buffer
    .subarray(0, 700)
    .toString('utf8')
    .trim()
    .toLowerCase();

  return (
    sample.startsWith('<!doctype html') ||
    sample.startsWith('<html') ||
    sample.includes('<html') ||
    sample.includes('microsoft') && sample.includes('sharepoint')
  );
}

function isProbablyXlsx(buffer) {
  return buffer.length > 4 && buffer[0] === 0x50 && buffer[1] === 0x4b;
}

async function downloadExcel() {
  const errors = [];

  for (const candidate of DOWNLOAD_CANDIDATES) {
    console.log(`Intentando descarga: ${candidate.name}`);
    console.log(candidate.url);

    try {
      const response = await fetch(candidate.url, {
        method: 'GET',
        redirect: 'follow',
        headers: {
          'User-Agent': 'Mozilla/5.0 GitHubActions DashboardDonacionesFIAS',
          'Accept': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream,*/*'
        }
      });

      const contentType = response.headers.get('content-type') || 'sin content-type';

      console.log(`HTTP ${response.status} ${response.statusText}`);
      console.log(`Content-Type: ${contentType}`);

      const arrayBuffer = await response.arrayBuffer();
      const buffer = Buffer.from(arrayBuffer);

      console.log(`Bytes descargados: ${buffer.length}`);

      if (!response.ok) {
        const sample = buffer.subarray(0, 1000).toString('utf8');

        errors.push(
          `${candidate.name}: HTTP ${response.status}. Muestra: ${sample}`
        );

        continue;
      }

      if (buffer.length < 1000) {
        const sample = buffer.toString('utf8');

        errors.push(
          `${candidate.name}: respuesta demasiado pequeña. Muestra: ${sample}`
        );

        continue;
      }

      if (isProbablyHtml(buffer) || contentType.toLowerCase().includes('text/html')) {
        const sample = buffer.subarray(0, 1200).toString('utf8');

        console.log('La respuesta parece HTML, no XLSX. Primeros caracteres:');
        console.log(sample);

        errors.push(
          `${candidate.name}: SharePoint devolvió HTML, no XLSX.`
        );

        continue;
      }

      if (!isProbablyXlsx(buffer)) {
        const sample = buffer.subarray(0, 500).toString('utf8');

        console.log('La respuesta no parece XLSX ZIP. Primeros caracteres:');
        console.log(sample);

        errors.push(
          `${candidate.name}: la respuesta no parece un XLSX válido.`
        );

        continue;
      }

      console.log(`Descarga válida detectada desde: ${candidate.name}`);

      return {
        buffer,
        sourceName: candidate.name,
        sourceUrl: candidate.url,
        contentType,
        bytes: buffer.length
      };

    } catch (error) {
      errors.push(`${candidate.name}: ${error.message}`);
    }
  }

  throw new Error(
    'No se pudo descargar un XLSX válido desde OneDrive/SharePoint.\n\n' +
    'Diagnóstico:\n' +
    errors.map((e, i) => `${i + 1}. ${e}`).join('\n') +
    '\n\nEsto significa que el enlace público abre en navegador, pero GitHub Actions recibe una página HTML de SharePoint en lugar del archivo Excel. ' +
    'En ese caso, genera un enlace directo desde OneDrive con la opción de descarga o reemplaza el Excel por un JSON público generado manualmente.'
  );
}

function findHeaderRow(matrix) {
  return matrix.findIndex(row => {
    if (!Array.isArray(row)) return false;

    const normalizedCells = row.map(normalizeKey);

    const hasNro =
      normalizedCells.includes('nro') ||
      normalizedCells.includes('no') ||
      normalizedCells.includes('numero') ||
      normalizedCells.includes('id');

    const hasProgram = normalizedCells.some(c =>
      c.includes('programa') ||
      c.includes('proyecto') ||
      c.includes('fondo')
    );

    const hasBeneficiary = normalizedCells.some(c =>
      c.includes('beneficiario')
    );

    return hasNro && (hasProgram || hasBeneficiary);
  });
}

function rowsFromWorkbook(workbook) {
  let selected = null;

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];

    const matrix = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      raw: true,
      defval: null
    });

    const headerIndex = findHeaderRow(matrix);

    console.log(
      `Hoja "${sheetName}": filas ${matrix.length}; encabezado detectado en índice ${headerIndex}`
    );

    if (headerIndex >= 0) {
      selected = {
        sheetName,
        matrix,
        headerIndex
      };

      break;
    }
  }

  if (!selected) {
    throw new Error(
      'No se encontró la fila de encabezados. Se esperaba una fila con Nro. y Programa/Beneficiario.\n' +
      `Encabezados esperados: ${EXPECTED_HEADERS.join(' | ')}`
    );
  }

  const headers = selected.matrix[selected.headerIndex].map(h =>
    String(h ?? '').trim()
  );

  console.log(`Hoja seleccionada: ${selected.sheetName}`);
  console.log('Encabezados detectados:');
  console.log(headers);

  const rows = selected.matrix
    .slice(selected.headerIndex + 1)
    .filter(row =>
      Array.isArray(row) &&
      row.some(cell =>
        cell !== null &&
        cell !== undefined &&
        String(cell).trim() !== ''
      )
    )
    .filter(row =>
      String(row[0] ?? '').trim().toLowerCase() !== 'total general'
    )
    .map(row => {
      const obj = {};

      headers.forEach((header, index) => {
        if (header) {
          obj[header] = row[index];
        }
      });

      return normalizeImported(obj);
    })
    .filter(row =>
      row.nro !== null &&
      row.nro !== undefined &&
      String(row.nro).trim() !== ''
    );

  return {
    rows,
    sheetName: selected.sheetName,
    headerIndex: selected.headerIndex,
    headers
  };
}

async function main() {
  console.log('Descargando Excel público desde OneDrive/SharePoint...');

  await fs.mkdir(OUT_DIR, {
    recursive: true
  });

  const download = await downloadExcel();

  console.log('Leyendo libro XLSX...');

  const workbook = XLSX.read(download.buffer, {
    type: 'buffer',
    cellDates: true
  });

  console.log(`Hojas encontradas: ${workbook.SheetNames.join(', ')}`);

  const parsed = rowsFromWorkbook(workbook);

  if (!parsed.rows.length) {
    throw new Error('No se generaron registros válidos desde el Excel.');
  }

  const meta = {
    generatedAt: new Date().toISOString(),
    records: parsed.rows.length,
    sheetName: parsed.sheetName,
    headerRowIndex: parsed.headerIndex,
    headers: parsed.headers,
    sourceName: download.sourceName,
    sourceUrl: download.sourceUrl,
    contentType: download.contentType,
    bytes: download.bytes
  };

  await fs.writeFile(
    OUT_JSON,
    JSON.stringify(parsed.rows, null, 2),
    'utf8'
  );

  await fs.writeFile(
    OUT_META,
    JSON.stringify(meta, null, 2),
    'utf8'
  );

  console.log(`JSON generado correctamente: ${OUT_JSON}`);
  console.log(`Registros: ${parsed.rows.length}`);
  console.log(`Metadata generada: ${OUT_META}`);
}

main().catch(error => {
  console.error('Error generando data/control-donaciones.json');
  console.error(error);
  process.exit(1);
});
