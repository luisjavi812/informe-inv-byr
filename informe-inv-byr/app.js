/* ============================================================
   BYR · Depuración semanal de inventario
   app.js — módulo ES único (sin build step)

   1. Configuración y utilidades
   2. Normalización y clasificación
   3. Lectura del PDF del sistema (pdf.js)
   4. Lectura del Excel físico (SheetJS)
   5. Comparación, modelo y anomalías
   6. Estado de la aplicación
   7. Render
   8. Acciones
   9. Historial (local + Supabase)
   10. Ajustes
   11. Eventos y arranque
   ============================================================ */

/* ---------------------------- 1. CONFIGURACIÓN ---------------------------- */

const CDN = {
  pdf: 'https://cdn.jsdelivr.net/npm/pdfjs-dist@4.6.82/build/pdf.min.mjs',
  pdfWorker: 'https://cdn.jsdelivr.net/npm/pdfjs-dist@4.6.82/build/pdf.worker.min.mjs',
  xlsx: 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js',
  supabase: 'https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2/+esm'
};

/* Credenciales por defecto. Se pueden dejar vacías y configurar desde
   Ajustes, o inyectar antes de cargar app.js:
   <script>window.BYR_SUPABASE = { url: '...', anonKey: '...' }</script> */
const SUPABASE_DEFAULT = { url: '', anonKey: '' };

const LS = {
  testPlates: 'byr.testplates.v1',
  supabase: 'byr.supabase.v1',
  cortes: 'byr.cortes.v1',
  ultimo: 'byr.ultimo.v1'
};

/* Estados del comparativo: sin estado "taller distinto" — si la placa está en
   ambos lados es OK. EXCLUIDA no entra al comparativo: son ingresos del log con
   fecha posterior a la fecha de corte (se muestran en la vista Excluidas). */
const ESTADOS = ['ELIMINAR', 'INCLUIR', 'OK'];
const PENDIENTES = ['ELIMINAR', 'INCLUIR'];
const ESTADO_LABEL = { ELIMINAR: 'Eliminar', INCLUIR: 'Incluir', OK: 'OK', EXCLUIDA: 'Excluida' };

const TIPO_ORDER = [
  'Turbo Grande', 'Turbo Pequeña', 'Turbo', 'Tractomula', 'Sencillo', 'Doble Troque',
  'Trailer', 'Bus Grande', 'Bus Pequeño', 'Bus', 'Cabezote', 'Patineta', 'Mensualidad', 'Otro'
];

const DEFAULT_TEST_PLATES = ['ELO12345', 'OOO1234', 'AND1234', 'KAR1234', 'MENSUALIDADTALLER1'];

/* Lista de placas de prueba editable por el usuario (Ajustes). */
function getTestPlates() {
  const stored = readJSON(LS.testPlates, null);
  if (!Array.isArray(stored)) return DEFAULT_TEST_PLATES.slice();
  const list = stored.map(p => normPlaca(p)).filter(Boolean);
  return list.length ? list : DEFAULT_TEST_PLATES.slice();
}

function saveTestPlates(list) {
  const clean = Array.from(new Set((list || []).map(p => normPlaca(p)).filter(Boolean)));
  writeJSON(LS.testPlates, clean);
  return clean;
}

const ICONS = {
  copy: '<svg viewBox="0 0 24 24" aria-hidden="true"><rect x="9" y="9" width="11" height="11" rx="2"/><path d="M15 5.5A1.5 1.5 0 0 0 13.5 4h-8A1.5 1.5 0 0 0 4 5.5v8A1.5 1.5 0 0 0 5.5 15"/></svg>',
  alert: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3 2.8 19h18.4z"/><path d="M12 9v4.5M12 16.6v.2"/></svg>',
  check: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="m4 12.5 5 5L20 6.5"/></svg>',
  down: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 4v12"/><path d="m7.5 11.5 4.5 4.5 4.5-4.5"/><path d="M4 19h16"/></svg>',
  trash: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 7h16M9 7V5h6v2M6 7l1 13h10l1-13"/></svg>',
  pdf: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5"/><path d="M9 13h6M9 17h4"/></svg>',
  sheet: '<svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3.5" y="4.5" width="17" height="15" rx="2"/><path d="M3.5 9.5h17M9.5 9.5v10M15 9.5v10"/></svg>',
  users: '<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="9" cy="8.5" r="3.5"/><path d="M2.8 20a6.2 6.2 0 0 1 12.4 0"/><path d="M16 5.4a3.4 3.4 0 0 1 0 6.3M18.2 20a6 6 0 0 0-2.6-4.9"/></svg>',
  box: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="m12 3 8 4.5v9L12 21l-8-4.5v-9z"/><path d="m4 7.5 8 4.5 8-4.5M12 12v9"/></svg>',
  minus: '<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="8.5"/><path d="M8 12h8"/></svg>',
  plus: '<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="8.5"/><path d="M12 8v8M8 12h8"/></svg>',
  help: '<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="9"/><path d="M9.6 9.2a2.5 2.5 0 1 1 3.7 2.2c-.8.5-1.3 1-1.3 2.1"/><path d="M12 17h.01"/></svg>',
  cards: '<svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="6" width="13" height="12" rx="2"/><path d="M8 4h11a2 2 0 0 1 2 2v11"/></svg>',
  link: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M9.5 14.5 14.5 9.5"/><path d="M7 13 5 15a3.5 3.5 0 0 0 5 5l2-2"/><path d="M17 11l2-2a3.5 3.5 0 0 0-5-5l-2 2"/></svg>',
  clock: '<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="8.5"/><path d="M12 7.5V12l3 2"/></svg>',
  x: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="m6 6 12 12M18 6 6 18"/></svg>'
};

const $ = (sel, root) => (root || document).querySelector(sel);
const $$ = (sel, root) => Array.from((root || document).querySelectorAll(sel));

function esc(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

const nf = new Intl.NumberFormat('es-CO');
const fmt = n => nf.format(Number(n) || 0);

function pct(part, total) {
  if (!total) return 0;
  return Math.round((part / total) * 100);
}

function stripAccents(s) {
  return String(s == null ? '' : s)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

function debounce(fn, wait) {
  let t = 0;
  return (...args) => {
    clearTimeout(t);
    t = setTimeout(() => fn(...args), wait);
  };
}

function todayISO() {
  const d = new Date();
  const off = d.getTimezoneOffset();
  return new Date(d.getTime() - off * 60000).toISOString().slice(0, 10);
}

function formatDate(iso, withTime) {
  if (!iso) return '—';
  const raw = String(iso);
  const date = new Date(raw.length <= 10 ? raw + 'T12:00:00' : raw);
  if (Number.isNaN(date.getTime())) return raw;
  const txt = date.toLocaleDateString('es-CO', { day: '2-digit', month: 'short', year: 'numeric' });
  if (!withTime) return txt.replace('.', '');
  return txt.replace('.', '') + ' · ' + date.toLocaleTimeString('es-CO', { hour: '2-digit', minute: '2-digit' });
}

function readJSON(key, fallback) {
  try {
    const raw = localStorage.getItem(key);
    return raw ? JSON.parse(raw) : fallback;
  } catch (e) {
    return fallback;
  }
}

function writeJSON(key, value) {
  try {
    localStorage.setItem(key, JSON.stringify(value));
    return true;
  } catch (e) {
    return false;
  }
}

function loadScript(src) {
  return new Promise((resolve, reject) => {
    const existing = $$('script').find(s => s.src === src);
    if (existing) {
      if (existing.dataset.loaded === '1') return resolve();
      existing.addEventListener('load', () => resolve(), { once: true });
      existing.addEventListener('error', () => reject(new Error('No se pudo cargar ' + src)), { once: true });
      return;
    }
    const s = document.createElement('script');
    s.src = src;
    s.async = true;
    s.onload = () => { s.dataset.loaded = '1'; resolve(); };
    s.onerror = () => reject(new Error('No se pudo cargar ' + src));
    document.head.appendChild(s);
  });
}

function downloadFile(filename, content, mime) {
  const blob = new Blob([content], { type: mime || 'text/plain;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1500);
}

async function copyText(text) {
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch (e) {
    try {
      const ta = document.createElement('textarea');
      ta.value = text;
      ta.setAttribute('readonly', '');
      ta.style.position = 'fixed';
      ta.style.opacity = '0';
      document.body.appendChild(ta);
      ta.select();
      const ok = document.execCommand('copy');
      ta.remove();
      return ok;
    } catch (err) {
      return false;
    }
  }
}

let toastSeq = 0;
function toast(message, type) {
  const host = $('#toasts');
  if (!host) return;
  if ([...host.children].some(c => c.dataset.msg === message)) return;
  const el = document.createElement('div');
  el.dataset.msg = message;
  el.className = 'toast toast--' + (type || 'ok');
  el.innerHTML = (type === 'err' ? ICONS.alert : ICONS.check) + '<span>' + esc(message) + '</span>';
  host.appendChild(el);
  const id = ++toastSeq;
  setTimeout(() => {
    el.style.transition = 'opacity .18s ease, transform .18s ease';
    el.style.opacity = '0';
    el.style.transform = 'translateY(4px)';
    setTimeout(() => { if (el.parentNode) el.remove(); }, 200);
  }, type === 'err' ? 5200 : 2800);
  return id;
}

/* ---------------------- 2. NORMALIZACIÓN Y CLASIFICACIÓN ---------------------- */

function normPlaca(value) {
  return stripAccents(String(value == null ? '' : value))
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, '');
}

function normR(value) {
  return stripAccents(String(value == null ? '' : value))
    .toUpperCase()
    .replace(/\s+/g, ' ')
    .trim();
}

function rKey(value) {
  return normR(value).replace(/[^A-Z0-9]/g, '');
}

function isGarbageR(value) {
  const raw = normR(value);
  if (!raw) return { garbage: true, razon: 'vacío' };
  const compact = raw.replace(/[^A-Z0-9]/g, '');
  if (!compact) return { garbage: true, razon: 'solo signos' };
  if (/^(NA|NV|NC|SN|NO|SIN|NINGUNO|NINGUNA|SINPLACA|NOPLACA|SINR)$/.test(compact)) {
    return { garbage: true, razon: 'sin código' };
  }
  if (/^X+$/.test(compact)) return { garbage: true, razon: 'relleno X' };
  if (/^0X+$/.test(compact)) return { garbage: true, razon: 'relleno 0x' };
  if (/^[A-Z]$/.test(compact)) return { garbage: true, razon: 'un carácter' };
  if (/^[,-.;:_]+$/.test(raw)) return { garbage: true, razon: 'solo signos' };
  return { garbage: false, razon: '' };
}

function plateFormat(plateNorm) {
  const p = normPlaca(plateNorm);
  if (!p) return 'vacia';
  if (/^[A-Z]{3}\d{3}$/.test(p)) return 'auto';
  if (/^[A-Z]\d{4,6}$/.test(p)) return 'remolque';
  if (getTestPlates().includes(p)) return 'prueba';
  return 'otro';
}

function isTestPlate(plateNorm) {
  const p = normPlaca(plateNorm);
  if (!p) return false;
  if (getTestPlates().includes(p)) return true;
  if (/^MENSUALIDAD/.test(p)) return true;
  if (/^(TEST|PRUEBA|DEMO|TEMP)/.test(p)) return true;
  return false;
}

function tipoFromName(name) {
  const s = stripAccents(name).toUpperCase();
  if (!s) return 'Otro';
  if (/MENSUALIDAD/.test(s)) return 'Mensualidad';
  if (/PATINETA/.test(s)) return 'Patineta';
  if (/TURBO\s*GRANDE/.test(s)) return 'Turbo Grande';
  if (/TURBO\s*PEQUE/.test(s)) return 'Turbo Pequeña';
  if (/TURBO/.test(s)) return 'Turbo';
  if (/TRACTOMULA|TRAC\s*PATI|\bTRAC\b|TRACTOCAMION/.test(s)) return 'Tractomula';
  if (/DOBLE\s*TROQUE/.test(s)) return 'Doble Troque';
  if (/SENCILLO/.test(s)) return 'Sencillo';
  if (/TRAILER|REMOLQUE/.test(s)) return 'Trailer';
  if (/BUS\s*GRANDE/.test(s)) return 'Bus Grande';
  if (/BUS\s*PEQUE/.test(s)) return 'Bus Pequeño';
  if (/\bBUS\b/.test(s)) return 'Bus';
  if (/CABEZOTE/.test(s)) return 'Cabezote';
  return 'Otro';
}

const TALLER_NOISE = new Set(['TALLER', 'TALLERES', 'PARQUEADERO', 'PARQUEO', 'PATIO', 'BYR', 'BR', 'SAS', 'SA', 'LTDA', 'SAS', 'DE', 'EL', 'LA', 'LOS', 'LAS']);

function normTaller(name) {
  const raw = stripAccents(String(name == null ? '' : name)).toUpperCase();
  const tokens = raw
    .replace(/&/g, ' ')
    .replace(/[^A-Z0-9\s]/g, ' ')
    .split(/\s+/)
    .filter(Boolean)
    .filter(t => !TALLER_NOISE.has(t));
  return tokens.join(' ');
}

function sameTaller(a, b) {
  const na = normTaller(a);
  const nb = normTaller(b);
  if (!na || !nb) {
    return stripAccents(String(a || '')).trim().toUpperCase() === stripAccents(String(b || '')).trim().toUpperCase();
  }
  return na === nb;
}

function cleanName(name) {
  return stripAccents(String(name == null ? '' : name))
    .replace(/\u00a0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/* ------------------------- 3. PDF DEL SISTEMA ------------------------- */

const PDF_NOISE = [
  /^REPRESENTACIONES BYR/i,
  /^NIT\s*:/i,
  /^TEL\s*:/i,
  /^RESPONSABLE DE IVA/i,
  /^KM\s*\d/i,
  /^INVENTARIO\s+REPRESENTACIONES/i,
  /^\d{4}-\d{2}-\d{2}/,
  /^-{5,}$/,
  /^_{5,}$/,
  /^PLACA\s+No\.?\s*R/i,
  /^VERIFICADO(\s+POR)?$/i,
  /^\.$/,
  /^\d{1,6}$/
];

function isPdfNoise(line) {
  const s = String(line || '').trim();
  if (!s) return true;
  return PDF_NOISE.some(re => re.test(s));
}

function looksLikePlateToken(token) {
  const t = normPlaca(token);
  if (t.length < 3 || t.length > 24) return false;
  if (!/[A-Z]/.test(t) || !/\d/.test(t)) return false;
  return true;
}

const PLATE_STOP = new Set(['NIT', 'TEL', 'PLACA', 'PLACAS', 'TOTAL', 'VERIFICADO', 'PAGINA', 'NO', 'R', 'N']);

/** Token de placa en las filas del PDF: admite casos como TEK, SCHRA o
 *  CHASISSINPLATA (sin dígitos) para no perder filas y reportarlos como anomalía. */
function isSystemPlateToken(token) {
  const t = normPlaca(token);
  if (t.length < 3 || t.length > 24) return false;
  if (!/^[A-Z0-9]+$/.test(t)) return false;
  if (!/[A-Z]/.test(t)) return false;
  if (PLATE_STOP.has(t)) return false;
  return true;
}

/**
 * Convierte las líneas de texto del PDF en filas del sistema.
 * @param {string[][]} pages líneas por página
 * @returns {{sections:Array, rows:Array, warnings:Array, declaredTotal:number|null}}
 */
function parseSystemLines(pages) {
  const sections = [];
  const rows = [];
  const warnings = [];
  let current = null;
  let declaredTotal = null;
  let sectionHeaderCount = 0;

  (pages || []).forEach((lines, pageIdx) => {
    (lines || []).forEach(raw => {
      const s = String(raw || '').replace(/\s+/g, ' ').trim();
      if (!s) return;
      if (/^TOTAL\s+VEH[ií]CULOS/i.test(s)) {
        const m = s.match(/(\d{1,5})\s*$/);
        if (m) declaredTotal = Number(m[1]);
        return;
      }
      if (isPdfNoise(s)) return;

      const header = s.match(/^(.+?)\s*\(\s*(\d{1,4})\s*\)\s*$/);
      if (header && !/^PLACA/i.test(header[1]) && !/TOTAL/i.test(header[1])) {
        const name = cleanName(header[1]);
        const key = 'sys:' + name;
        sectionHeaderCount += 1;
        const existing = sections.find(x => x.key === key);
        if (existing) {
          existing.declared += Number(header[2]) || 0;
          current = existing;
        } else {
          current = { key, name, declared: Number(header[2]) || 0, page: pageIdx + 1 };
          sections.push(current);
        }
        return;
      }

      const withMark = s.match(/^(\S+)\s+(.*?)\s*_{3,}$/);
      if (withMark && current) {
        const placa = withMark[1];
        if (!isSystemPlateToken(placa)) {
          warnings.push({ tipo: 'línea', mensaje: `Línea no reconocida en «${current.name}»: ${s}` });
          return;
        }
        rows.push(makeSystemRow(placa, withMark[2], current, pageIdx));
        return;
      }

      if (current) {
        const parts = s.split(' ');
        if (parts.length >= 1 && looksLikePlateToken(parts[0]) && !/[a-z]{4,}/.test(s)) {
          rows.push(makeSystemRow(parts[0], parts.slice(1).join(' '), current, pageIdx));
          warnings.push({ tipo: 'formato', mensaje: `Fila sin marca de verificación en «${current.name}»: ${s}` });
          return;
        }
        if (!/^\d+$/.test(s)) {
          warnings.push({ tipo: 'línea', mensaje: `Línea ignorada en «${current.name}»: ${s}` });
        }
      }
    });
  });

  if (declaredTotal != null && rows.length !== declaredTotal) {
    warnings.push({
      tipo: 'total',
      mensaje: `El PDF declara ${fmt(declaredTotal)} vehículos y se leyeron ${fmt(rows.length)} filas.`
    });
  }

  sections.forEach(sec => {
    const n = rows.filter(r => r.sectionKey === sec.key).length;
    sec.count = n;
    if (sec.declared && sec.declared !== n) {
      warnings.push({ tipo: 'cuenta', mensaje: `«${sec.name}» declara ${sec.declared} y se leyeron ${n}.` });
    }
  });

  if (!sections.length) {
    warnings.push({ tipo: 'estructura', mensaje: 'No se detectaron bloques con el formato «TALLER (n)». Revisa que el PDF sea el inventario del sistema.' });
  }

  return { sections, rows, warnings, declaredTotal, sectionHeaderCount };
}

function makeSystemRow(placaRaw, rRaw, section, pageIdx) {
  const placa = String(placaRaw || '').trim().toUpperCase();
  const r = normR(rRaw);
  return {
    placa,
    placaNorm: normPlaca(placa),
    r,
    taller: section.name,
    sectionKey: section.key,
    tipo: tipoFromName(section.name),
    origen: 'sistema',
    page: pageIdx + 1
  };
}

/* Texto extraído con pdf.js → líneas ordenadas por eje Y.
   Cada página es un arreglo de líneas. */
function pdfItemsToLines(items) {
  const clean = (items || [])
    .filter(it => it && it.str && it.str.trim())
    .map(it => ({ x: it.transform[4], y: it.transform[5], str: it.str }));
  clean.sort((a, b) => b.y - a.y || a.x - b.x);

  const lines = [];
  let cur = null;
  for (const it of clean) {
    if (!cur || Math.abs(cur.y - it.y) > 2.5) {
      cur = { y: it.y, parts: [] };
      lines.push(cur);
    }
    cur.parts.push(it);
  }
  return lines.map(l => l.parts.sort((a, b) => a.x - b.x).map(p => p.str).join(' ').replace(/\s+/g, ' ').trim());
}

async function readPdfFile(file) {
  const pdfjs = await import(CDN.pdf);
  pdfjs.GlobalWorkerOptions.workerSrc = CDN.pdfWorker;
  const data = await file.arrayBuffer();
  const doc = await pdfjs.getDocument({ data }).promise;
  const pages = [];
  for (let i = 1; i <= doc.numPages; i++) {
    const page = await doc.getPage(i);
    const content = await page.getTextContent();
    pages.push(pdfItemsToLines(content.items));
    page.cleanup();
  }
  doc.destroy();
  return pages;
}

/* ------------------------- 4. EXCEL FÍSICO ------------------------- */

function isTitleCell(value) {
  if (typeof value !== 'string') return false;
  const s = value.trim();
  if (!s || s.length < 3) return false;
  if (/^[A-Z0-9]{4,}$/.test(s) && /\d/.test(s)) return false;
  if (/\d/.test(s)) return false;
  if (!/[A-Za-zÁÉÍÓÚÑÜáéíóúñü]{3}/.test(s)) return false;
  if (/^(PLACA|NOMBRE|OBSERV|CODIGO|CÓDIGO|TIPO|CLASE|FECHA|TOTAL|ITEM|No|N\.?)$/i.test(s)) return false;
  return true;
}

function plateFromCell(value) {
  if (typeof value === 'number') return null;
  if (value instanceof Date) return null;
  const s = String(value == null ? '' : value).replace(/\s+/g, ' ').trim().toUpperCase();
  if (!s) return null;
  if (/^\d+$/.test(s)) return null;
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return null;
  const first = s.split(' ')[0].replace(/-/g, '');
  if (!looksLikePlateToken(first)) return null;
  if (/^(NIT|TEL|PLACA|TOTAL|FECHA)$/i.test(first)) return null;
  return { placa: s.split(' ')[0], r: s.split(' ').slice(1).join(' ') };
}

function colLetter(idx) {
  let n = Number(idx) + 1;
  let out = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

function detectTabular(aoa) {
  const limit = Math.min((aoa || []).length, 12);
  for (let r = 0; r < limit; r++) {
    const row = aoa[r] || [];
    const cells = row.map(v => stripAccents(String(v == null ? '' : v)).trim().toUpperCase());
    const pi = cells.findIndex(v => /^PLACA$/.test(v) || /^PLACAS?$/.test(v));
    if (pi < 0) continue;
    const nonEmpty = cells.filter(Boolean).length;
    if (nonEmpty < 2) continue;
    const ti = cells.findIndex(v => /TALLER|PARQUE|UBICAC|PATIO|SITIO|BODEGA/.test(v));
    const ri = cells.findIndex(v => /^R$|^NO\.? ?R$|^NRO\.? ?R$|CODIGO|CÓDIGO/.test(v));
    const toi = cells.findIndex(v => /^TIPO$|CLASE|VEHICULO|VEHÍCULO/.test(v));
    return { headerRow: r, pi, ti, ri, toi };
  }
  return null;
}

/* ---- Detección de la hoja de ingresos por CONTENIDO (tiene fechas), no por
   nombre: el nombre cambia cada semana (§2B/§4 del flujo). ---- */

const FECHA_RE = /^(\d{4})-(\d{1,2})-(\d{1,2})([T\s].*)?$|^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})([\sT].*)?$/;

const MESES = { ENE: 1, FEB: 2, MAR: 3, ABR: 4, MAY: 5, JUN: 6, JUL: 7, AGO: 8, SEP: 9, OCT: 10, NOV: 11, DIC: 12, JAN: 1, APR: 4, AUG: 8, DEC: 12 };

/** Devuelve 'YYYY-MM-DD' si la celda contiene fecha/hora, o ''. */
function cellDateISO(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const d = value;
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  }
  const s = String(value == null ? '' : value).trim();
  if (!s || s.length < 8) return '';
  const m = s.match(FECHA_RE);
  if (m) {
    if (m[1]) {
      const y = Number(m[1]);
      const mo = Number(m[2]);
      const d = Number(m[3]);
      return (y > 1990 && mo >= 1 && mo <= 12 && d >= 1 && d <= 31)
        ? `${y}-${String(mo).padStart(2, '0')}-${String(d).padStart(2, '0')}` : '';
    }
    const d = Number(m[5]);
    const mo = Number(m[6]);
    let y = Number(m[7]);
    if (y < 100) y += 2000;
    return (y > 1990 && mo >= 1 && mo <= 12 && d >= 1 && d <= 31)
      ? `${y}-${String(mo).padStart(2, '0')}-${String(d).padStart(2, '0')}` : '';
  }
  // Formato con mes en texto: 16-nov-2026 / 16 NOV 2026
  const t = s.match(/^(\d{1,2})[\s\-\/]([A-Za-z]{3,4})[\s\-\/](\d{2,4})/);
  if (t) {
    const mo = MESES[t[2].toUpperCase().slice(0, 4).replace('.', '')] || MESES[t[2].toUpperCase().slice(0, 3)];
    const d = Number(t[1]);
    let y = Number(t[3]);
    if (y < 100) y += 2000;
    return (y > 1990 && mo && d >= 1 && d <= 31)
      ? `${y}-${String(mo).padStart(2, '0')}-${String(d).padStart(2, '0')}` : '';
  }
  return '';
}

/** Token de placa suelto dentro de una celda de la hoja de ingresos. */
function placaFromIngresoCell(value) {
  if (value == null) return null;
  if (value instanceof Date) return null;
  const s = String(value).trim();
  if (!s || /\s/.test(s) || s.includes('@')) return null;
  if (/^\d+([.,]\d+)?$/.test(s)) return null;               // id numérico o cédula
  if (cellDateISO(s)) return null;
  const t = normPlaca(s);
  if (t.length < 5 || t.length > 12) return null;
  if (!/^[A-Z]/.test(t) || !/\d/.test(t)) return null;
  if (/^(NIT|TEL|PLACA|TOTAL|FECHA|ID)/.test(t)) return null;
  return t;
}

function looksLikeIngresosSheet(aoa) {
  let rows = 0;
  const limit = Math.min((aoa || []).length, 400);
  for (let r = 0; r < limit; r++) {
    const row = aoa[r] || [];
    let hasDate = false;
    let hasPlaca = false;
    for (const cell of row) {
      if (!hasDate && cellDateISO(cell)) hasDate = true;
      if (!hasPlaca && placaFromIngresoCell(cell)) hasPlaca = true;
      if (hasDate && hasPlaca) break;
    }
    if (hasDate && hasPlaca) rows += 1;
  }
  return rows >= 3;
}

/** Extrae [placa, categoría, fecha] de la hoja de ingresos. */
function parseIngresosSheet(aoa, sheetName) {
  const rows = [];
  (aoa || []).forEach(row => {
    if (!row || !row.length) return;
    let placa = null;
    let placaIdx = -1;
    let fecha = '';
    row.forEach((cell, i) => {
      if (!placa) {
        const p = placaFromIngresoCell(cell);
        if (p) { placa = p; placaIdx = i; }
      }
      if (!fecha) fecha = cellDateISO(cell);
    });
    if (!placa || !fecha) return;
    // Categoría: celda posterior a la placa que parezca taller/parqueadero.
    let categoria = '';
    let codigoR = '';
    for (let i = placaIdx + 1; i < Math.min(placaIdx + 6, row.length); i++) {
      const c = row[i];
      if (typeof c === 'number') continue;
      const txt = String(c == null ? '' : c).trim();
      if (!txt || cellDateISO(txt)) continue;
      if (!categoria && /TALLER|PARQUEO|PATIO/i.test(txt)) { categoria = cleanName(txt); continue; }
      if (!codigoR && /^[RS]-?\d{4,6}$/i.test(txt.replace(/\s+/g, ''))) codigoR = txt.toUpperCase();
    }
    rows.push({ placa, placaNorm: placa, categoria, codigoR, conductor: '', fecha, hoja: sheetName });
  });
  return rows;
}

/**
 * Recibe hojas ya convertidas a matrices (array de filas) y devuelve
 * secciones + filas del inventario físico, más la hoja de ingresos detectada.
 * @param {Array<{name:string, rows:Array<Array>}>} sheets
 */
function parsePhysicalSheets(sheets) {
  const out = { rows: [], sections: [], sheetStats: [], skipped: [], ingresos: [], warnings: [] };
  let seq = 0;

  const addSection = (sheetName, name, col, startRow, keyOverride) => {
    const key = keyOverride || `fis:${seq + 1}:${col == null ? 'tab' : colLetter(col)}`;
    const existing = out.sections.find(s => s.key === key);
    if (existing) return existing;
    seq += 1;
    const sec = {
      key,
      name: cleanName(name) || cleanName(sheetName),
      sheet: sheetName,
      col: col == null ? null : col,
      startRow: startRow == null ? null : startRow,
      count: 0
    };
    out.sections.push(sec);
    return sec;
  };

  (sheets || []).forEach(sheet => {
    const sheetName = cleanName(sheet.name || 'Hoja');
    const aoa = sheet.rows || [];
    if (!aoa.length) {
      out.warnings.push({ tipo: 'hoja', mensaje: `La hoja «${sheetName}» está vacía.` });
      return;
    }

    // Hoja de ingresos: se detecta por contenido (filas con fecha/hora + placa),
    // nunca por el nombre de la hoja.
    if (!detectTabular(aoa) && looksLikeIngresosSheet(aoa)) {
      const ingresos = parseIngresosSheet(aoa, sheetName);
      if (ingresos.length) {
        out.ingresos = out.ingresos.concat(ingresos);
        out.skipped.push({
          nombre: sheetName,
          motivo: `hoja de ingresos (${ingresos.length} registros con fecha): se usa para excluir placas ingresadas después del corte`,
          ingresos: true
        });
        return;
      }
    }

    const before = out.rows.length;
    const tab = detectTabular(aoa);

    if (tab) {
      const width = aoa.reduce((m, r) => Math.max(m, (r || []).length), 0);
      const byTaller = new Map();
      let carry = '';
      for (let r = tab.headerRow + 1; r < aoa.length; r++) {
        const row = aoa[r] || [];
        const cell = row[tab.pi];
        const parsed = plateFromCell(cell);
        if (!parsed) continue;
        let taller = tab.ti >= 0 ? cleanName(row[tab.ti]) : '';
        if (taller) carry = taller;
        else taller = carry || sheetName;
        const key = taller.toUpperCase();
        if (!byTaller.has(key)) {
          const sec = addSection(sheetName, taller, tab.pi, r);
          sec.tabular = true;
          byTaller.set(key, sec);
        }
        const sec = byTaller.get(key);
        const tipo = tab.toi >= 0 ? cleanName(row[tab.toi]) : '';
        pushPhysicalRow(out, sec, parsed, sheetName, tipo);
      }
      if (byTaller.size === 0) {
        out.warnings.push({ tipo: 'hoja', mensaje: `La hoja «${sheetName}» tiene encabezado tipo tabla pero no se encontraron placas.` });
      }
    } else {
      const width = aoa.reduce((m, r) => Math.max(m, Array.isArray(r) ? r.length : 0), 0);
      const titleCols = new Set();
      for (let c = 0; c < width; c++) {
        for (let r = 0; r < aoa.length; r++) {
          const v = (aoa[r] || [])[c];
          if (isTitleCell(v)) { titleCols.add(c); break; }
        }
      }

      let activeCols = Array.from(titleCols).sort((a, b) => a - b);
      const hasTitles = activeCols.length > 0;
      if (!hasTitles) {
        // Hoja sin títulos: cuadrícula simple (ej. PARQUEO) → una sola sección.
        activeCols = [];
        for (let c = 0; c < width; c++) {
          const hasPlate = aoa.some(row => plateFromCell((row || [])[c]));
          if (hasPlate) activeCols.push(c);
        }
      }

      if (!activeCols.length) {
        out.warnings.push({ tipo: 'hoja', mensaje: `No se detectaron placas en la hoja «${sheetName}».` });
        return;
      }

      const defaultName = sheetName;

      if (!hasTitles) {
        const sec = addSection(sheetName, sheetName, null, null);
        activeCols.forEach(c => {
          for (let r = 0; r < aoa.length; r++) {
            const row = aoa[r] || [];
            const raw = row[c];
            if (raw == null) continue;
            const parsed = plateFromCell(raw);
            if (!parsed) continue;
            pushPhysicalRow(out, sec, parsed, sheetName, '');
          }
        });
      } else {
        activeCols.forEach(c => {
          let cur = null;
          for (let r = 0; r < aoa.length; r++) {
            const row = aoa[r] || [];
            const raw = row[c];
            if (raw == null) continue;
            if (typeof raw === 'string' && isTitleCell(raw)) {
              // Títulos que son solo "TALLER": un solo bloque "TALLER (sin nombre)".
              const bare = /^TALLER$/i.test(stripAccents(raw.trim()));
              cur = addSection(sheetName, bare ? 'TALLER (sin nombre)' : raw.trim(), c, r, bare ? 'fis:unnamed' : null);
              continue;
            }
            const parsed = plateFromCell(raw);
            if (!parsed) continue;
            if (!cur) cur = addSection(sheetName, defaultName, c, null);

            // Código R en una columna contigua que no es columna de placas.
            let rCode = parsed.r;
            if (!rCode) {
              for (let cc = c + 1; cc < Math.min(c + 3, width); cc++) {
                if (activeCols.includes(cc)) continue;
                const neighbor = row[cc];
                if (neighbor == null) continue;
                if (typeof neighbor === 'number') { rCode = String(neighbor); break; }
                if (neighbor instanceof Date) continue;
                const txt = String(neighbor).trim();
                if (!txt || isTitleCell(txt)) continue;
                rCode = txt;
                break;
              }
            }
            pushPhysicalRow(out, cur, { placa: parsed.placa, r: rCode }, sheetName, '');
          }
        });
      }
    }

    out.sheetStats.push({ name: sheetName, rows: out.rows.length - before, tabular: !!tab });
  });

  // Nombres visibles: si un título se repite, se agrega su ubicación.
  const nameCount = new Map();
  out.sections.forEach(s => {
    const k = s.name.toUpperCase();
    nameCount.set(k, (nameCount.get(k) || 0) + 1);
  });
  out.sections.forEach(s => {
    const duplicated = nameCount.get(s.name.toUpperCase()) > 1;
    s.display = duplicated && s.col != null
      ? `${s.name} · col ${colLetter(s.col)}${s.startRow != null ? ' f' + (s.startRow + 1) : ''}`
      : s.name;
    s.tipo = tipoFromName(s.name);
  });

  const byKey = new Map(out.sections.map(s => [s.key, s]));
  out.rows.forEach(row => {
    const sec = byKey.get(row.sectionKey);
    row.taller = sec ? sec.display : row.taller;
  });
  out.sections = out.sections.filter(s => out.rows.some(r => r.sectionKey === s.key));
  out.sections.forEach(s => { s.count = out.rows.filter(r => r.sectionKey === s.key).length; });

  if (!out.rows.length) {
    out.warnings.push({ tipo: 'estructura', mensaje: 'No se detectaron placas en el archivo físico. Revisa el formato del Excel.' });
  }
  return out;
}

function pushPhysicalRow(out, section, parsed, sheetName, tipoOverride) {
  const placa = String(parsed.placa || '').trim().toUpperCase();
  const r = normR(parsed.r);
  out.rows.push({
    placa,
    placaNorm: normPlaca(placa),
    r,
    taller: section.name,
    sectionKey: section.key,
    tipo: tipoOverride || tipoFromName(section.name),
    origen: 'fisico',
    hoja: sheetName
  });
}

async function readPhysicalFiles(files) {
  await loadScript(CDN.xlsx);
  const XLSX = globalThis.XLSX;
  if (!XLSX) throw new Error('SheetJS no está disponible.');
  const sheets = [];
  for (const file of files) {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(new Uint8Array(buf), { type: 'array', cellDates: true });
    wb.SheetNames.forEach(name => {
      const ws = wb.Sheets[name];
      /* raw:true conserva los valores: con cellDates las fechas llegan como
         Date, que es lo que usa la detección de la hoja de ingresos. */
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });
      sheets.push({ name, rows });
    });
  }
  return parsePhysicalSheets(sheets);
}

/* ------------------- 5. COMPARACIÓN, MODELO Y ANOMALÍAS ------------------- */

function groupByPlate(rows) {
  const map = new Map();
  rows.forEach(row => {
    if (!row.placaNorm) return;
    if (!map.has(row.placaNorm)) map.set(row.placaNorm, []);
    map.get(row.placaNorm).push(row);
  });
  return map;
}

function pickR(sysRow, fisRow) {
  if (sysRow && !isGarbageR(sysRow.r).garbage) return sysRow.r;
  if (fisRow && !isGarbageR(fisRow.r).garbage) return fisRow.r;
  return (sysRow && sysRow.r) || (fisRow && fisRow.r) || '';
}

function buildModel() {
  const sysBy = groupByPlate(state.system.rows);
  const fisBy = groupByPlate(state.physical.rows);
  const rows = [];

  /* ---- Exclusiones por fecha (§4): ingresos del log con fecha posterior a la
     fecha de corte. No se les hace nada: ya están en el sistema. ---- */
  const excludedMap = new Map();
  const corte = state.corteFecha || '';
  if (corte) {
    (state.ingresos || []).forEach(ing => {
      if (!ing.fecha || ing.fecha <= corte) return;
      const pn = normPlaca(ing.placa);
      if (!pn || excludedMap.has(pn)) return;
      excludedMap.set(pn, {
        placa: ing.placa,
        placaNorm: pn,
        categoria: ing.categoria || '',
        codigoR: ing.codigoR || '',
        fecha: ing.fecha,
        enSistema: false,
        catSistema: ''
      });
    });
  }

  const esParqueo = f => /^PARQUEO\b/i.test(String(f.taller || ''));

  for (const [plate, list] of sysBy) {
    const s = list[0];
    if (excludedMap.has(plate)) {
      const ex = excludedMap.get(plate);
      ex.enSistema = true;
      ex.catSistema = s.taller;
      continue;
    }
    const flist = fisBy.get(plate) || [];
    const f = flist[0] || null;
    const estado = f ? 'OK' : 'ELIMINAR';
    const motivo = f ? 'Coincide con el conteo físico' : 'No aparece en el conteo físico';

    // PARQUEO (sin nombre): el parqueadero efectivo de la placa es la categoría
    // con la que aparece en el sistema; si no existe, "PARQUEO (sin categoría)".
    let fisTaller = f ? f.taller : '';
    let fisGroupKey = f ? f.sectionKey : '';
    let fisGroupLabel = f ? f.taller : '';
    if (f && esParqueo(f)) {
      fisTaller = s.taller;
      fisGroupKey = 'fis:parq:' + (normPlaca(s.taller) || 'sin');
      fisGroupLabel = s.taller ? `PARQUEO · ${s.taller}` : 'PARQUEO (sin categoría)';
    }

    rows.push({
      placa: s.placa,
      placaNorm: plate,
      r: pickR(s, f),
      rSistema: s.r,
      rFisico: f ? f.r : '',
      estado,
      motivo,
      sysTaller: s.taller,
      fisTaller,
      fisGroupKey,
      fisGroupLabel,
      sysKey: s.sectionKey,
      fisKey: f ? f.sectionKey : '',
      tipo: tipoFromName(s.taller),
      origen: f ? 'ambos' : 'sistema'
    });
  }

  for (const [plate, list] of fisBy) {
    if (sysBy.has(plate) || excludedMap.has(plate)) continue;
    const f = list[0];
    let fisTaller = f.taller;
    let fisGroupKey = f.sectionKey;
    let fisGroupLabel = f.taller;
    if (esParqueo(f)) {
      fisTaller = 'PARQUEO (sin categoría)';
      fisGroupKey = 'fis:parq:sin';
      fisGroupLabel = 'PARQUEO (sin categoría)';
    }
    rows.push({
      placa: f.placa,
      placaNorm: plate,
      r: pickR(null, f),
      rSistema: '',
      rFisico: f.r,
      estado: 'INCLUIR',
      motivo: 'No existe en el inventario del sistema',
      sysTaller: '',
      fisTaller,
      fisGroupKey,
      fisGroupLabel,
      sysKey: '',
      fisKey: f.sectionKey,
      tipo: tipoFromName(f.taller),
      origen: 'fisico'
    });
  }

  rows.sort((a, b) => a.placaNorm.localeCompare(b.placaNorm, 'es'));

  const counts = { ELIMINAR: 0, INCLUIR: 0, OK: 0 };
  rows.forEach(r => { counts[r.estado] += 1; });

  const excluded = Array.from(excludedMap.values()).sort((a, b) => b.fecha.localeCompare(a.fecha));
  const stats = {
    total: rows.length,
    sistema: state.system.rows.length,
    sistemaUnicos: sysBy.size,
    /* Total físico en placas únicas: los bloques pueden repetir una placa
       (se informa como anomalía), pero el conteo es 1 placa = 1 vehículo. */
    fisico: fisBy.size,
    fisicoFilas: state.physical.rows.length,
    fisicoUnicos: fisBy.size,
    excluidas: excluded.length,
    excluidasEnSistema: excluded.filter(e => e.enSistema).length,
    ...counts
  };

  const model = { rows, stats, sysBy, fisBy, excluded, sections: buildSectionIndex(rows) };
  model.anomalies = detectAnomalies(model);
  model.anomalyCount = Object.values(model.anomalies).reduce((n, arr) => n + arr.length, 0);
  return model;
}

function buildSectionIndex(rows) {
  const map = new Map();
  rows.forEach(row => {
    const entries = [
      row.sysKey ? { key: row.sysKey, label: row.sysTaller, origen: 'sistema' } : null,
      row.fisKey ? { key: row.fisKey, label: row.fisTaller, origen: 'fisico' } : null
    ].filter(Boolean);
    entries.forEach(e => {
      if (!map.has(e.key)) map.set(e.key, { key: e.key, label: e.label, tipo: tipoFromName(e.label), declared: null, count: 0 });
      map.get(e.key).count += 1;
    });
  });
  (state.system.sections || []).forEach(s => {
    if (map.has(s.key)) map.get(s.key).declared = s.declared;
  });
  return Array.from(map.values());
}

function detectAnomalies(model) {
  const sys = state.system.rows;
  const fis = state.physical.rows;
  const all = sys.concat(fis);

  const dup = sideRows => {
    const map = groupByPlate(sideRows);
    const out = [];
    for (const [plate, list] of map) {
      if (list.length < 2) continue;
      out.push({
        placa: plate,
        detalle: list.map(r => r.taller).filter((v, i, a) => a.indexOf(v) === i).join(' · '),
        veces: list.length
      });
    }
    return out.sort((a, b) => a.placa.localeCompare(b.placa));
  };

  const pruebas = [];
  const seenTest = new Set();
  all.forEach(r => {
    if (!isTestPlate(r.placaNorm)) return;
    const k = r.origen + '|' + r.placaNorm;
    if (seenTest.has(k)) return;
    seenTest.add(k);
    pruebas.push({ placa: r.placa, detalle: (r.origen === 'sistema' ? 'Sistema' : 'Físico') + ' · ' + r.taller });
  });

  const rBasura = [];
  all.forEach(r => {
    if (r.origen === 'fisico' && !r.r) return;
    const g = isGarbageR(r.r);
    if (!g.garbage) return;
    rBasura.push({
      placa: r.placa,
      detalle: `${r.origen === 'sistema' ? 'Sistema' : 'Físico'} · ${r.taller}`,
      extra: r.r ? `R: ${r.r}` : 'R vacío',
      razon: g.razon
    });
  });

  const byR = new Map();
  all.forEach(r => {
    if (r.origen === 'fisico' && !r.r) return;
    if (isGarbageR(r.r).garbage) return;
    const k = rKey(r.r);
    if (!k || k.length < 3) return;
    if (!/\d/.test(k)) return;
    if (!byR.has(k)) byR.set(k, { codigo: normR(r.r), placas: new Set(), talleres: new Set() });
    const item = byR.get(k);
    item.placas.add(r.placaNorm);
    item.talleres.add(r.taller);
  });
  const rRepetidos = [];
  for (const item of byR.values()) {
    if (item.placas.size < 2) continue;
    rRepetidos.push({
      placa: item.codigo,
      detalle: `${item.placas.size} placas: ${Array.from(item.placas).join(', ')}`,
      extra: Array.from(item.talleres).join(' · ')
    });
  }

  const plateUniverse = new Set(all.map(r => r.placaNorm));
  const rCruzado = [];
  all.forEach(r => {
    if (r.origen === 'fisico' && !r.r) return;
    if (isGarbageR(r.r).garbage) return;
    const k = rKey(r.r);
    if (!k || k === r.placaNorm) return;
    if (!plateUniverse.has(k)) return;
    if (k.length !== r.placaNorm.length) return;
    rCruzado.push({
      placa: r.placa,
      detalle: `R ${normR(r.r)} corresponde a otra placa del inventario`,
      extra: r.taller
    });
  });

  const invalidas = [];
  all.forEach(r => {
    if (isTestPlate(r.placaNorm)) return;
    if (plateFormat(r.placaNorm) === 'otro') {
      invalidas.push({
        placa: r.placa,
        detalle: `${r.origen === 'sistema' ? 'Sistema' : 'Físico'} · ${r.taller}`
      });
    }
  });

  const dedupeBy = (arr, keyFn) => {
    const seen = new Set();
    return arr.filter(item => {
      const k = keyFn(item);
      if (seen.has(k)) return false;
      seen.add(k);
      return true;
    });
  };

  /* Placas casi idénticas: distancia de edición ≤ 1 (sustitución, inserción,
     borrado o transposición adyacente). Ej. QZS756 / QSZ756. */
  const casiIdenticas = detectarCasiIdenticas(plateUniverse);

  return {
    duplicadosFisico: dup(fis),
    duplicadosSistema: dup(sys),
    casiIdenticas,
    pruebas: dedupeBy(pruebas, x => x.placa),
    rBasura: dedupeBy(rBasura, x => x.placa + '|' + x.extra),
    rRepetidos,
    rCruzado: dedupeBy(rCruzado, x => x.placa + '|' + x.detalle),
    invalidas: dedupeBy(invalidas, x => x.placa + '|' + x.detalle)
  };
}

/** true si a→b cuesta 1 edición (sustitución, inserción, borrado o
 *  transposición de caracteres adyacentes). */
function edicionUno(a, b) {
  if (a === b) return false;
  const la = a.length;
  const lb = b.length;
  if (Math.abs(la - lb) > 1) return false;
  if (la === lb) {
    let diffs = [];
    for (let i = 0; i < la; i++) {
      if (a[i] !== b[i]) diffs.push(i);
      if (diffs.length > 2) return false;
    }
    if (diffs.length === 1) return true;
    if (diffs.length === 2) {
      const [i, j] = diffs;
      return j === i + 1 && a[i] === b[j] && a[j] === b[i];
    }
    return false;
  }
  const [corta, larga] = la < lb ? [a, b] : [b, a];
  let i = 0;
  let j = 0;
  let saltado = false;
  while (i < corta.length && j < larga.length) {
    if (corta[i] === larga[j]) { i++; j++; continue; }
    if (saltado) return false;
    saltado = true;
    j++;
  }
  return true;
}

function detectarCasiIdenticas(plateUniverse) {
  const plates = Array.from(plateUniverse).filter(p => p && p.length >= 4 && !isTestPlate(p));
  const byLen = new Map();
  plates.forEach(p => {
    if (!byLen.has(p.length)) byLen.set(p.length, []);
    byLen.get(p.length).push(p);
  });
  const out = [];
  const seenPair = new Set();
  const lens = Array.from(byLen.keys()).sort((a, b) => a - b);
  lens.forEach(l => {
    const group = byLen.get(l);
    const vecinos = [];
    if (byLen.has(l)) vecinos.push(group);
    if (byLen.has(l + 1)) vecinos.push(byLen.get(l + 1));
    for (let i = 0; i < group.length; i++) {
      vecinos.forEach(other => {
        const start = other === group ? i + 1 : 0;
        for (let j = start; j < other.length; j++) {
          const a = group[i];
          const b = other[j];
          if (!edicionUno(a, b)) continue;
          const pairKey = a < b ? a + '|' + b : b + '|' + a;
          if (seenPair.has(pairKey)) continue;
          seenPair.add(pairKey);
          out.push({ placa: `${a} / ${b}`, detalle: 'Distancia de edición 1: revisa que no sea un error de digitación' });
        }
      });
    }
  });
  return out.sort((x, y) => x.placa.localeCompare(y.placa, 'es'));
}

/* ---------------------------- 6. ESTADO ---------------------------- */

const state = {
  files: { sistema: null, fisico: [] },
  importStatus: { sistema: 'empty', fisico: 'empty', message: '' },
  system: { rows: [], sections: [], warnings: [], meta: {} },
  physical: { rows: [], sections: [], warnings: [], meta: {} },
  ingresos: [],
  corteFecha: '',
  model: null,
  filters: { taller: '', tipo: '', estado: '', q: '' },
  ui: { view: 'resumen', open: new Set(), groupBy: 'sistema' },
  history: {
    cortes: [], loaded: false, loading: false, selectedId: null,
    detalles: new Map(), comparison: null, bucket: 'todos', error: ''
  },
  auth: { configured: false, user: null, client: null, error: '', mode: 'login', loading: false }
};

/* ---------------------------- 7. RENDER ---------------------------- */

function renderAll() {
  const empty = !(state.model && state.model.rows.length);
  if (typeof document !== 'undefined') document.body.classList.toggle('is-empty', empty);
  renderKpis();
  renderDonut();
  renderBars();
  renderImport();
  renderFilters();
  renderSections();
  renderAnomalies();
  renderExcluidas();
  renderSideCard();
  const nav = $('#navAnomalias');
  if (nav) nav.textContent = fmt(state.model ? state.model.anomalyCount : 0);
  const navEx = $('#navExcluidas');
  if (navEx) navEx.textContent = fmt(state.model ? state.model.stats.excluidas || 0 : 0);
  const footer = $('#footerMeta');
  if (footer) {
    footer.textContent = state.model && state.model.rows.length
      ? `${fmt(state.model.rows.length)} placas comparadas · ${fmt(state.model.stats.excluidas || 0)} excluidas · ${formatDate(todayISO())}`
      : 'Sin archivos cargados';
  }
  if (state.ui.view === 'ajustes') renderSettings();
}

function renderKpis() {
  const host = $('#kpis');
  if (!host) return;
  const st = state.model ? state.model.stats : { total: 0, sistema: 0, fisico: 0, eliminar: 0, incluir: 0, ok: 0, excluidas: 0, sistemaUnicos: 0, fisicoUnicos: 0 };
  const anom = state.model ? state.model.anomalyCount : 0;
  const sysSections = state.system.sections.length;
  const fisSections = state.physical.sections.length;
  const sheets = state.physical.meta.sheets || 0;

  const cards = [
    {
      cls: 'kpi--primary',
      label: 'Total sistema',
      value: st.sistema,
      icon: 'pdf',
      badge: st.sistema ? 'PDF' : '-',
      badgeCls: '',
      foot: st.sistema ? `${fmt(sysSections)} talleres · ${fmt(st.sistemaUnicos)} únicas` : 'Carga el PDF del sistema'
    },
    {
      cls: '',
      label: 'Total físico',
      value: st.fisico,
      icon: 'sheet',
      badge: sheets ? (sheets === 1 ? '1 hoja' : sheets + ' hojas') : '-',
      badgeCls: 'kpi__badge--muted',
      foot: st.fisico ? `${fmt(st.fisicoUnicos)} únicas · ${fmt(fisSections)} bloques` : 'Carga el Excel físico'
    },
    {
      cls: 'kpi--ok',
      label: 'Coinciden',
      value: st.OK || 0,
      icon: 'check',
      badge: st.total ? pct(st.OK || 0, st.total) + '% total' : '-',
      badgeCls: 'kpi__badge--muted',
      foot: 'En sistema y en el conteo físico'
    },
    {
      cls: 'kpi--eliminar',
      label: 'A eliminar',
      value: st.ELIMINAR || 0,
      icon: 'minus',
      badge: st.sistema ? pct(st.ELIMINAR || 0, st.sistema) + '% sist.' : '-',
      badgeCls: 'kpi__badge--red',
      foot: 'En sistema y no en físico'
    },
    {
      cls: 'kpi--incluir',
      label: 'A incluir',
      value: st.INCLUIR || 0,
      icon: 'plus',
      badge: st.fisico ? pct(st.INCLUIR || 0, st.fisico) + '% fís.' : '-',
      badgeCls: '',
      foot: 'En físico y no en sistema'
    },
    {
      cls: 'kpi--excluidas',
      label: 'Excluidas por fecha',
      value: st.excluidas || 0,
      icon: 'clock',
      badge: st.excluidas ? fmt(st.excluidasEnSistema || 0) + ' en sist.' : '-',
      badgeCls: 'kpi__badge--muted',
      foot: state.corteFecha ? `Ingresadas después del ${formatDate(state.corteFecha)}` : 'Define la fecha de corte'
    },
    {
      cls: 'kpi--anomalias',
      label: 'Anomalías',
      value: anom,
      icon: 'box',
      badge: state.model ? countAnomalyTypes() + ' tipos' : '-',
      badgeCls: 'kpi__badge--muted',
      foot: 'Duplicados, R y pruebas'
    }
  ];

  host.innerHTML = cards.map(c => `
    <article class="kpi ${c.cls}">
      <div class="kpi__top">
        <span class="kpi__icon">${ICONS[c.icon] || ICONS.box}</span>
        ${c.badge && c.badge !== '-' ? `<span class="kpi__badge ${c.badgeCls}">${esc(c.badge)}</span>` : ''}
      </div>
      <p class="kpi__label">${esc(c.label)}</p>
      <strong class="kpi__value">${fmt(c.value)}</strong>
      <span class="kpi__foot">${esc(c.foot)}</span>
    </article>
  `).join('');
}

function countAnomalyTypes() {
  if (!state.model) return 0;
  return Object.values(state.model.anomalies).filter(arr => arr.length > 0).length;
}

function renderDonut() {
  const donut = $('#donut');
  if (!donut) return;
  const st = state.model ? state.model.stats : null;
  const total = st ? st.total : 0;
  const counts = {
    ELIMINAR: st ? st.ELIMINAR : 0,
    INCLUIR: st ? st.INCLUIR : 0,
    OK: st ? st.OK : 0
  };
  const colors = { ELIMINAR: 'var(--red)', INCLUIR: 'var(--green)', OK: 'var(--c-ok)' };

  let acc = 0;
  const stops = [];
  ESTADOS.forEach(k => {
    const p = total ? (counts[k] / total) * 100 : 0;
    if (p > 0) {
      stops.push(`${colors[k]} ${acc.toFixed(2)}% ${(acc + p).toFixed(2)}%`);
      acc += p;
    }
  });
  if (acc < 99.99) stops.push(`var(--c-ok) ${acc.toFixed(2)}% 100%`);
  donut.style.background = stops.length ? `conic-gradient(${stops.join(',')})` : 'conic-gradient(var(--c-ok) 0 100%)';

  const okPct = total ? pct(counts.OK, total) : 0;
  const valueEl = $('#donutValue');
  if (valueEl) valueEl.textContent = okPct + '%';
  const chip = $('#donutChip');
  if (chip) chip.textContent = total ? `${fmt(total)} placas` : 'Sin datos';

  const legend = $('#donutLegend');
  if (legend) {
    legend.innerHTML = ESTADOS.map(k => `
      <li><i style="background:${colors[k]}"></i>
        <span>${ESTADO_LABEL[k]}</span>
        <span class="legend__pct">${total ? pct(counts[k], total) : 0}%</span>
        <b>${fmt(counts[k])}</b>
      </li>
    `).join('');
  }
}

function renderBars() {
  const host = $('#bars');
  if (!host) return;
  const meta = $('#barsMeta');
  const mode = state.ui.groupBy;
  if (meta) meta.textContent = mode === 'sistema' ? 'Agrupado por categoría del sistema' : 'Agrupado por taller físico · Eliminar + incluir';

  if (!state.model || !state.model.rows.length) {
    host.innerHTML = '<p class="bars__empty">Carga los dos archivos para ver el comparativo por taller.</p>';
    return;
  }

  const groups = groupedSections(state.model.rows, mode);
  const withPending = groups.map(g => ({ ...g, pend: g.counts.ELIMINAR + g.counts.INCLUIR }));
  const totalPending = withPending.reduce((sum, g) => sum + g.pend, 0);
  const ranked = withPending
    .filter(g => g.pend > 0)
    .sort((a, b) => b.pend - a.pend)
    .slice(0, 6);

  if (!ranked.length) {
    host.innerHTML = '<p class="bars__empty">Sin pendientes: todo coincide.</p>';
    return;
  }

  const max = Math.max(...ranked.map(g => g.pend));
  host.innerHTML = ranked.map(g => {
    const seg = ['ELIMINAR', 'INCLUIR'].map(k => {
      const w = (g.counts[k] / max) * 100;
      return w > 0 ? `<span class="seg--${k.toLowerCase()}" style="width:${w.toFixed(2)}%" title="${ESTADO_LABEL[k]}: ${g.counts[k]}"></span>` : '';
    }).join('');
    return `
      <button class="bar-row" type="button" data-bar="${esc(g.key)}" title="Filtrar por ${esc(g.label)}">
        <span class="bar-row__label">${esc(g.label)}</span>
        <span class="bar-row__track">${seg}</span>
        <span class="bar-row__value">${fmt(g.pend)}<small>${totalPending ? pct(g.pend, totalPending) : 0}%</small></span>
      </button>
    `;
  }).join('');
}

function renderImport() {
  const host = $('#importCard');
  if (!host) return;
  const hasData = state.model && state.model.rows.length > 0;
  const status = state.importStatus;
  const busy = status.sistema === 'loading' || status.fisico === 'loading';

  if (!hasData && !busy && status.sistema === 'empty' && status.fisico === 'empty') {
    host.innerHTML = `
      <div class="empty-hero">
        <h3>Empieza con el corte de esta semana</h3>
        <p>Sube el PDF del sistema y el Excel del conteo físico. Todo se procesa en tu navegador: ningún archivo se envía a un servidor.</p>
        <p class="muted small">¿Ya guardaste un corte antes? Ábrelo en <button type="button" class="link-btn" data-view="historial">Historial</button>: ahí puedes verlo y descargar el CSV con las placas a eliminar e incluir.</p>
        <div class="drops">
          ${dropHtml('sistema')}
          ${dropHtml('fisico')}
        </div>
      </div>`;
    return;
  }

  if (state.restaurado && !state.files.sistema && !(state.files.fisico || []).length) {
    const u = readJSON(LS.ultimo, {}) || {};
    host.innerHTML = `
      <div class="restore-note">
        <div>
          <strong>📋 Última comparación restaurada</strong>
          <span>${fmt(u.stats ? u.stats.total : 0)} placas · guardada el ${esc(formatDate(u.guardado))}${u.files && u.files.sistema ? ` · ${esc(u.files.sistema)} + ${esc(u.files.fisico)}` : ''}. Es solo de consulta: vuelve a cargar los archivos para actualizarla.</span>
        </div>
        <button class="btn btn--ghost btn--sm" type="button" data-descartar-restaurado>Descartar</button>
      </div>
      <div class="drops">
        ${dropHtml('sistema')}
        ${dropHtml('fisico')}
      </div>`;
    return;
  }

  const sys = state.system.meta;
  const fis = state.physical.meta;
  const warnings = (state.system.warnings || []).length + (state.physical.warnings || []).length;
  const skipped = (state.physical.meta.skipped || []);

  host.innerHTML = `
    <div class="import-card__head">
      <div>
        <h2>Archivos del corte</h2>
        <p>${busy ? 'Procesando archivos…' : 'Los datos se normalizan y comparan al instante.'}</p>
      </div>
      ${warnings ? `<button class="btn btn--ghost" type="button" data-avisos="1">${ICONS.alert}<span>${warnings} aviso${warnings === 1 ? '' : 's'}</span></button>` : ''}
      ${hasData ? '<button class="btn btn--ghost" type="button" id="btnReimport">Reemplazar</button>' : ''}
    </div>
    <div class="drops">
      ${dropHtml('sistema')}
      ${dropHtml('fisico')}
    </div>
    ${state.ingresos.length ? `
      <div class="corte-fecha-row">
        <label class="field field--inline">
          <span class="field__label">Fecha de corte del conteo físico</span>
          <input type="date" id="corteFechaInput" value="${esc(state.corteFecha || '')}" />
        </label>
        <p class="muted small">Las placas de la hoja de ingresos con fecha posterior al corte se excluyen del comparativo y se listan en «Excluidas». Hoja de ingresos: <strong>${esc(skipped.map(s => s.nombre).join(', '))}</strong> (${fmt(state.ingresos.length)} registros).</p>
      </div>` : ''}
    ${skipped.length && !state.ingresos.length ? `<p class="muted small" style="margin-top:10px">Hoja ignorada: ${esc(skipped.map(s => s.nombre).join(', '))} (${esc(skipped[0].motivo)}).</p>` : ''}
  `;

  const reimport = $('#btnReimport');
  if (reimport) reimport.addEventListener('click', openImportModal);
}

function dropHtml(kind) {
  const isSys = kind === 'sistema';
  const status = state.importStatus[kind];
  const file = isSys ? state.files.sistema : (state.files.fisico || [])[0];
  const count = isSys ? state.system.rows.length : state.physical.rows.length;
  const sections = isSys ? state.system.sections.length : state.physical.sections.length;
  const names = isSys
    ? (state.files.sistema ? state.files.sistema.name : '')
    : (state.files.fisico || []).map(f => f.name).join(' · ');

  const label = isSys ? 'Inventario del sistema' : 'Conteo físico';
  const hint = isSys ? 'PDF · ~760 placas por taller/parqueadero' : 'Excel .xlsx · una hoja por zona o bloques por columna';
  const accept = isSys ? 'application/pdf' : 'Excel / CSV';

  if (status === 'loading') {
    return `
      <div class="drop" aria-busy="true">
        <span class="drop__icon"><span class="spinner"></span></span>
        <span class="drop__text"><strong>${esc(label)}</strong><span>Leyendo ${esc(names || (isSys ? 'PDF' : 'Excel'))}…</span></span>
      </div>`;
  }

  if (status === 'error') {
    return `
      <div class="drop drop--error" data-drop="${kind}" role="button" tabindex="0">
        <span class="drop__icon">${ICONS.alert}</span>
        <span class="drop__text">
          <strong>No se pudo leer ${esc(label.toLowerCase())}</strong>
          <span>${esc(state.importStatus.message || 'Revisa el archivo e inténtalo de nuevo.')}</span>
          <em>${ICONS.down} Reintentar</em>
        </span>
      </div>`;
  }

  if (!file) {
    return `
      <div class="drop" data-drop="${kind}" role="button" tabindex="0">
        <span class="drop__icon">${isSys ? ICONS.pdf : ICONS.sheet}</span>
        <span class="drop__text">
          <strong>${esc(label)}</strong>
          <span>${esc(hint)}</span>
          <em>${ICONS.down} Arrastra o haz clic · ${accept}</em>
        </span>
      </div>`;
  }

  const extra = isSys
    ? `<span class="mini-tag">${fmt(count)} placas</span>
       <span class="mini-tag mini-tag--muted">${fmt(sections)} talleres</span>
       ${state.system.meta.pages ? `<span class="mini-tag mini-tag--muted">${state.system.meta.pages} pág.</span>` : ''}`
    : `<span class="mini-tag">${fmt(count)} registros</span>
       <span class="mini-tag mini-tag--muted">${fmt(sections)} bloques</span>`;

  return `
    <div class="drop drop--loaded" data-drop="${kind}" role="button" tabindex="0">
      <span class="drop__icon">${isSys ? ICONS.pdf : ICONS.sheet}</span>
      <span class="drop__text">
        <strong title="${esc(names)}">${esc(names.length > 46 ? names.slice(0, 44) + '…' : names)}</strong>
        <span class="drop__stats">${extra}</span>
        <em>${ICONS.down} Cambiar archivo</em>
      </span>
      <button class="drop__reset" type="button" data-reset="${kind}" title="Quitar archivo" aria-label="Quitar archivo">${ICONS.x}</button>
    </div>`;
}

/* Clave y etiqueta de agrupación de una fila según el modo.
   En modo físico, las placas del PARQUEO se agrupan por la categoría del
   sistema con la que aparecen (o "PARQUEO (sin categoría)"). */
function groupKeyOf(row, mode) {
  return mode === 'fisico'
    ? (row.fisGroupKey || row.fisKey || row.sysKey)
    : (row.sysKey || row.fisKey);
}

function groupLabelOf(row, mode) {
  return mode === 'fisico'
    ? (row.fisGroupLabel || row.fisTaller || row.sysTaller)
    : (row.sysTaller || row.fisTaller);
}

function groupedSections(rows, mode) {
  const meta = new Map();
  if (state.model) state.model.sections.forEach(s => meta.set(s.key, s));
  const map = new Map();
  rows.forEach(row => {
    const key = groupKeyOf(row, mode);
    if (!key) return;
    const label = groupLabelOf(row, mode);
    if (!map.has(key)) {
      map.set(key, {
        key,
        label: label || 'Sin taller',
        meta: meta.get(key) || null,
        counts: { ELIMINAR: 0, INCLUIR: 0, OK: 0 },
        rows: []
      });
    }
    const sec = map.get(key);
    sec.counts[row.estado] += 1;
    sec.rows.push(row);
  });
  return Array.from(map.values())
    .map(s => ({ ...s, tipo: tipoFromName(s.label), pend: s.counts.ELIMINAR + s.counts.INCLUIR }))
    .sort((a, b) => b.pend - a.pend || a.label.localeCompare(b.label, 'es'));
}

function filteredRows() {
  if (!state.model) return [];
  const f = state.filters;
  let rows = state.model.rows;
  if (f.estado) rows = rows.filter(r => r.estado === f.estado);
  if (f.tipo) rows = rows.filter(r => r.tipo === f.tipo);
  if (f.taller) rows = rows.filter(r => groupKeyOf(r, state.ui.groupBy) === f.taller);
  if (f.q) {
    const q = normPlaca(f.q);
    if (q) rows = rows.filter(r => r.placaNorm.includes(q));
  }
  return rows;
}

function renderFilters() {
  const rows = state.model ? state.model.rows : [];
  const groups = groupedSections(rows, state.ui.groupBy);
  const tallerSel = $('#filterTaller');
  const tipoSel = $('#filterTipo');
  if (!tallerSel || !tipoSel) return;

  const options = [...groups].sort((a, b) => a.label.localeCompare(b.label, 'es'));
  const wanted = state.filters.taller;
  tallerSel.innerHTML = '<option value="">Todos los talleres</option>' + options.map(g => {
    const n = g.rows.length;
    return `<option value="${esc(g.key)}">${esc(g.label)} (${n})</option>`;
  }).join('');
  tallerSel.value = options.some(g => g.key === wanted) ? wanted : '';
  if (tallerSel.value !== wanted) state.filters.taller = '';

  const tipos = Array.from(new Set(rows.map(r => r.tipo))).sort((a, b) => {
    const ia = TIPO_ORDER.indexOf(a);
    const ib = TIPO_ORDER.indexOf(b);
    return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib) || a.localeCompare(b, 'es');
  });
  const wantedTipo = state.filters.tipo;
  tipoSel.innerHTML = '<option value="">Todos los tipos</option>' + tipos.map(t => {
    const n = rows.filter(r => r.tipo === t).length;
    return `<option value="${esc(t)}">${esc(t)} (${n})</option>`;
  }).join('');
  tipoSel.value = tipos.includes(wantedTipo) ? wantedTipo : '';
  if (tipoSel.value !== wantedTipo) state.filters.tipo = '';

  const estadoSel = $('#filterEstado');
  if (estadoSel) {
    const st = state.model ? state.model.stats : null;
    const counts = {
      ELIMINAR: st ? st.ELIMINAR : 0,
      INCLUIR: st ? st.INCLUIR : 0,
      OK: st ? st.OK : 0
    };
    const wantedEstado = state.filters.estado;
    estadoSel.innerHTML = '<option value="">Todos los estados</option>' + ESTADOS.map(e =>
      `<option value="${e}">${ESTADO_LABEL[e]} (${fmt(counts[e])})</option>`
    ).join('');
    estadoSel.value = ESTADOS.includes(wantedEstado) ? wantedEstado : '';
  }

  const search = $('#searchPlaca');
  if (search && search.value !== state.filters.q) search.value = state.filters.q;

  const visibles = filteredRows();
  const inc = visibles.filter(r => r.estado === 'INCLUIR').length;
  const eli = visibles.filter(r => r.estado === 'ELIMINAR').length;
  const bi = $('#btnCopyIncluir');
  const be = $('#btnCopyEliminar');
  const bx = $('#btnExport');
  if (bi) {
    bi.disabled = inc === 0;
    const label = bi.querySelector('span');
    if (label) label.textContent = inc ? `Copiar incluir (${fmt(inc)})` : 'Copiar incluir';
  }
  if (be) {
    be.disabled = eli === 0;
    const label = be.querySelector('span');
    if (label) label.textContent = eli ? `Copiar eliminar (${fmt(eli)})` : 'Copiar eliminar';
  }
  if (bx) bx.disabled = visibles.length === 0;
  const bxx = $('#btnExportExcel');
  if (bxx) bxx.disabled = visibles.length === 0;

  const count = $('#resultCount');
  if (count) {
    count.textContent = visibles.length
      ? `${fmt(visibles.length)} placas visibles · ${fmt(new Set(visibles.map(r => groupKeyOf(r, state.ui.groupBy))).size)} talleres`
      : (state.model ? 'Ninguna placa coincide con los filtros' : 'Sin datos');
  }
}

function renderSections() {
  const host = $('#sections');
  if (!host) return;
  if (!state.model || !state.model.rows.length) {
    host.innerHTML = '<div class="sections__empty"><strong>Aún no hay comparativo</strong>Carga el PDF del sistema y el Excel físico para ver las placas agrupadas por taller.</div>';
    return;
  }

  const rows = filteredRows();
  if (!rows.length) {
    host.innerHTML = '<div class="sections__empty"><strong>Sin resultados</strong>Ajusta los filtros o limpia la búsqueda para ver todas las placas.</div>';
    return;
  }

  const groups = groupedSections(rows, state.ui.groupBy);
  const forceOpen = state.filters.q.length > 0;
  host.innerHTML = groups.map(g => sectionHtml(g, forceOpen || state.ui.open.has(g.key))).join('');
}

function sectionHtml(group, open) {
  const c = group.counts;
  const total = group.rows.length;
  const declared = group.meta && group.meta.declared;
  const sysOnly = group.key.startsWith('sys:');
  const tag = (estado) => {
    const n = c[estado];
    return `<span class="tag tag--${estado.toLowerCase()} ${n ? '' : 'tag--zero'}">${ESTADO_LABEL[estado]} ${fmt(n)}</span>`;
  };

  const note = declared && declared !== total && !state.filters.estado
    ? `<div class="section__note">${ICONS.alert}<span>El PDF declara ${fmt(declared)} placas para este bloque y se leyeron ${fmt(total)}.</span></div>`
    : '';

  return `
    <article class="section ${open ? 'is-open' : ''}" data-section="${esc(group.key)}">
      <div class="section__head">
        <button class="section__chev" type="button" data-toggle="${esc(group.key)}" aria-expanded="${open ? 'true' : 'false'}" aria-label="Mostrar u ocultar placas">
          <svg viewBox="0 0 24 24" aria-hidden="true"><path d="m9 6 6 6-6 6"/></svg>
        </button>
        <button class="section__id" type="button" data-toggle="${esc(group.key)}">
          <span class="section__name">${esc(group.label)} ${sysOnly ? '' : '<span class="origin">Físico</span>'}</span>
          <span class="section__sub">
            <span>${esc(group.tipo)}</span>
            <span>·</span>
            <span>${fmt(total)} placas</span>
            ${declared ? `<span>·</span><span>PDF declara ${fmt(declared)}</span>` : ''}
          </span>
        </button>
        <div class="section__tags">${tag('ELIMINAR')}${tag('INCLUIR')}${tag('OK')}</div>
        <div class="section__tools">
          <button class="btn btn--soft" type="button" data-copy-section="${esc(group.key)}" data-estado="INCLUIR" ${c.INCLUIR ? '' : 'disabled'} title="Copiar las placas a incluir de este taller" aria-label="Copiar placas a incluir de este taller">
            ${ICONS.copy}<span>Copiar incluir${c.INCLUIR ? ' (' + fmt(c.INCLUIR) + ')' : ''}</span>
          </button>
          <button class="icon-btn" type="button" data-copy-section="${esc(group.key)}" data-estado="ELIMINAR" ${c.ELIMINAR ? '' : 'disabled'} title="Copiar las placas a eliminar de este taller" aria-label="Copiar placas a eliminar">
            ${ICONS.minus}
          </button>
        </div>
      </div>
      ${open ? `<div class="section__body">${note}${tableHtml(group.rows)}</div>` : ''}
    </article>
  `;
}

function markScrollables(root) {
  const scope = root || document;
  const lists = [];
  if (scope.matches && scope.matches('.anomalia-list, .mapping')) lists.push(scope);
  lists.push(...$$('.anomalia-list, .mapping', scope));
  lists.forEach(el => {
    const scrollable = el.scrollHeight - el.clientHeight > 6;
    el.classList.toggle('is-scrollable', scrollable);
    el.classList.toggle('at-end', !scrollable || el.scrollTop + el.clientHeight >= el.scrollHeight - 6);
  });
}

function tableHtml(rows) {
  return `
    <p class="table-hint">Desliza la tabla para ver todas las columnas →</p>
    <div class="table-wrap">
      <table class="data">
        <thead>
          <tr>
            <th>Placa</th>
            <th>Código R</th>
            <th>Estado</th>
            <th>Categoría / Parqueadero (sistema)</th>
            <th>Taller físico</th>
            <th>Tipo</th>
            <th><span class="sr-only">Copiar</span></th>
          </tr>
        </thead>
        <tbody>
          ${rows.map(rowHtml).join('')}
        </tbody>
      </table>
    </div>`;
}

function rowHtml(row) {
  const estado = row.estado;
  const badge = `<span class="badge badge--${estado.toLowerCase()}" title="${esc(row.motivo)}"><i></i>${estado}</span>`;
  const r = row.r ? esc(row.r) : '<span class="muted-cell">—</span>';
  const canCopy = estado === 'ELIMINAR' || estado === 'INCLUIR';
  return `
    <tr class="row--${estado.toLowerCase()}">
      <td><span class="plate">${esc(row.placa)}</span></td>
      <td><span class="r-code">${r}</span></td>
      <td>${badge}</td>
      <td class="taller-cell"><span title="${esc(row.sysTaller)}">${row.sysTaller ? esc(row.sysTaller) : '<span class="muted-cell">—</span>'}</span></td>
      <td class="taller-cell"><span title="${esc(row.fisTaller)}">${row.fisTaller ? esc(row.fisTaller) : '<span class="muted-cell">—</span>'}</span></td>
      <td class="muted-cell">${esc(row.tipo)}</td>
      <td>${canCopy ? `<button class="copy-one" type="button" data-copy-placa="${esc(row.placa)}" title="Copiar placa" aria-label="Copiar ${esc(row.placa)}">${ICONS.copy}</button>` : ''}</td>
    </tr>`;
}

const ANOMALIA_DEFS = [
  { key: 'duplicadosFisico', titulo: 'Duplicados en el físico', desc: 'La misma placa aparece más de una vez en el conteo.', tone: 'red', icon: 'cards' },
  { key: 'duplicadosSistema', titulo: 'Duplicados en el sistema', desc: 'La placa está repetida en el PDF del sistema.', tone: 'red', icon: 'cards' },
  { key: 'casiIdenticas', titulo: 'Placas casi idénticas', desc: 'Pares con distancia de edición 1 (ej. QZS756 / QSZ756). Suelen ser errores de digitación.', tone: 'amber', icon: 'cards' },
  { key: 'pruebas', titulo: 'Placas de prueba', desc: 'Lista editable en Ajustes: ELO12345, OOO1234, AND1234 y similares.', tone: 'violet', icon: 'box' },
  { key: 'rBasura', titulo: 'Basura en columna R', desc: 'Vacío, Xxxx, N/V, N/A, comas o valores sin código real.', tone: 'amber', icon: 'alert' },
  { key: 'rRepetidos', titulo: 'Códigos R repetidos', desc: 'Un mismo código R asignado a varias placas.', tone: 'amber', icon: 'link' },
  { key: 'rCruzado', titulo: 'R que apunta a otra placa', desc: 'El código R coincide con la placa de otro vehículo.', tone: 'amber', icon: 'link' },
  { key: 'invalidas', titulo: 'Filas sin placa válida', desc: 'No cumplen AAA123 ni el formato de remolque (letra + 4-6 dígitos).', tone: 'violet', icon: 'box' }
];

function renderAnomalies() {
  const host = $('#anomalies');
  if (!host) return;
  if (!state.model) {
    host.innerHTML = '<div class="sections__empty"><strong>Sin datos</strong>Carga los archivos para detectar anomalías.</div>';
    return;
  }

  const a = state.model.anomalies;
  host.innerHTML = ANOMALIA_DEFS.map(def => {
    const items = a[def.key] || [];
    const tone = def.tone === 'red' ? ' anomalia-card__icon--red' : def.tone === 'violet' ? ' anomalia-card__icon--violet' : '';
    const list = items.length
      ? `<div class="anomalia-list">${items.map(it => `
          <div class="anomalia-item">
            <span class="plate" title="${esc(it.placa)}">${esc(it.placa)}</span>
            ${it.extra ? `<span class="r" title="${esc(it.extra)}">${esc(it.extra)}</span>` : ''}
            <span class="where" title="${esc(it.detalle)}">${esc(it.detalle)}</span>
          </div>`).join('')}</div>`
      : '<p class="anomalia-empty">Sin hallazgos en esta categoría.</p>';
    return `
      <article class="card anomalia-card" data-anomalia="${def.key}">
        <div class="anomalia-card__head">
          <span class="anomalia-card__icon${tone}">${ICONS[def.icon] || ICONS.alert}</span>
          <div>
            <h3>${esc(def.titulo)}</h3>
            <p>${esc(def.desc)}</p>
          </div>
          <strong class="anomalia-card__count">${fmt(items.length)}</strong>
        </div>
        ${list}
      </article>`;
  }).join('');

  const bx = $('#btnExportAnomalias');
  const bc = $('#btnCopyAnomalias');
  const total = state.model.anomalyCount;
  if (bx) bx.disabled = total === 0;
  if (bc) bc.disabled = total === 0;

  const sub = $('#anomaliasSub');
  if (sub) {
    sub.textContent = total
      ? `${fmt(total)} hallazgos en ${countAnomalyTypes()} categorías. Revisa antes de cargar cambios al sistema.`
      : 'Sin anomalías detectadas en el corte.';
  }

  markScrollables(host);
}

/* ---------------------------- VISTA: EXCLUIDAS ---------------------------- */

function renderExcluidas() {
  const host = $('#excluidasBody');
  if (!host) return;
  const excluded = state.model ? state.model.excluded : [];
  if (!state.model) {
    host.innerHTML = '<div class="sections__empty"><strong>Sin datos</strong>Carga los archivos para ver las placas excluidas por fecha.</div>';
    return;
  }
  if (!excluded.length) {
    host.innerHTML = `<div class="sections__empty"><strong>Sin placas excluidas</strong>${state.corteFecha
      ? 'Ninguna placa de la hoja de ingresos es posterior al corte del ' + esc(formatDate(state.corteFecha)) + '.'
      : 'Define la fecha de corte del conteo físico en «Archivos del corte» para excluir los ingresos posteriores.'}</div>`;
    return;
  }

  const enSistema = excluded.filter(e => e.enSistema).length;
  host.innerHTML = `
    <div class="card toolbar toolbar--tight">
      <div class="toolbar__filters">
        <div>
          <h2 class="h2">Ingresadas después del corte</h2>
          <p class="muted">Corte del ${esc(formatDate(state.corteFecha))} · ${fmt(excluded.length)} placas · ${fmt(enSistema)} ya están en el sistema.</p>
          <p class="muted small">A estas placas <strong>no se les hace nada</strong>: ya aparecen en el PDF del sistema, así que no se proponen para eliminar ni para incluir.</p>
        </div>
      </div>
      <div class="toolbar__actions">
        <button class="btn btn--dark" id="btnExportExcluidas" type="button">${ICONS.down}<span>Exportar CSV</span></button>
      </div>
    </div>
    <div class="card">
      <p class="table-hint">Desliza la tabla para ver categoría, código R y fecha →</p>
      <div class="table-wrap">
        <table class="data">
          <thead>
            <tr>
              <th>Placa</th>
              <th>Categoría (ingreso)</th>
              <th>Categoría / Parqueadero (sistema)</th>
              <th>Código R</th>
              <th>Fecha ingreso</th>
              <th>¿Está en el sistema?</th>
            </tr>
          </thead>
          <tbody>
            ${excluded.map(e => `
              <tr class="row--excluida">
                <td><span class="plate">${esc(e.placa)}</span></td>
                <td class="taller-cell"><span>${esc(e.categoria || '—')}</span></td>
                <td class="taller-cell"><span>${esc(e.catSistema || '—')}</span></td>
                <td class="r-code">${esc(e.codigoR || '—')}</td>
                <td class="muted-cell">${esc(formatDate(e.fecha, true))}</td>
                <td>${e.enSistema
                  ? `<span class="tag tag--ok">${ICONS.check} Sí</span>`
                  : '<span class="tag tag--zero">Aún no</span>'}</td>
              </tr>`).join('')}
          </tbody>
        </table>
      </div>
    </div>`;

  const btn = $('#btnExportExcluidas');
  if (btn) btn.addEventListener('click', exportExcluidasCsv);
}

function exportExcluidasCsv() {
  const excluded = state.model ? state.model.excluded : [];
  if (!excluded.length) return toast('No hay placas excluidas para exportar.', 'err');
  const data = excluded.map(e => [
    e.placa,
    e.categoria,
    e.catSistema,
    e.codigoR,
    e.fecha,
    e.enSistema ? 'Si' : 'No'
  ]);
  downloadFile(
    `BYR_excluidas_${state.corteFecha || todayISO()}.csv`,
    toCsv(['Placa', 'Categoria ingreso', 'Categoria sistema', 'Codigo R', 'Fecha ingreso', 'Esta en sistema'], data),
    'text/csv;charset=utf-8'
  );
  toast(`${fmt(excluded.length)} placas excluidas exportadas.`);
}

/* ---------------------------- 8. ACCIONES ---------------------------- */

function renderSideCard() {
  const has = state.model && state.model.rows.length > 0;
  const fecha = $('#sideCorteFecha');
  const hint = $('#sideCorteHint');
  const btn = $('#btnSaveCorte');
  if (fecha) fecha.textContent = has ? formatDate(state.corteFecha || todayISO()) : 'Sin corte cargado';
  if (hint) {
    if (!has) hint.textContent = 'Carga el PDF del sistema y el Excel físico para empezar.';
    else {
      const ex = state.model.stats.excluidas || 0;
      hint.textContent = `${fmt(state.model.stats.sistema)} en sistema · ${fmt(state.model.stats.fisico)} en físico` + (ex ? ` · ${fmt(ex)} excluidas` : '');
    }
  }
  if (btn) {
    btn.disabled = !has;
    btn.textContent = has ? 'Guardar corte' : 'Guardar corte';
  }
}

function setView(view) {
  const meta = {
    resumen: ['Resumen del corte', 'Comparativo de inventario físico vs sistema'],
    comparacion: ['Comparación', 'Placas agrupadas por categoría, taller o parqueadero'],
    anomalias: ['Anomalías', 'Duplicados, placas casi idénticas, prueba y códigos R por limpiar'],
    excluidas: ['Excluidas por fecha', 'Ingresos posteriores al corte: no se les hace nada'],
    historial: ['Historial de cortes', 'Pendientes que persisten y pendientes resueltos'],
    ajustes: ['Ajustes', 'Placas de prueba, conexión de datos y reglas']
  }[view] || ['Resumen del corte', ''];
  state.ui.view = view;
  $$('.view').forEach(v => { v.hidden = v.id !== 'view-' + view; });
  $$('.nav__item[data-view]').forEach(b => {
    const active = b.dataset.view === view;
    b.classList.toggle('is-active', active);
    if (active) b.setAttribute('aria-current', 'page');
    else b.removeAttribute('aria-current');
  });
  const title = $('#viewTitle');
  const sub = $('#viewSub');
  if (title) title.textContent = meta[0];
  if (sub) sub.textContent = meta[1];
  closeSidebar();
  window.scrollTo({ top: 0, behavior: 'smooth' });
  if (view === 'historial') renderHistory();
  if (view === 'ajustes') renderSettings();
  if (view === 'anomalias') renderAnomalies();
  if (view === 'excluidas') renderExcluidas();
  if (view === 'comparacion') { renderFilters(); renderSections(); }
}

function openSidebar() {
  const sb = $('#sidebar');
  const scrim = $('#scrim');
  if (sb) sb.classList.add('is-open');
  if (scrim) scrim.hidden = false;
  const btn = $('#menuBtn');
  if (btn) btn.setAttribute('aria-expanded', 'true');
}

function closeSidebar() {
  const sb = $('#sidebar');
  const scrim = $('#scrim');
  if (sb) sb.classList.remove('is-open');
  if (scrim) scrim.hidden = true;
  const btn = $('#menuBtn');
  if (btn) btn.setAttribute('aria-expanded', 'false');
}

/* ------------------------------ IMPORTACIÓN ------------------------------ */

function pickFile(kind) {
  const input = kind === 'sistema' ? $('#fileSistema') : $('#fileFisico');
  if (!input) return;
  input.value = '';
  input.click();
}

function resetImport(kind) {
  if (kind === 'sistema') {
    state.files.sistema = null;
    state.system = { rows: [], sections: [], warnings: [], meta: {} };
    state.importStatus.sistema = 'empty';
  } else {
    state.files.fisico = [];
    state.physical = { rows: [], sections: [], warnings: [], meta: {} };
    state.ingresos = [];
    state.corteFecha = '';
    state.importStatus.fisico = 'empty';
  }
  state.importStatus.message = '';
  rebuild();
}

async function handleFileList(kind, fileList) {
  const files = Array.from(fileList || []);
  if (!files.length) return;
  if (kind === 'sistema') {
    const pdf = files.find(f => /\.pdf$/i.test(f.name) || f.type === 'application/pdf');
    if (!pdf) return toast('El inventario del sistema debe ser un archivo PDF.', 'err');
    await loadSystemFile(pdf);
    return;
  }
  const valid = files.filter(f => /\.(xlsx|xls|csv)$/i.test(f.name) || /sheet|excel/i.test(f.type));
  if (!valid.length) return toast('El conteo físico debe ser Excel (.xlsx / .xls) o CSV.', 'err');
  await loadPhysicalFile(valid);
}

async function loadSystemFile(file) {
  state.importStatus.sistema = 'loading';
  state.importStatus.message = '';
  renderImport();
  try {
    const pages = await readPdfFile(file);
    const parsed = parseSystemLines(pages);
    if (!parsed.rows.length) throw new Error('No se detectaron placas. Verifica que el PDF sea el inventario del sistema.');
    state.files.sistema = file;
    state.system = {
      rows: parsed.rows,
      sections: parsed.sections,
      warnings: parsed.warnings,
      meta: { pages: pages.length, declaredTotal: parsed.declaredTotal, file: file.name }
    };
    state.importStatus.sistema = 'ok';
    state.ui.open.clear();
    rebuild();
    toast(`${fmt(parsed.rows.length)} placas leídas del PDF.`);
    if (parsed.warnings.length) {
      toast(`${parsed.warnings.length} avisos de lectura en el PDF.`, 'err');
    }
    maybeAskCorteFecha();
  } catch (err) {
    console.error(err);
    state.importStatus.sistema = 'error';
    state.importStatus.message = err.message || 'No se pudo leer el PDF.';
    state.files.sistema = null;
    renderImport();
    toast('No se pudo leer el PDF del sistema.', 'err');
  }
}

async function loadPhysicalFile(files) {
  state.importStatus.fisico = 'loading';
  state.importStatus.message = '';
  renderImport();
  try {
    const parsed = await readPhysicalFiles(files);
    if (!parsed.rows.length) throw new Error('No se detectaron placas en el Excel. Revisa el formato.');
    state.files.fisico = files;
    state.physical = {
      rows: parsed.rows,
      sections: parsed.sections,
      warnings: parsed.warnings,
      meta: { sheets: parsed.sheetStats.length, skipped: parsed.skipped, file: files.map(f => f.name).join(' · ') }
    };
    state.ingresos = parsed.ingresos || [];
    state.corteFecha = '';
    state.importStatus.fisico = 'ok';
    state.ui.open.clear();
    rebuild();
    toast(`${fmt(parsed.rows.length)} placas leídas del conteo físico.`);
    if (state.ingresos.length) maybeAskCorteFecha();
  } catch (err) {
    console.error(err);
    state.importStatus.fisico = 'error';
    state.importStatus.message = err.message || 'No se pudo leer el Excel.';
    state.files.fisico = [];
    state.ingresos = [];
    state.corteFecha = '';
    renderImport();
    toast('No se pudo leer el Excel físico.', 'err');
  }
}

/* ---- Fecha de corte (§4): se pide al cargar los archivos cuando existe una
   hoja de ingresos; su valor es editable en «Archivos del corte». ---- */
function defaultCorteFecha() {
  const fechas = Array.from(new Set((state.ingresos || []).map(i => i.fecha).filter(Boolean))).sort();
  if (!fechas.length) return todayISO();
  /* Si el log cubre varios días (la semana completa), el conteo físico se hizo
     el último día: sugerir ese día. Si trae un solo día, es un log de ingresos
     posteriores al conteo: sugerir un día antes (comportamiento original). */
  const target = fechas.length > 1 ? fechas[fechas.length - 1] : fechas[0];
  const d = new Date(target + 'T12:00:00');
  if (fechas.length === 1) d.setDate(d.getDate() - 1);
  const off = d.getTimezoneOffset();
  return new Date(d.getTime() - off * 60000).toISOString().slice(0, 10);
}

function applyCorteFecha(fecha) {
  state.corteFecha = fecha || '';
  state.ui.open.clear();
  rebuild();
  const n = state.model ? state.model.stats.excluidas || 0 : 0;
  if (state.corteFecha) {
    toast(fecha ? `Fecha de corte: ${formatDate(fecha)} · ${fmt(n)} placa${n === 1 ? '' : 's'} excluida${n === 1 ? '' : 's'}.` : 'Sin fecha de corte: no se excluyen ingresos.');
  }
}

function maybeAskCorteFecha() {
  if (!state.ingresos.length || state.corteFecha) return;
  if (!state.system.rows.length) return; // se pregunta también al cargar el PDF
  askCorteFecha();
}

function askCorteFecha() {
  const def = defaultCorteFecha();
  const rango = (() => {
    const fechas = Array.from(new Set((state.ingresos || []).map(i => i.fecha).filter(Boolean))).sort();
    return fechas.length > 1 ? [fechas[0], fechas[fechas.length - 1]] : null;
  })();
  openModal({
    title: 'Fecha de corte del conteo físico',
    body: `
      <p>El Excel trae una hoja de ingresos con ${fmt(state.ingresos.length)} registros. Indica la fecha del conteo físico: toda placa ingresada <strong>después</strong> de esa fecha se excluye del comparativo (no se elimina ni se incluye) y se lista en la vista «Excluidas».</p>
      <label class="field" style="margin-top:10px">
        <span class="field__label">Fecha de corte</span>
        <input type="date" id="askCorteFechaInput" value="${esc(def)}" />
      </label>
      <p class="muted small">Sugerencia: ${esc(formatDate(def))}${rango ? ` · el log trae ingresos del ${esc(formatDate(rango[0]))} al ${esc(formatDate(rango[1]))} y el conteo se hizo el último día` : ' (un día después del único ingreso del log)'}. Revísala: debe ser el día en que se hizo el conteo físico.</p>
    `,
    foot: `
      <button class="btn btn--ghost" type="button" data-fecha="skip">Sin exclusiones</button>
      <button class="btn btn--primary" type="button" data-fecha="ok">Confirmar fecha</button>
    `,
    onMount: (_body, foot) => {
      foot.querySelector('[data-fecha="ok"]').addEventListener('click', () => {
        const value = ($('#askCorteFechaInput') || {}).value || def;
        closeModal();
        applyCorteFecha(value);
      });
      foot.querySelector('[data-fecha="skip"]').addEventListener('click', () => {
        closeModal();
        applyCorteFecha('');
      });
    }
  });
}

function rebuild() {
  const hasFiles = state.system.rows.length || state.physical.rows.length;
  if (!hasFiles) {
    // Sin archivos: conserva la comparación restaurada si la hay.
    if (state.restaurado && state.model) {
      saveUltimo();
      renderAll();
      return;
    }
    state.model = null;
  } else {
    state.model = buildModel();
    state.restaurado = false;
    saveUltimo();
  }
  renderAll();
}

/* -------- Recordatorio de la última comparación (§ contra recargas) --------
   Los archivos nunca se suben a un servidor, así que no sobreviven a un
   recargado. Lo que SÍ se guarda en este navegador es el RESULTADO completo
   (placas, estados, anomalías y totales): al volver, la app lo restaura en
   modo consulta hasta que se carguen archivos nuevos. */

function saveUltimo() {
  if (!state.model || !state.model.rows.length) return;
  const compact = state.model.rows.map(r => [r.placaNorm, r.estado, r.sysTaller || '', r.fisTaller || '', r.tipo || '', r.r || '']);
  writeJSON(LS.ultimo, {
    guardado: new Date().toISOString(),
    corteFecha: state.corteFecha || '',
    stats: state.model.stats,
    anomalyCount: state.model.anomalyCount,
    anomalies: state.model.anomalies,
    rows: compact,
    files: {
      sistema: state.files.sistema ? state.files.sistema.name : '',
      fisico: (state.files.fisico || []).map(f => f.name).join(' · ')
    }
  });
}

function restoreUltimo() {
  const u = readJSON(LS.ultimo, null);
  if (!u || !Array.isArray(u.rows) || !u.rows.length) return false;
  try {
    const rows = u.rows.map(a => ({
      placa: a[0],
      placaNorm: a[0],
      r: a[5] || '',
      estado: a[1],
      motivo: 'Restaurado de la última comparación',
      sysTaller: a[2] || '',
      fisTaller: a[3] || '',
      fisGroupKey: a[3] ? 'fis:' + a[3] : '',
      fisGroupLabel: a[3] || '',
      sysKey: a[2] ? 'sys:' + a[2] : '',
      fisKey: '',
      tipo: a[4] || 'Otro',
      origen: 'restaurado'
    }));
    state.model = {
      rows,
      stats: u.stats,
      sysBy: new Map(),
      fisBy: new Map(),
      excluded: [],
      sections: buildSectionIndex(rows),
      anomalies: u.anomalies || {},
      anomalyCount: u.anomalyCount || 0,
      restaurado: true
    };
    state.corteFecha = u.corteFecha || '';
    state.restaurado = true;
    state.ui.open.clear();
    return true;
  } catch (e) {
    console.error(e);
    return false;
  }
}

function discardUltimo() {
  localStorage.removeItem(LS.ultimo);
  state.model = null;
  state.restaurado = false;
  state.ui.open.clear();
  renderAll();
  toast('Comparación restaurada descartada.');
}

/* ------------------------------ PORTAPAPELES ------------------------------ */

/* Seguridad operativa (§5): confirmación con la cantidad antes de copiar o
   exportar la lista de ELIMINAR. */
async function confirmEliminar(n, verbo) {
  if (!n) return true;
  return confirmModal({
    title: `${verbo} ${fmt(n)} placa${n === 1 ? '' : 's'} a eliminar`,
    message: `Vas a ${verbo.toLowerCase()} ${fmt(n)} placas marcadas como ELIMINAR (están en el sistema y no aparecen en el conteo físico). Revisa que el conteo esté completo antes de continuar.`,
    confirmLabel: 'Sí, continuar'
  });
}

async function copyPlates(plates, label) {
  const unique = Array.from(new Set(plates.filter(Boolean))).sort((a, b) => a.localeCompare(b, 'es'));
  if (!unique.length) return toast('No hay placas para copiar con los filtros actuales.', 'err');
  const ok = await copyText(unique.join('\n'));
  if (ok) toast(`${fmt(unique.length)} placas copiadas${label ? ' · ' + label : ''}.`);
  else toast('El navegador bloqueó el portapapeles. Usa Exportar CSV.', 'err');
}

function rowsForSection(key, estado) {
  return filteredRows().filter(r => groupKeyOf(r, state.ui.groupBy) === key && (!estado || r.estado === estado));
}

/* --------------------------------- CSV --------------------------------- */

function csvCell(value) {
  const s = String(value == null ? '' : value);
  return /[";\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}

function toCsv(header, rows) {
  const lines = [header.map(csvCell).join(';')];
  rows.forEach(r => lines.push(r.map(csvCell).join(';')));
  return '\uFEFF' + lines.join('\r\n');
}

function exportFilteredRows(rows, ext) {
  const data = rows.map(r => [
    ESTADO_LABEL[r.estado] || r.estado,
    r.placa,
    r.r,
    r.sysTaller,
    r.fisTaller,
    r.tipo,
    r.motivo
  ]);
  const header = ['Estado', 'Placa', 'Codigo R', 'Categoria sistema', 'Taller fisico', 'Tipo', 'Observacion'];
  if (ext === 'xlsx') exportRowsExcel(`BYR_depuracion_${todayISO()}.xlsx`, header, data, rows.length);
  else {
    downloadFile(
      `BYR_depuracion_${todayISO()}.csv`,
      toCsv(header, data),
      'text/csv;charset=utf-8'
    );
    toast(`${fmt(rows.length)} filas exportadas.`);
  }
}

async function exportRowsExcel(filename, header, data, count) {
  try {
    await loadScript(CDN.xlsx);
    const XLSX = globalThis.XLSX;
    if (!XLSX) throw new Error('SheetJS no está disponible.');
    const ws = XLSX.utils.aoa_to_sheet([header, ...data]);
    ws['!cols'] = header.map((h, i) => ({ wch: Math.min(38, Math.max(h.length + 2, ...data.slice(0, 200).map(r => String(r[i] || '').length + 2))) }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Depuracion');
    XLSX.writeFile(wb, filename);
    toast(`${fmt(count)} filas exportadas a Excel.`);
  } catch (err) {
    console.error(err);
    toast(err.message || 'No se pudo generar el Excel. Usa Exportar CSV.', 'err');
  }
}

async function exportFiltered(kind) {
  const rows = filteredRows();
  if (!rows.length) return toast('No hay filas para exportar.', 'err');
  const nEli = rows.filter(r => r.estado === 'ELIMINAR').length;
  if (nEli && !(await confirmEliminar(nEli, 'Exportar'))) return;
  exportFilteredRows(rows, kind);
}

async function exportFilteredCsv() {
  return exportFiltered('csv');
}

async function exportFilteredExcel() {
  return exportFiltered('xlsx');
}

function exportAnomaliasCsv() {
  if (!state.model) return;
  const rows = [];
  ANOMALIA_DEFS.forEach(def => {
    (state.model.anomalies[def.key] || []).forEach(it => {
      rows.push([def.titulo, it.placa, it.extra || '', it.detalle || '']);
    });
  });
  if (!rows.length) return toast('No hay anomalías para exportar.', 'err');
  downloadFile(
    `BYR_anomalias_${todayISO()}.csv`,
    toCsv(['Categoria', 'Placa o R', 'Detalle', 'Ubicacion'], rows),
    'text/csv;charset=utf-8'
  );
  toast(`${fmt(rows.length)} anomalías exportadas.`);
}

function copyAllAnomalias() {
  if (!state.model) return;
  const plates = [];
  ANOMALIA_DEFS.forEach(def => {
    (state.model.anomalies[def.key] || []).forEach(it => {
      if (plateFormat(it.placa) !== 'otro') plates.push(normPlaca(it.placa));
    });
  });
  copyPlates(plates, 'anomalías');
}

/* -------------------------------- MODAL -------------------------------- */

function openModal({ title, body, foot, onMount }) {
  const modal = $('#modal');
  if (!modal) return;
  $('#modalTitle').textContent = title || '';
  $('#modalBody').innerHTML = body || '';
  $('#modalFoot').innerHTML = foot || '';
  modal.hidden = false;
  document.body.style.overflow = 'hidden';
  const focusable = modal.querySelector('input, select, textarea, button');
  if (focusable) setTimeout(() => focusable.focus(), 40);
  if (typeof onMount === 'function') onMount($('#modalBody'), $('#modalFoot'));
}

function closeModal() {
  const modal = $('#modal');
  if (!modal) return;
  modal.hidden = true;
  document.body.style.overflow = '';
  if (state.ui.modalResolve) {
    const resolve = state.ui.modalResolve;
    state.ui.modalResolve = null;
    resolve(false);
  }
}

function confirmModal({ title, message, confirmLabel, danger }) {
  return new Promise(resolve => {
    state.ui.modalResolve = resolve;
    openModal({
      title,
      body: `<p>${esc(message)}</p>`,
      foot: `<button class="btn btn--ghost" type="button" data-confirm="no">Cancelar</button>
             <button class="btn ${danger ? 'btn--danger' : 'btn--primary'}" type="button" data-confirm="yes">${esc(confirmLabel || 'Confirmar')}</button>`,
      onMount: (_body, foot) => {
        foot.querySelector('[data-confirm="no"]').addEventListener('click', () => closeModal());
        foot.querySelector('[data-confirm="yes"]').addEventListener('click', () => {
          state.ui.modalResolve = null;
          closeModal();
          resolve(true);
        });
      }
    });
  });
}

function showHelp() {
  openModal({
    title: 'Cómo funciona la depuración',
    body: `
      <h3>1. Carga los dos archivos</h3>
      <ul class="bullets">
        <li><strong>PDF del sistema:</strong> cada bloque «CATEGORÍA (n)» se toma como sección; debajo vienen las placas y su código R.</li>
        <li><strong>Excel físico:</strong> acepta bloques con título por columna, hojas sin títulos como PARQUEO y tablas con encabezados PLACA / TALLER / R. La hoja de ingresos se detecta por tener fechas, sin importar su nombre.</li>
      </ul>
      <h3>2. Fecha de corte</h3>
      <p>Al cargar se pide la <strong>fecha de corte</strong> del conteo físico. Las placas ingresadas después de esa fecha se excluyen del comparativo: no se les hace nada (ya están en el sistema) y se listan en la vista «Excluidas».</p>
      <h3>3. Reglas de comparación</h3>
      <ul class="bullets">
        <li><strong>Eliminar</strong> (fondo rojo, texto blanco): está en el sistema y no en el físico.</li>
        <li><strong>Incluir</strong> (fondo verde): está en el físico y no en el sistema.</li>
        <li><strong>OK</strong>: está en ambos. Las categorías del sistema y los talleres físicos no se comparan entre sí.</li>
      </ul>
      <p class="muted small">Las placas se normalizan antes de comparar: mayúsculas, sin espacios, guiones ni puntos. El código R nunca se usa para emparejar. Antes de copiar o exportar la lista de eliminar se pide confirmación.</p>
      <h3>4. PARQUEO</h3>
      <p>La cuadrícula PARQUEO no tiene nombres: cada placa toma como parqueadero la categoría con la que aparece en el sistema, o «PARQUEO (sin categoría)» si no existe allí. Los títulos que son solo «TALLER» se agrupan como «TALLER (sin nombre)».</p>
      <h3>5. Guardar el corte</h3>
      <p>«Guardar corte» archiva los conteos y el detalle del comparativo. En Historial verás qué pendientes del corte anterior siguen abiertos y cuáles se resolvieron. No se pueden guardar dos cortes con la misma fecha sin reemplazar el anterior.</p>
      <p class="muted small">Los archivos se procesan en el navegador. Solo se envían a Supabase los totales y el detalle del corte cuando guardas.</p>
    `,
    foot: '<button class="btn btn--primary" type="button" data-close-modal>Cerrar</button>',
    onMount: (_b, foot) => foot.querySelector('[data-close-modal]').addEventListener('click', closeModal)
  });
}

function showWarnings() {
  const sys = state.system.warnings || [];
  const fis = state.physical.warnings || [];
  const skipped = (state.physical.meta && state.physical.meta.skipped) || [];
  const block = (titulo, items) => items.length
    ? `<h3>${esc(titulo)} (${items.length})</h3><ul class="bullets">${items.map(w => `<li>${esc(w.mensaje)}</li>`).join('')}</ul>`
    : '';
  openModal({
    title: 'Avisos de lectura',
    body: `
      <p class="muted small">Estos avisos no bloquean el comparativo, pero conviene revisarlos para que los totales cuadren con el documento original.</p>
      ${block('PDF del sistema', sys)}
      ${block('Excel físico', fis)}
      ${skipped.length ? `<h3>Hojas ignoradas</h3><ul class="bullets">${skipped.map(s => `<li><strong>${esc(s.nombre)}</strong>: ${esc(s.motivo)}</li>`).join('')}</ul>` : ''}
      ${!sys.length && !fis.length && !skipped.length ? '<p>No hay avisos.</p>' : ''}
    `,
    foot: '<button class="btn btn--primary" type="button" data-close-modal>Cerrar</button>',
    onMount: (_b, foot) => foot.querySelector('[data-close-modal]').addEventListener('click', closeModal)
  });
}

/* --------------------- 9. HISTORIAL Y AUTENTICACIÓN --------------------- */

function getConfig() {
  const stored = readJSON(LS.supabase, null);
  const injected = (typeof window !== 'undefined' && window.BYR_SUPABASE) || SUPABASE_DEFAULT;
  const url = (stored && stored.url) || injected.url || '';
  const key = (stored && (stored.key || stored.anonKey)) || injected.anonKey || '';
  return { url: String(url).trim(), key: String(key).trim() };
}

function isSupabaseConfigured() {
  const c = getConfig();
  return Boolean(c.url && c.key);
}

async function getSupabase() {
  const cfg = getConfig();
  if (!cfg.url || !cfg.key) return null;
  if (state.auth.client) return state.auth.client;
  const mod = await import(CDN.supabase);
  const client = mod.createClient(cfg.url, cfg.key, {
    auth: { persistSession: true, autoRefreshToken: true, detectSessionInUrl: true }
  });
  client.auth.onAuthStateChange((_event, session) => {
    state.auth.user = session && session.user ? session.user : null;
    state.history.loaded = false;
    state.history.cortes = [];
    state.history.selectedId = null;
    state.history.detalles = [];
    state.history.comparison = null;
    syncIdentity();
    if (state.ui.view === 'historial') renderHistory();
  });
  state.auth.client = client;
  state.auth.configured = true;
  return client;
}

/* Puente de sesión del portal: si el usuario ya inició sesión en
   problemacero.com (cookie de Supabase legible por JS), la app lo reconoce
   y no le pide entrar de nuevo — igual que las micro-apps del portal. */
async function bridgePortalSession(sb) {
  try {
    const url = getConfig().url;
    const ref = String(url || '').replace(/^https?:\/\//, '').split('.')[0];
    if (!ref) return false;
    const base = 'sb-' + ref + '-auth-token';

    const readCookie = name => {
      const raw = document.cookie.split('; ').find(c => c.startsWith(name + '='));
      return raw ? decodeURIComponent(raw.split('=').slice(1).join('=')) : null;
    };

    let value = readCookie(base);
    if (!value) {
      // Sesión partida en chunks (sb-{ref}-auth-token.0, .1, ...)
      const chunks = [];
      for (let i = 0; i < 40; i++) {
        const c = readCookie(base + '.' + i);
        if (c == null) break;
        chunks.push(c);
      }
      if (chunks.length) value = chunks.join('');
    }
    if (!value) return false;

    let session = null;
    if (value.startsWith('base64-')) value = value.slice(7);
    try {
      session = JSON.parse(value);
    } catch (e) {
      let b64 = value.replace(/-/g, '+').replace(/_/g, '/');
      while (b64.length % 4) b64 += '=';
      session = JSON.parse(decodeURIComponent(Array.prototype.map.call(atob(b64), c =>
        '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2)).join('')));
    }
    if (!session || !session.access_token || !session.refresh_token) return false;

    const { data, error } = await sb.auth.setSession({
      access_token: session.access_token,
      refresh_token: session.refresh_token
    });
    if (error || !data || !data.user) return false;
    return true;
  } catch (e) {
    return false;
  }
}

async function initAuth() {
  if (!isSupabaseConfigured()) {
    state.auth.configured = false;
    syncIdentity();
    return;
  }
  try {
    const sb = await getSupabase();
    if (!sb) return;
    const { data } = await sb.auth.getSession();
    let session = data && data.session ? data.session : null;
    if (!session) {
      const bridged = await bridgePortalSession(sb);
      if (bridged) {
        const again = await sb.auth.getSession();
        session = again && again.data ? again.data.session : null;
      }
    }
    state.auth.user = session ? session.user : null;
    state.auth.configured = true;
    syncIdentity();
    if (state.ui.view === 'historial') renderHistory();
  } catch (err) {
    console.error(err);
    state.auth.error = err.message || 'No se pudo conectar con Supabase.';
    state.auth.configured = false;
    syncIdentity();
    toast('Revisa la URL y la llave de Supabase en Ajustes.', 'err');
  }
}

function initialsOf(email) {
  const s = String(email || 'BY').split('@')[0].replace(/[^A-Za-z0-9]/g, '');
  return (s.slice(0, 2) || 'BY').toUpperCase();
}

function syncIdentity() {
  const user = state.auth.user;
  const configured = isSupabaseConfigured();
  const email = user && user.email ? user.email : '';
  const nameEl = $('#sideUserEmail');
  const roleEl = $('#sideUserRole');
  const avatar = $('#sideAvatar');
  const topAvatar = $('#topAvatar');
  const logout = $('#btnLogout');
  if (nameEl) nameEl.textContent = email || (configured ? 'Sin sesión' : 'Modo local');
  if (roleEl) roleEl.textContent = email ? 'Supabase · sesión activa' : (configured ? 'Entra en Historial' : 'Sin sesión de Supabase');
  if (avatar) avatar.textContent = initialsOf(email);
  if (topAvatar) topAvatar.title = email || 'Representaciones BYR SAS';
  if (logout) logout.hidden = !email;
  if (state.ui.view === 'ajustes') {
    const block = $('#authBlock');
    if (block) block.innerHTML = authBlockHtml();
  }
}

function authBlockHtml() {
  if (!isSupabaseConfigured()) {
    return '<p class="muted small">Configura la URL y la llave para activar el historial compartido con autenticación.</p>';
  }
  const user = state.auth.user;
  if (user && user.email) {
    return `<strong>Sesión activa</strong><p>${esc(user.email)}</p>
      <button class="btn btn--ghost" type="button" data-auth="logout" style="margin-top:8px">Cerrar sesión</button>`;
  }
  return `<strong>Sin sesión</strong><p>El historial requiere autenticación. Entra desde la vista Historial o crea una cuenta con tu correo.</p>
    <button class="btn btn--soft" type="button" data-auth="go" style="margin-top:8px">Ir a Historial</button>`;
}

/* --------------------------- LECTURA/ESCRITURA --------------------------- */

function sortCortes(list) {
  return [...list].sort((a, b) => {
    const fa = String(a.fecha_corte || '');
    const fb = String(b.fecha_corte || '');
    if (fa !== fb) return fb.localeCompare(fa);
    return String(b.created_at || '').localeCompare(String(a.created_at || ''));
  });
}

async function loadCortes() {
  if (state.history.loading) return;
  state.history.loading = true;
  state.history.error = '';
  try {
    if (!isSupabaseConfigured()) {
      state.history.cortes = sortCortes(readJSON(LS.cortes, []));
    } else {
      if (!state.auth.user) state.history.cortes = [];
      else {
        const sb = await getSupabase();
        const { data, error } = await sb
          .from('cortes')
          .select('*')
          .order('fecha_corte', { ascending: false })
          .order('created_at', { ascending: false });
        if (error) throw error;
        state.history.cortes = data || [];
      }
    }
  } catch (err) {
    console.error(err);
    state.history.error = err.message || 'No se pudo cargar el historial.';
    state.history.cortes = [];
  } finally {
    state.history.loading = false;
    state.history.loaded = true;
  }
}

function normDetalle(d) {
  return {
    placa: d.placa ? String(d.placa) : '',
    placaNorm: d.placa_norm ? String(d.placa_norm) : normPlaca(d.placa),
    estado: d.estado || 'OK',
    tallerSistema: d.taller_sistema || '',
    tallerFisico: d.taller_fisico || '',
    tipo: d.tipo || '',
    r: d.codigo_r || '',
    observacion: d.observacion || ''
  };
}

async function fetchDetalles(corte) {
  if (!isSupabaseConfigured()) {
    return (corte.detalles || []).map(normDetalle).filter(d => d.estado !== 'OK' && d.estado !== 'EXCLUIDA');
  }
  const sb = await getSupabase();
  if (!sb) return [];
  let all = [];
  let from = 0;
  for (;;) {
    const { data, error } = await sb
      .from('cortes_detalle')
      .select('placa,placa_norm,estado,taller_sistema,taller_fisico,tipo,codigo_r,observacion')
      .eq('corte_id', corte.id)
      .range(from, from + 999);
    if (error) throw error;
    all = all.concat(data || []);
    if (!data || data.length < 1000) break;
    from += 1000;
  }
  return all.map(normDetalle).filter(d => d.estado !== 'OK' && d.estado !== 'EXCLUIDA');
}

function compareCortes(prev, curr) {
  const prevMap = new Map(prev.map(d => [d.placaNorm, d]));
  const currMap = new Map(curr.map(d => [d.placaNorm, d]));
  const persistentes = [];
  const resueltos = [];
  const nuevos = [];

  prevMap.forEach(d => {
    if (currMap.has(d.placaNorm)) {
      const c = currMap.get(d.placaNorm);
      persistentes.push({
        ...c,
        estadoPrev: d.estado,
        estadoCurr: c.estado,
        bucket: 'persistente',
        cambio: d.estado !== c.estado
      });
    } else {
      resueltos.push({ ...d, estadoPrev: d.estado, estadoCurr: '—', bucket: 'resuelto', cambio: false });
    }
  });
  currMap.forEach(d => {
    if (!prevMap.has(d.placaNorm)) {
      nuevos.push({ ...d, estadoPrev: '—', estadoCurr: d.estado, bucket: 'nuevo', cambio: false });
    }
  });

  const sort = arr => arr.sort((a, b) => a.placaNorm.localeCompare(b.placaNorm, 'es'));
  return { persistentes: sort(persistentes), resueltos: sort(resueltos), nuevos: sort(nuevos) };
}

/* -------------------------------- MODAL -------------------------------- */

function openSaveModal() {
  if (!state.model || !state.model.rows.length) {
    return toast('Carga el PDF y el Excel antes de guardar el corte.', 'err');
  }
  if (isSupabaseConfigured() && !state.auth.user) {
    toast('Inicia sesión para guardar el corte en Supabase.', 'err');
    return setView('historial');
  }
  const st = state.model.stats;
  const destino = isSupabaseConfigured() ? 'Supabase' : 'este navegador (modo local)';
  openModal({
    title: 'Guardar corte semanal',
    body: `
      <p class="muted small">Se guardarán el resumen y las ${fmt(state.model.rows.length)} placas comparadas (+${fmt(st.excluidas || 0)} excluidas) en ${esc(destino)}.</p>
      <div class="summary-grid">
        <div><span>Sistema</span><strong>${fmt(st.sistema)}</strong></div>
        <div><span>Físico</span><strong>${fmt(st.fisico)}</strong></div>
        <div><span>Eliminar</span><strong>${fmt(st.ELIMINAR)}</strong></div>
        <div><span>Incluir</span><strong>${fmt(st.INCLUIR)}</strong></div>
        <div><span>OK</span><strong>${fmt(st.OK)}</strong></div>
        <div><span>Excluidas</span><strong>${fmt(st.excluidas || 0)}</strong></div>
        <div><span>Anomalías</span><strong>${fmt(state.model.anomalyCount)}</strong></div>
      </div>
      <div class="form" style="margin-top:12px">
        <label class="field">
          <span class="field__label">Fecha del corte</span>
          <input type="date" id="corteFecha" value="${todayISO()}" />
        </label>
        <label class="field">
          <span class="field__label">Notas (opcional)</span>
          <textarea id="corteNotas" maxlength="800" placeholder="Novedades del conteo, placas pendientes por verificar…"></textarea>
        </label>
      </div>
    `,
    foot: `
      <button class="btn btn--ghost" type="button" data-corte="cancel">Cancelar</button>
      <button class="btn btn--primary" type="button" data-corte="save">Guardar corte</button>
    `,
    onMount: (_body, foot) => {
      foot.querySelector('[data-corte="cancel"]').addEventListener('click', closeModal);
      foot.querySelector('[data-corte="save"]').addEventListener('click', saveCorteFromModal);
    }
  });
}

async function saveCorteFromModal() {
  const fecha = $('#corteFecha') ? $('#corteFecha').value : todayISO();
  const notas = $('#corteNotas') ? $('#corteNotas').value.trim() : '';
  if (!fecha) return toast('Indica la fecha del corte.', 'err');
  const btn = $('#modalFoot [data-corte="save"]');
  if (btn) { btn.disabled = true; btn.textContent = 'Guardando…'; }

  const st = state.model.stats;
  const header = {
    fecha_corte: fecha,
    archivo_sistema: state.files.sistema ? state.files.sistema.name : '',
    archivo_fisico: (state.files.fisico || []).map(f => f.name).join(' · '),
    total_sistema: st.sistema,
    total_fisico: st.fisico,
    total_eliminar: st.ELIMINAR,
    total_incluir: st.INCLUIR,
    total_excluidas: st.excluidas || 0,
    total_ok: st.OK,
    total_anomalias: state.model.anomalyCount,
    notas
  };
  const detalles = state.model.rows.map(r => ({
    placa: r.placa,
    placa_norm: r.placaNorm,
    estado: r.estado,
    taller_sistema: r.sysTaller,
    taller_fisico: r.fisTaller,
    tipo: r.tipo,
    codigo_r: r.r,
    observacion: r.motivo
  })).concat((state.model.excluded || []).map(e => ({
    placa: e.placa,
    placa_norm: e.placaNorm,
    estado: 'EXCLUIDA',
    taller_sistema: e.catSistema || '',
    taller_fisico: e.categoria || '',
    tipo: '',
    codigo_r: e.codigoR || '',
    observacion: `Ingresada el ${e.fecha} (después del corte)${e.enSistema ? ' · ya está en el sistema' : ''}`
  })));

  try {
    // §8: no guardar dos veces el mismo fecha_corte — avisar y permitir reemplazar.
    const dup = await findCorteByFecha(fecha);
    if (dup) {
      const reemplazar = await confirmModal({
        title: 'Ya existe un corte de esa fecha',
        message: `Hay un corte guardado del ${formatDate(fecha)}. Si continúas, ese corte se reemplaza por el actual.`,
        confirmLabel: 'Reemplazar corte',
        danger: true
      });
      if (!reemplazar) {
        if (btn) { btn.disabled = false; btn.textContent = 'Guardar corte'; }
        return;
      }
      await deleteCorteById(dup.id, { silent: true });
    }
    await persistCorte(header, detalles);
    closeModal();
    state.history.loaded = false;
    state.history.selectedId = null;
    state.history.detalles = [];
    state.history.comparison = null;
    toast(`Corte del ${formatDate(fecha)} guardado.`);
    if (state.ui.view === 'historial') renderHistory();
  } catch (err) {
    console.error(err);
    toast(err.message || 'No se pudo guardar el corte.', 'err');
    if (btn) { btn.disabled = false; btn.textContent = 'Guardar corte'; }
  }
}

async function findCorteByFecha(fecha) {
  if (isSupabaseConfigured()) {
    if (!state.auth.user) return null;
    try {
      const sb = await getSupabase();
      const { data } = await sb
        .from('cortes')
        .select('id,fecha_corte')
        .eq('user_id', state.auth.user.id)
        .eq('fecha_corte', fecha)
        .limit(1);
      return data && data[0] ? data[0] : null;
    } catch (e) {
      return null;
    }
  }
  return readJSON(LS.cortes, []).find(c => c.fecha_corte === fecha) || null;
}

async function deleteCorteById(id, opts) {
  if (isSupabaseConfigured()) {
    const sb = await getSupabase();
    const { error } = await sb.from('cortes').delete().eq('id', id);
    if (error) throw error;
    state.history.cortes = state.history.cortes.filter(c => c.id !== id);
  } else {
    const list = readJSON(LS.cortes, []).filter(c => c.id !== id);
    writeJSON(LS.cortes, list);
    state.history.cortes = sortCortes(list);
  }
  if (state.history.selectedId === id) {
    state.history.selectedId = null;
    state.history.detalles = [];
    state.history.comparison = null;
  }
  if (!(opts && opts.silent)) toast('Corte eliminado.');
}

async function persistCorte(header, detalles) {
  if (isSupabaseConfigured()) {
    const user = state.auth.user;
    if (!user) throw new Error('Inicia sesión para guardar en Supabase.');
    const sb = await getSupabase();
    const { data, error } = await sb
      .from('cortes')
      .insert({ ...header, user_id: user.id })
      .select()
      .single();
    if (error) throw error;

    try {
      const rows = detalles.map(d => ({ ...d, corte_id: data.id, user_id: user.id }));
      for (let i = 0; i < rows.length; i += 400) {
        const chunk = rows.slice(i, i + 400);
        const { error: detailError } = await sb.from('cortes_detalle').insert(chunk);
        if (detailError) throw detailError;
      }
    } catch (detailError) {
      await sb.from('cortes').delete().eq('id', data.id);
      throw detailError;
    }
    return data;
  }

  const corte = {
    ...header,
    id: 'local-' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
    created_at: new Date().toISOString(),
    detalles
  };
  const list = readJSON(LS.cortes, []);
  list.unshift(corte);
  const ok = writeJSON(LS.cortes, list);
  if (!ok) throw new Error('El navegador no tiene espacio. Configura Supabase o borra cortes locales.');
  state.history.cortes = sortCortes(list);
  return corte;
}

async function deleteCorte(corte) {
  const ok = await confirmModal({
    title: 'Borrar corte',
    message: `¿Eliminar el corte del ${formatDate(corte.fecha_corte)}? Esta acción no se puede deshacer.`,
    confirmLabel: 'Borrar corte',
    danger: true
  });
  if (!ok) return;
  try {
    if (!isSupabaseConfigured()) {
      const list = readJSON(LS.cortes, []).filter(c => c.id !== corte.id);
      writeJSON(LS.cortes, list);
      state.history.cortes = sortCortes(list);
    } else {
      const sb = await getSupabase();
      const { error } = await sb.from('cortes').delete().eq('id', corte.id);
      if (error) throw error;
      state.history.cortes = state.history.cortes.filter(c => c.id !== corte.id);
    }
    if (state.history.selectedId === corte.id) {
      state.history.selectedId = null;
      state.history.detalles = [];
      state.history.comparison = null;
    }
    toast('Corte eliminado.');
    renderHistory();
  } catch (err) {
    console.error(err);
    toast(err.message || 'No se pudo borrar el corte.', 'err');
  }
}

async function selectCorte(id) {
  const corte = state.history.cortes.find(c => c.id === id);
  if (!corte) return;
  state.history.selectedId = id;
  state.history.loadingDetail = true;
  state.history.detalles = [];
  state.history.comparison = null;
  state.history.bucket = 'todos';
  renderHistory();
  try {
    const idx = state.history.cortes.findIndex(c => c.id === id);
    const prev = idx >= 0 ? state.history.cortes[idx + 1] : null;
    const [currDet, prevDet] = await Promise.all([
      fetchDetalles(corte),
      prev ? fetchDetalles(prev) : Promise.resolve([])
    ]);
    state.history.detalles = currDet;
    state.history.prevCorte = prev;
    state.history.comparison = prev ? compareCortes(prevDet, currDet) : null;
  } catch (err) {
    console.error(err);
    toast('No se pudo cargar el detalle del corte.', 'err');
    state.history.error = err.message || '';
  } finally {
    state.history.loadingDetail = false;
    renderHistory();
  }
}

/* -------------------------------- VISTA -------------------------------- */

function renderHistory() {
  const authHost = $('#historyAuth');
  const bodyHost = $('#historyBody');
  if (!authHost || !bodyHost) return;
  const configured = isSupabaseConfigured();

  if (configured && !state.auth.user) {
    authHost.innerHTML = authPanelHtml();
    bodyHost.innerHTML = '';
    wireAuthPanel();
    return;
  }
  authHost.innerHTML = '';
  bodyHost.innerHTML = historyBodyHtml(configured);
  wireHistory();

  if (!state.history.loaded && !state.history.loading) {
    loadCortes().then(() => renderHistory());
  }
}

function authPanelHtml() {
  const mode = state.auth.mode || 'login';
  return `
    <div class="card auth-panel">
      <span class="chip chip--primary">Supabase · historial</span>
      <h2>${mode === 'signup' ? 'Crea tu cuenta de trabajo' : 'Inicia sesión para continuar'}</h2>
      <p>Usa el <strong>mismo correo y contraseña de problemacero.com</strong>: es tu misma cuenta (un solo usuario para todo el portal). Si ya iniciaste sesión en el sitio, la app te reconoce automáticamente.</p>
      <p class="muted small" style="margin-top:6px">Los cortes se guardan por usuario y las políticas RLS impiden que otros vean tu historial.</p>
      <div class="auth-tabs">
        <button type="button" data-mode="login" class="${mode === 'login' ? 'is-active' : ''}">Ingresar</button>
        <button type="button" data-mode="signup" class="${mode === 'signup' ? 'is-active' : ''}">Crear cuenta</button>
      </div>
      <form class="form" id="authForm" novalidate>
        <label class="field">
          <span class="field__label">Correo</span>
          <input id="authEmail" type="email" autocomplete="username" placeholder="tu@correo.com" required />
        </label>
        <label class="field">
          <span class="field__label">Contraseña</span>
          <input id="authPass" type="password" autocomplete="${mode === 'signup' ? 'new-password' : 'current-password'}" minlength="6" placeholder="Mínimo 6 caracteres" required />
        </label>
        <button class="btn btn--primary btn--block" type="submit">${mode === 'signup' ? 'Crear cuenta' : 'Ingresar'}</button>
      </form>
      <button class="btn btn--ghost btn--block" type="button" id="btnMagic" style="margin-top:8px">Enviarme un enlace mágico</button>
      <p class="auth-note">El enlace mágico llega al correo y abre esta app con la sesión iniciada.</p>
      <div class="auth-error" id="authError" ${state.auth.error ? '' : 'hidden'}>${esc(state.auth.error || '')}</div>
    </div>`;
}

function wireAuthPanel() {
  const form = $('#authForm');
  const magic = $('#btnMagic');
  $$('#historyAuth [data-mode]').forEach(btn => {
    btn.addEventListener('click', () => {
      state.auth.mode = btn.dataset.mode;
      state.auth.error = '';
      renderHistory();
    });
  });
  if (form) {
    form.addEventListener('submit', async e => {
      e.preventDefault();
      const email = ($('#authEmail') || {}).value || '';
      const pass = ($('#authPass') || {}).value || '';
      if (!email || !pass) return showAuthError('Completa correo y contraseña.');
      await submitAuth(state.auth.mode || 'login', email, pass);
    });
  }
  if (magic) {
    magic.addEventListener('click', async () => {
      const email = ($('#authEmail') || {}).value || '';
      if (!email) return showAuthError('Escribe tu correo para enviarte el enlace.');
      try {
        const sb = await getSupabase();
        const { error } = await sb.auth.signInWithOtp({
          email,
          options: { emailRedirectTo: location.origin + location.pathname }
        });
        if (error) throw error;
        showAuthError('Enlace enviado. Revisa tu correo (y la carpeta de spam).', true);
      } catch (err) {
        showAuthError(traduceAuthError(err.message));
      }
    });
  }
}

function showAuthError(message, ok) {
  const el = $('#authError');
  if (!el) return;
  el.hidden = false;
  el.textContent = message;
  el.style.background = ok ? 'var(--green-soft)' : '';
  el.style.color = ok ? 'var(--green-ink)' : '';
}

async function submitAuth(mode, email, password) {
  try {
    const sb = await getSupabase();
    if (!sb) throw new Error('Supabase no está configurado.');
    if (mode === 'signup') {
      const { data, error } = await sb.auth.signUp({ email, password });
      if (error) throw error;
      if (!data.session) {
        return showAuthError('Cuenta creada. Revisa tu correo para confirmarla y luego ingresa.', true);
      }
    } else {
      const { error } = await sb.auth.signInWithPassword({ email, password });
      if (error) throw error;
    }
    state.auth.error = '';
    state.history.loaded = false;
    renderHistory();
    toast('Sesión iniciada.');
  } catch (err) {
    showAuthError(traduceAuthError(err.message));
  }
}

function traduceAuthError(message) {
  const m = String(message || '');
  if (/Invalid login credentials/i.test(m)) return 'Correo o contraseña incorrectos.';
  if (/Email not confirmed/i.test(m)) return 'Debes confirmar tu correo antes de ingresar.';
  if (/User already registered/i.test(m)) return 'Ese correo ya tiene cuenta. Usa «Ingresar».';
  if (/Password should be at least/i.test(m)) return 'La contraseña debe tener al menos 6 caracteres.';
  if (/Unable to validate email/i.test(m)) return 'El correo no es válido.';
  if (/fetch/i.test(m)) return 'No se pudo conectar con Supabase. Revisa la URL y tu red.';
  return m || 'Ocurrió un error de autenticación.';
}

async function signOutUser() {
  try {
    const sb = await getSupabase();
    if (sb) await sb.auth.signOut();
    state.auth.user = null;
    state.history.loaded = false;
    syncIdentity();
    renderHistory();
    toast('Sesión cerrada.');
  } catch (err) {
    toast('No se pudo cerrar la sesión.', 'err');
  }
}

function historyBodyHtml(configured) {
  if (state.history.loading && !state.history.cortes.length) {
    return `<div class="card"><p class="muted">Cargando cortes…</p></div>`;
  }

  const list = state.history.cortes || [];
  const hasModel = state.model && state.model.rows.length > 0;
  const modeChip = configured
    ? `<span class="chip chip--ok">Supabase</span>`
    : `<span class="chip chip--warn">Modo local</span>`;

  const listHtml = list.length
    ? `<div class="corte-list">${list.map(c => corteItemHtml(c)).join('')}</div>`
    : `<div class="hist-empty"><strong>Sin cortes guardados</strong>Guarda el corte de esta semana para empezar a comparar.</div>`;

  return `
    <div class="card toolbar toolbar--tight">
      <div class="toolbar__filters">
        <div>
          <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap">
            <h2 class="h2">Cortes guardados</h2>
            ${modeChip}
          </div>
          <p class="muted">${configured
            ? 'Historial privado por usuario en Supabase.'
            : 'El historial vive en este navegador. Configura Supabase en Ajustes para compartirlo.'}</p>
        </div>
      </div>
      <div class="toolbar__actions">
        <button class="btn btn--ghost" type="button" data-hist="refresh">Actualizar</button>
        <button class="btn btn--primary" type="button" data-hist="save" ${hasModel ? '' : 'disabled'}>Guardar corte actual</button>
      </div>
    </div>
    ${state.history.error ? `<div class="card" style="margin-bottom:14px"><p class="auth-error" style="margin:0">${esc(state.history.error)}</p></div>` : ''}
    <div class="history-grid">
      <div class="card">
        <div class="card__head"><div><h2>Cortes</h2><p>${list.length} registro${list.length === 1 ? '' : 's'}</p></div></div>
        ${listHtml}
      </div>
      <div class="card" id="historyDetail">${detailHtml()}</div>
    </div>`;
}

function corteItemHtml(corte) {
  const active = state.history.selectedId === corte.id;
  const local = String(corte.id).startsWith('local-');
  return `
    <div class="corte-item ${active ? 'is-active' : ''}" data-corte-id="${esc(corte.id)}" role="button" tabindex="0">
      <div class="corte-item__top">
        <span class="corte-item__date">${esc(formatDate(corte.fecha_corte))}</span>
        ${local ? '<span class="chip chip--warn">local</span>' : ''}
        <button class="corte-item__del" type="button" data-corte-del="${esc(corte.id)}" title="Borrar corte" aria-label="Borrar corte">${ICONS.trash}</button>
      </div>
      <span class="corte-item__meta">
        <span>${fmt(corte.total_sistema)} sistema</span>
        <span>·</span>
        <span>${fmt(corte.total_fisico)} físico</span>
      </span>
      <div class="corte-item__counts">
        <span class="tag tag--eliminar">E ${fmt(corte.total_eliminar)}</span>
        <span class="tag tag--incluir">I ${fmt(corte.total_incluir)}</span>
        <span class="tag tag--ok">OK ${fmt(corte.total_ok)}</span>
        ${corte.total_excluidas ? `<span class="tag tag--excluida">X ${fmt(corte.total_excluidas)}</span>` : ''}
      </div>
    </div>`;
}

function detailHtml() {
  const id = state.history.selectedId;
  if (!id) {
    return `<div class="hist-empty"><strong>Selecciona un corte</strong>Verás el comparativo contra el corte anterior: pendientes que persisten, resueltos y nuevos.</div>`;
  }
  const corte = state.history.cortes.find(c => c.id === id);
  if (!corte) return `<div class="hist-empty">Corte no encontrado.</div>`;
  if (state.history.loadingDetail) {
    return `<div class="detail-head"><div><h2>${esc(formatDate(corte.fecha_corte))}</h2><p>Cargando detalle…</p></div></div><p class="muted">Un momento…</p>`;
  }

  const comp = state.history.comparison;
  const prev = state.history.prevCorte;
  const details = state.history.detalles || [];

  const cards = comp ? `
    <div class="compare-cards">
      <div class="compare-card compare-card--resueltos"><strong>${fmt(comp.resueltos.length)}</strong><span>Resueltos</span></div>
      <div class="compare-card compare-card--persistentes"><strong>${fmt(comp.persistentes.length)}</strong><span>Persisten</span></div>
      <div class="compare-card compare-card--nuevos"><strong>${fmt(comp.nuevos.length)}</strong><span>Nuevos</span></div>
    </div>` : `
    <div class="compare-cards">
      <div class="compare-card"><strong>${fmt(corte.total_eliminar)}</strong><span>Eliminar</span></div>
      <div class="compare-card"><strong>${fmt(corte.total_incluir)}</strong><span>Incluir</span></div>
      <div class="compare-card"><strong>${fmt(corte.total_ok)}</strong><span>OK</span></div>
    </div>`;

  const rows = comparisonRows();
  const buckets = comp
    ? [
        { key: 'todos', label: 'Todos', n: comp.persistentes.length + comp.resueltos.length + comp.nuevos.length },
        { key: 'persistente', label: 'Persisten', n: comp.persistentes.length },
        { key: 'resuelto', label: 'Resueltos', n: comp.resueltos.length },
        { key: 'nuevo', label: 'Nuevos', n: comp.nuevos.length }
      ]
    : [];

  const table = comp ? comparisonTableHtml(rows) : firstCorteTableHtml(details);

  return `
    <div class="detail-head">
      <div>
        <h2>Corte del ${esc(formatDate(corte.fecha_corte))}</h2>
        <p>${esc(corte.archivo_sistema || 'Sin PDF')} · ${esc(corte.archivo_fisico || 'Sin Excel')}${prev ? ' · vs corte del ' + esc(formatDate(prev.fecha_corte)) : ''}</p>
        ${corte.notas ? `<p class="muted small" style="margin-top:5px">${esc(corte.notas)}</p>` : ''}
      </div>
      <button class="btn btn--ghost" type="button" data-hist="export">${ICONS.down}<span>CSV</span></button>
      <button class="btn btn--danger" type="button" data-hist="delete">${ICONS.trash}<span>Borrar</span></button>
    </div>
    ${comp ? '' : '<p class="muted small" style="margin-bottom:12px">Este es el primer corte guardado: aún no hay uno anterior para comparar.</p>'}
    ${cards}
    ${comp ? `<div class="segmented" role="group" aria-label="Filtrar comparativo" style="margin-bottom:12px;flex-wrap:wrap">
      ${buckets.map(b => `<button type="button" data-bucket="${b.key}" class="${state.history.bucket === b.key ? 'is-active' : ''}">${b.label} (${fmt(b.n)})</button>`).join('')}
    </div>` : ''}
    ${table}
  `;
}

function comparisonRows() {
  const comp = state.history.comparison;
  if (!comp) return [];
  const all = [...comp.persistentes, ...comp.nuevos, ...comp.resueltos];
  if (state.history.bucket === 'todos') return all;
  return all.filter(r => r.bucket === state.history.bucket);
}

function estadoTag(estado) {
  if (!estado || estado === '-') return '<span class="muted">-</span>';
  const map = { ELIMINAR: 'eliminar', INCLUIR: 'incluir', REVISAR: 'revisar', OK: 'ok', EXCLUIDA: 'excluida' };
  return `<span class="tag tag--${map[estado] || 'ok'}">${ESTADO_LABEL[estado] || estado}</span>`;
}

function comparisonTableHtml(rows) {
  if (!rows.length) return `<div class="hist-empty">No hay placas en este grupo.</div>`;
  return `
    <p class="table-hint">Desliza la tabla para ver taller, tipo y situación →</p>
    <div class="table-wrap">
      <table class="data">
        <thead><tr><th>Placa</th><th>Anterior</th><th>Actual</th><th>Taller</th><th>Tipo</th><th>Situación</th></tr></thead>
        <tbody>
          ${rows.map(r => `
            <tr>
              <td><span class="plate">${esc(r.placa)}</span></td>
              <td>${estadoTag(r.estadoPrev)}</td>
              <td>${estadoTag(r.estadoCurr)}</td>
              <td class="taller-cell"><span>${esc(r.tallerSistema || r.tallerFisico || '—')}</span></td>
              <td class="muted-cell">${esc(r.tipo || '—')}</td>
              <td><span class="tag ${r.bucket === 'resuelto' ? 'tag--incluir' : r.bucket === 'nuevo' ? 'tag--nuevo' : 'tag--revisar'}">${r.bucket === 'resuelto' ? 'Resuelto' : r.bucket === 'nuevo' ? 'Nuevo' : 'Persiste'}${r.cambio ? ' · cambió' : ''}</span></td>
            </tr>`).join('')}
        </tbody>
      </table>
    </div>`;
}

function firstCorteTableHtml(rows) {
  if (!rows.length) return `<div class="hist-empty">Este corte no tiene pendientes.</div>`;
  return `
    <p class="table-hint">Desliza la tabla para ver taller, tipo y código R →</p>
    <div class="table-wrap">
      <table class="data">
        <thead><tr><th>Placa</th><th>Estado</th><th>Taller</th><th>Tipo</th><th>Código R</th></tr></thead>
        <tbody>
          ${rows.map(r => `
            <tr>
              <td><span class="plate">${esc(r.placa)}</span></td>
              <td>${estadoTag(r.estado)}</td>
              <td class="taller-cell"><span>${esc(r.tallerSistema || r.tallerFisico || '—')}</span></td>
              <td class="muted-cell">${esc(r.tipo || '—')}</td>
              <td class="r-code">${esc(r.r || '—')}</td>
            </tr>`).join('')}
        </tbody>
      </table>
    </div>`;
}

function exportDetailCsv() {
  const corte = state.history.cortes.find(c => c.id === state.history.selectedId);
  if (!corte) return toast('Selecciona un corte.', 'err');
  const comp = state.history.comparison;
  let rows = [];
  if (comp) {
    rows = comparisonRows().map(r => [
      r.placa,
      r.bucket === 'resuelto' ? 'Resuelto' : r.bucket === 'nuevo' ? 'Nuevo' : 'Persiste',
      r.estadoPrev, r.estadoCurr,
      r.tallerSistema || r.tallerFisico, r.tipo, r.cambio ? 'Sí' : ''
    ]);
  } else {
    rows = (state.history.detalles || []).map(r => [r.placa, 'Pendiente', r.estado, r.estado, r.tallerSistema || r.tallerFisico, r.tipo, '']);
  }
  if (!rows.length) return toast('No hay filas para exportar.', 'err');
  downloadFile(
    `BYR_corte_${corte.fecha_corte}.csv`,
    toCsv(['Placa', 'Situación', 'Estado anterior', 'Estado actual', 'Taller', 'Tipo', 'Cambió de estado'], rows),
    'text/csv;charset=utf-8'
  );
  toast(`${fmt(rows.length)} filas exportadas.`);
}

/* ---------------------------- 10. AJUSTES ---------------------------- */

function renderTestPlatesCard() {
  const input = $('#testPlatesInput');
  if (input && document.activeElement !== input) input.value = getTestPlates().join(', ');
}

function renderSettings() {
  const cfg = getConfig();
  const urlInput = $('#supaUrl');
  const keyInput = $('#supaKey');
  if (urlInput && document.activeElement !== urlInput) urlInput.value = cfg.url;
  if (keyInput && document.activeElement !== keyInput) keyInput.value = cfg.key;

  const chip = $('#supaChip');
  if (chip) {
    const ok = isSupabaseConfigured();
    chip.textContent = ok ? (state.auth.user ? 'Conectado' : 'Configurado') : 'Modo local';
    chip.className = 'chip ' + (ok ? 'chip--ok' : 'chip--warn');
  }

  const block = $('#authBlock');
  if (block) block.innerHTML = authBlockHtml();

  renderTestPlatesCard();

  const info = $('#storageInfo');
  if (info) {
    const cortes = readJSON(LS.cortes, []);
    let bytes = 0;
    Object.values(LS).forEach(k => { bytes += (localStorage.getItem(k) || '').length * 2; });
    info.textContent = `${cortes.length} corte${cortes.length === 1 ? '' : 's'} local${cortes.length === 1 ? '' : 'es'} · ~${Math.max(1, Math.round(bytes / 1024))} KB usados en este navegador.`;
  }
}

function applyTestPlatesInput() {
  const input = $('#testPlatesInput');
  if (!input) return;
  const list = saveTestPlates(input.value.split(/[,;\s]+/));
  input.value = list.join(', ');
  rebuild();
  toast(`${fmt(list.length)} placa${list.length === 1 ? '' : 's'} de prueba configurada${list.length === 1 ? '' : 's'}.`);
}

async function clearData() {
  const ok = await confirmModal({
    title: 'Vaciar datos cargados',
    message: 'Se quitarán los archivos y el comparativo actual. Los cortes guardados y las equivalencias se conservan.',
    confirmLabel: 'Vaciar datos',
    danger: true
  });
  if (!ok) return;
  state.files = { sistema: null, fisico: [] };
  state.system = { rows: [], sections: [], warnings: [], meta: {} };
  state.physical = { rows: [], sections: [], warnings: [], meta: {} };
  state.ingresos = [];
  state.corteFecha = '';
  state.importStatus = { sistema: 'empty', fisico: 'empty', message: '' };
  state.model = null;
  state.restaurado = false;
  localStorage.removeItem(LS.ultimo);
  state.ui.open.clear();
  renderAll();
  toast('Datos cargados vaciados.');
}

async function clearLocalHistory() {
  const ok = await confirmModal({
    title: 'Borrar historial local',
    message: 'Se eliminarán todos los cortes guardados en este navegador. Los de Supabase no se tocan.',
    confirmLabel: 'Borrar historial',
    danger: true
  });
  if (!ok) return;
  localStorage.removeItem(LS.cortes);
  state.history.cortes = [];
  state.history.selectedId = null;
  state.history.detalles = [];
  state.history.comparison = null;
  state.history.loaded = true;
  if (state.ui.view === 'historial') renderHistory();
  renderSettings();
  toast('Historial local borrado.');
}

/* ---------------------------- 11. EVENTOS ---------------------------- */

function wireHistory() { /* los eventos del historial viven en un único delegado */ }

function openImportModal() {
  openModal({
    title: 'Cargar archivos del corte',
    body: `
      <p class="muted small">Sube el PDF del sistema y el Excel del conteo físico. Puedes reemplazar uno sin perder el otro.</p>
      <div class="drops" style="margin-top:12px">
        <button class="drop" type="button" data-import="sistema">
          <span class="drop__icon">${ICONS.pdf}</span>
          <span class="drop__text"><strong>Inventario del sistema</strong><span>PDF agrupado por taller/parqueadero</span><em>Elegir PDF</em></span>
        </button>
        <button class="drop" type="button" data-import="fisico">
          <span class="drop__icon">${ICONS.sheet}</span>
          <span class="drop__text"><strong>Conteo físico</strong><span>Excel .xlsx / .xls / .csv</span><em>Elegir Excel</em></span>
        </button>
      </div>
      <p class="muted small" style="margin-top:12px">Los archivos se procesan localmente en el navegador; nada se sube hasta que guardas el corte.</p>
    `,
    foot: '<button class="btn btn--ghost" type="button" data-close-modal>Cancelar</button>',
    onMount: (body, foot) => {
      body.querySelectorAll('[data-import]').forEach(btn => {
        btn.addEventListener('click', () => {
          const kind = btn.dataset.import;
          closeModal();
          pickFile(kind);
        });
      });
      foot.querySelector('[data-close-modal]').addEventListener('click', closeModal);
    }
  });
}

function bindEvents() {
  $$('.nav__item[data-view]').forEach(btn => {
    btn.addEventListener('click', () => setView(btn.dataset.view));
  });

  const menu = $('#menuBtn');
  if (menu) menu.addEventListener('click', openSidebar);
  const scrim = $('#scrim');
  if (scrim) scrim.addEventListener('click', closeSidebar);

  const importBtns = ['#btnImport', '#btnImportSide'];
  importBtns.forEach(sel => {
    const el = $(sel);
    if (el) el.addEventListener('click', openImportModal);
  });

  const help = $('#btnHelp');
  if (help) help.addEventListener('click', showHelp);

  // Modo claro / oscuro: alterna data-theme en <html> y recuerda la elección.
  const themeBtn = $('#themeToggle');
  if (themeBtn) {
    const paint = () => {
      const dark = document.documentElement.getAttribute('data-theme') === 'dark';
      themeBtn.setAttribute('aria-pressed', dark ? 'true' : 'false');
      themeBtn.title = dark ? 'Cambiar a modo claro' : 'Cambiar a modo oscuro';
    };
    paint();
    themeBtn.addEventListener('click', () => {
      const dark = document.documentElement.getAttribute('data-theme') === 'dark';
      if (dark) {
        document.documentElement.removeAttribute('data-theme');
        try { localStorage.setItem('byr.theme', 'light'); } catch (e) {}
      } else {
        document.documentElement.setAttribute('data-theme', 'dark');
        try { localStorage.setItem('byr.theme', 'dark'); } catch (e) {}
      }
      paint();
    });
  }

  const save = $('#btnSaveCorte');
  if (save) save.addEventListener('click', openSaveModal);

  const logout = $('#btnLogout');
  if (logout) logout.addEventListener('click', signOutUser);

  const fileSys = $('#fileSistema');
  if (fileSys) fileSys.addEventListener('change', e => handleFileList('sistema', e.target.files));
  const fileFis = $('#fileFisico');
  if (fileFis) fileFis.addEventListener('change', e => handleFileList('fisico', e.target.files));

  // Import card (delegado: el contenido se re-renderiza)
  const importCard = $('#importCard');
  if (importCard) {
    importCard.addEventListener('click', e => {
      const reset = e.target.closest('[data-reset]');
      if (reset) {
        resetImport(reset.dataset.reset);
        return;
      }
      if (e.target.closest('[data-avisos]')) {
        showWarnings();
        return;
      }
      if (e.target.closest('[data-descartar-restaurado]')) {
        discardUltimo();
        return;
      }
      const historialLink = e.target.closest('[data-view="historial"]');
      if (historialLink) {
        setView('historial');
        return;
      }
      const drop = e.target.closest('[data-drop]');
      if (drop) pickFile(drop.dataset.drop);
    });
    importCard.addEventListener('keydown', e => {
      if (e.key !== 'Enter' && e.key !== ' ') return;
      const drop = e.target.closest('[data-drop]');
      if (!drop) return;
      e.preventDefault();
      pickFile(drop.dataset.drop);
    });
    // Fecha de corte editable en «Archivos del corte»
    importCard.addEventListener('change', e => {
      if (e.target.id !== 'corteFechaInput') return;
      applyCorteFecha(e.target.value);
    });
  }

  // Filtros
  const ft = $('#filterTaller');
  if (ft) ft.addEventListener('change', () => { state.filters.taller = ft.value; renderFilters(); renderSections(); });
  const ftipo = $('#filterTipo');
  if (ftipo) ftipo.addEventListener('change', () => { state.filters.tipo = ftipo.value; renderFilters(); renderSections(); });
  const festado = $('#filterEstado');
  if (festado) festado.addEventListener('change', () => { state.filters.estado = festado.value; renderFilters(); renderSections(); });

  const search = $('#searchPlaca');
  if (search) {
    search.addEventListener('input', debounce(() => {
      state.filters.q = search.value.trim();
      renderFilters();
      renderSections();
    }, 150));
  }

  const clear = $('#btnClearFilters');
  if (clear) clear.addEventListener('click', () => {
    state.filters = { taller: '', tipo: '', estado: '', q: '' };
    if (search) search.value = '';
    renderFilters();
    renderSections();
  });

  const copyInc = $('#btnCopyIncluir');
  if (copyInc) copyInc.addEventListener('click', () => {
    copyPlates(filteredRows().filter(r => r.estado === 'INCLUIR').map(r => r.placaNorm), 'incluir');
  });
  const copyEli = $('#btnCopyEliminar');
  if (copyEli) copyEli.addEventListener('click', async () => {
    const rows = filteredRows().filter(r => r.estado === 'ELIMINAR');
    if (!rows.length) return toast('No hay placas para copiar con los filtros actuales.', 'err');
    if (!(await confirmEliminar(rows.length, 'Copiar'))) return;
    copyPlates(rows.map(r => r.placaNorm), 'eliminar');
  });
  const exportBtn = $('#btnExport');
  if (exportBtn) exportBtn.addEventListener('click', exportFilteredCsv);
  const exportXls = $('#btnExportExcel');
  if (exportXls) exportXls.addEventListener('click', exportFilteredExcel);

  const groupBy = $('#groupBy');
  if (groupBy) {
    groupBy.addEventListener('click', e => {
      const btn = e.target.closest('[data-group]');
      if (!btn) return;
      state.ui.groupBy = btn.dataset.group;
      state.filters.taller = '';
      $$('#groupBy button').forEach(b => b.classList.toggle('is-active', b === btn));
      renderFilters();
      renderSections();
      renderBars();
    });
  }

  // Secciones
  const sections = $('#sections');
  if (sections) {
    sections.addEventListener('click', e => {
      const toggle = e.target.closest('[data-toggle]');
      if (toggle) {
        const key = toggle.dataset.toggle;
        if (state.ui.open.has(key)) state.ui.open.delete(key);
        else state.ui.open.add(key);
        renderSections();
        return;
      }
      const sectionCopy = e.target.closest('[data-copy-section]');
      if (sectionCopy) {
        const rows = rowsForSection(sectionCopy.dataset.copySection, sectionCopy.dataset.estado);
        const group = groupedSections(filteredRows(), state.ui.groupBy).find(g => g.key === sectionCopy.dataset.copySection);
        const esEliminar = sectionCopy.dataset.estado === 'ELIMINAR';
        const label = `${esEliminar ? 'eliminar' : 'incluir'}${group ? ' · ' + group.label : ''}`;
        if (esEliminar) {
          if (!rows.length) return toast('No hay placas a eliminar en esta sección.', 'err');
          confirmEliminar(rows.length, 'Copiar').then(ok => { if (ok) copyPlates(rows.map(r => r.placaNorm), label); });
        } else {
          copyPlates(rows.map(r => r.placaNorm), label);
        }
        return;
      }
      const one = e.target.closest('[data-copy-placa]');
      if (one) copyPlates([normPlaca(one.dataset.copyPlaca)], one.dataset.copyPlaca);
    });
  }

  // Barras → filtro por taller
  const bars = $('#bars');
  if (bars) {
    bars.addEventListener('click', e => {
      const btn = e.target.closest('[data-bar]');
      if (!btn) return;
      state.filters.taller = btn.dataset.bar;
      state.filters.estado = '';
      const sel = $('#filterEstado');
      if (sel) sel.value = '';
      if (state.ui.view !== 'comparacion') setView('comparacion');
      renderFilters();
      renderSections();
      const sec = $('#sections');
      if (sec) sec.scrollIntoView({ behavior: 'smooth', block: 'start' });
    });
  }

  // Anomalías
  const copyAnom = $('#btnCopyAnomalias');
  if (copyAnom) copyAnom.addEventListener('click', copyAllAnomalias);
  const expAnom = $('#btnExportAnomalias');
  if (expAnom) expAnom.addEventListener('click', exportAnomaliasCsv);

  // Historial
  const historyBody = $('#historyBody');
  if (historyBody) {
    historyBody.addEventListener('click', e => {
      const refresh = e.target.closest('[data-hist="refresh"]');
      if (refresh) {
        state.history.loaded = false;
        loadCortes().then(() => renderHistory());
        return;
      }
      if (e.target.closest('[data-hist="save"]')) return openSaveModal();
      if (e.target.closest('[data-hist="export"]')) return exportDetailCsv();
      if (e.target.closest('[data-hist="delete"]')) {
        const corte = state.history.cortes.find(c => c.id === state.history.selectedId);
        if (corte) deleteCorte(corte);
        return;
      }
      const bucket = e.target.closest('[data-bucket]');
      if (bucket) {
        state.history.bucket = bucket.dataset.bucket;
        renderHistory();
        return;
      }
      const del = e.target.closest('[data-corte-del]');
      if (del) {
        const corte = state.history.cortes.find(c => c.id === del.dataset.corteDel);
        if (corte) deleteCorte(corte);
        return;
      }
      const item = e.target.closest('[data-corte-id]');
      if (item) selectCorte(item.dataset.corteId);
    });
    historyBody.addEventListener('keydown', e => {
      if (e.key !== 'Enter' && e.key !== ' ') return;
      const item = e.target.closest('[data-corte-id]');
      if (!item) return;
      e.preventDefault();
      selectCorte(item.dataset.corteId);
    });
  }

  // Ajustes
  const authBlock = $('#authBlock');
  if (authBlock) {
    authBlock.addEventListener('click', e => {
      const btn = e.target.closest('[data-auth]');
      if (!btn) return;
      if (btn.dataset.auth === 'logout') signOutUser();
      if (btn.dataset.auth === 'go') setView('historial');
    });
  }

  // Placas de prueba editables (Ajustes)
  const testInput = $('#testPlatesInput');
  if (testInput) {
    testInput.addEventListener('change', applyTestPlatesInput);
    testInput.addEventListener('blur', applyTestPlatesInput);
  }

  const supaForm = $('#supaForm');
  if (supaForm) {
    supaForm.addEventListener('submit', async e => {
      e.preventDefault();
      const url = ($('#supaUrl') || {}).value.trim().replace(/\/+$/, '');
      const key = ($('#supaKey') || {}).value.trim();
      if (!/^https?:\/\/.+\..+/.test(url)) return toast('La URL del proyecto no es válida.', 'err');
      if (key.length < 20) return toast('La llave anon parece incompleta.', 'err');
      writeJSON(LS.supabase, { url, key });
      state.auth.client = null;
      state.auth.user = null;
      state.history.loaded = false;
      state.history.cortes = [];
      state.history.selectedId = null;
      await initAuth();
      renderSettings();
      renderHistory();
      toast('Credenciales guardadas. Inicia sesión en Historial.');
    });
  }

  const supaClear = $('#btnSupaClear');
  if (supaClear) supaClear.addEventListener('click', () => {
    localStorage.removeItem(LS.supabase);
    state.auth.client = null;
    state.auth.user = null;
    state.history.loaded = false;
    state.history.cortes = [];
    state.history.selectedId = null;
    syncIdentity();
    renderSettings();
    renderHistory();
    toast('Credenciales eliminadas. Modo local activo.');
  });

  const clearDataBtn = $('#btnClearData');
  if (clearDataBtn) clearDataBtn.addEventListener('click', clearData);
  const clearHist = $('#btnClearLocalHistory');
  if (clearHist) clearHist.addEventListener('click', clearLocalHistory);

  // Modal
  $$('#modal [data-close]').forEach(el => el.addEventListener('click', closeModal));
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape' && !$('#modal').hidden) closeModal();
  });

  // Listas internas: actualiza el degradado inferior al desplazarse
  document.addEventListener('scroll', e => {
    const el = e.target;
    if (!el || !el.classList || !el.classList.contains('is-scrollable')) return;
    if (!el.classList.contains('anomalia-list') && !el.classList.contains('mapping')) return;
    el.classList.toggle('at-end', el.scrollTop + el.clientHeight >= el.scrollHeight - 6);
  }, true);
  let scrollHintFrame = 0;
  window.addEventListener('resize', () => {
    if (scrollHintFrame) cancelAnimationFrame(scrollHintFrame);
    scrollHintFrame = requestAnimationFrame(() => { scrollHintFrame = 0; markScrollables(); });
  });

  // Drag & drop global (evita que el navegador abra el archivo)
  const hasFiles = e => e.dataTransfer && Array.from(e.dataTransfer.types || []).indexOf('Files') >= 0;
  document.addEventListener('dragenter', e => {
    if (!hasFiles(e)) return;
    e.preventDefault();
    const drop = e.target.closest && e.target.closest('[data-drop]');
    if (drop) drop.classList.add('is-over');
  });
  document.addEventListener('dragover', e => {
    if (!hasFiles(e)) return;
    e.preventDefault();
    const drop = e.target.closest && e.target.closest('[data-drop]');
    $$('[data-drop]').forEach(d => { if (d !== drop) d.classList.remove('is-over'); });
    if (drop) drop.classList.add('is-over');
  });
  document.addEventListener('dragleave', e => {
    const drop = e.target.closest && e.target.closest('[data-drop]');
    if (drop) drop.classList.remove('is-over');
  });
  document.addEventListener('drop', e => {
    if (!hasFiles(e)) return;
    e.preventDefault();
    $$('[data-drop]').forEach(d => d.classList.remove('is-over'));
    const drop = e.target.closest && e.target.closest('[data-drop]');
    if (!drop) return toast('Suelta los archivos sobre una zona de carga.', 'err');
    handleFileList(drop.dataset.drop, e.dataTransfer.files);
  });
}

/* ---------------------------- ARRANQUE ---------------------------- */

function boot() {
  restoreUltimo();
  bindEvents();
  renderAll();
  syncIdentity();
  initAuth();
  const hash = String(location.hash || '').replace('#', '');
  if (['resumen', 'comparacion', 'anomalias', 'excluidas', 'historial', 'ajustes'].includes(hash)) setView(hash);
}

if (typeof document !== 'undefined') {
  boot();
}

if (typeof window !== 'undefined') {
  window.__BYR = {
    state,
    parseSystemLines,
    parsePhysicalSheets,
    buildModel,
    normPlaca,
    normTaller,
    tipoFromName,
    plateFormat,
    isGarbageR,
    rebuild,
    renderAll,
    setView,
    readPdfFile,
    readPhysicalFiles,
    loadSystemFile,
    loadPhysicalFile,
    formatDate
  };
}

export {
  state,
  normPlaca,
  normR,
  normTaller,
  sameTaller,
  tipoFromName,
  plateFormat,
  isGarbageR,
  isTestPlate,
  parseSystemLines,
  parsePhysicalSheets,
  buildModel,
  detectAnomalies,
  pdfItemsToLines
};




