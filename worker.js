// Motor de análisis de Nube BI.
// Corre en un hilo aparte: lee el Excel, detecta tipos de columna,
// limpia y genera las recomendaciones de gráficos ya agregadas.
importScripts("https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js");

self.onmessage = (event) => {
  const { id, op, buffer, name, parts } = event.data;
  try {
    const result = op === "combine" ? combineWorkbooks(parts, name) : analyzeWorkbook(buffer, name);
    self.postMessage({ id, ok: true, result });
  } catch (err) {
    self.postMessage({ id, ok: false, error: String((err && err.message) || err) });
  }
};

const MAX_SHEETS = 6;
const TOP_N = 12;

const SEMANTIC = {
  measure: /(venta|monto|total|importe|ingreso|precio|valor|costo|coste|gasto|pago|saldo|utilidad|ganancia|cantidad|unidades|amount|sales|revenue|price|cost|qty|quantity|profit)/i,
  dimension: /(categor|producto|cliente|regi[oó]n|zona|ciudad|pa[ií]s|vendedor|tipo|estado|sucursal|tienda|departamento|area|área|canal|marca|proveedor|segmento|category|product|customer|region|city|country|type|status|store|channel|brand)/i,
  identifier: /(^id$|^id[_ ]|[_ ]id$|c[oó]digo|^cod|n[uú]mero|^nro|^num|dni|rut|cedula|c[eé]dula|tel[eé]fono|celular|email|correo|sku|folio|^#)/i,
  year: /(^a[nñ]o$|^year$)/i,
};

const FILE_CATEGORIES = [
  ["Ventas", /(venta|cliente|factura|pedido|producto|sales|revenue|ingreso|orden)/i],
  ["Finanzas", /(gasto|presupuesto|costo|pago|banco|saldo|cuenta|impuesto|contab|flujo|caja|budget|expense)/i],
  ["Inventario", /(stock|inventario|almac[eé]n|existencia|bodega|sku|lote|inventory)/i],
  ["Personal", /(empleado|sueldo|salario|n[oó]mina|cargo|rrhh|asistencia|vacacion|employee|payroll)/i],
  ["Marketing", /(campa[nñ]a|clic|impresion|lead|ctr|conversi[oó]n|seguidor|alcance|campaign)/i],
];

// ---------- Lectura ----------

// Los CSV se leen como texto (UTF-8 o Windows-1252) y sin interpretar,
// para que las fechas dd/mm/aaaa no se confundan con el formato de EE. UU.
function readBook(buffer, fileName) {
  if (/\.(csv|txt)$/i.test(fileName)) {
    let text;
    try { text = new TextDecoder("utf-8", { fatal: true }).decode(buffer); }
    catch { text = new TextDecoder("windows-1252").decode(buffer); }
    const firstLine = text.slice(0, text.indexOf("\n") >>> 0);
    const FS = (firstLine.match(/;/g) || []).length > (firstLine.match(/,/g) || []).length ? ";" : ",";
    return XLSX.read(text.replace(/^﻿/, ""), { type: "string", raw: true, dense: true, FS });
  }
  return XLSX.read(new Uint8Array(buffer), { type: "array", cellDates: true, dense: true });
}

function analyzeWorkbook(buffer, fileName) {
  const wb = readBook(buffer, fileName);
  const sheets = [];
  for (const sheetName of wb.SheetNames.slice(0, MAX_SHEETS)) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, raw: true, defval: null, blankrows: false });
    const table = extractTable(rows);
    if (!table || table.rows.length === 0) continue;
    sheets.push(analyzeTable(sheetName, table));
  }
  if (sheets.length === 0) throw new Error("No encontré una tabla con datos en este archivo.");
  sheets.sort((a, b) => b.score - a.score);

  const allHeaders = sheets.flatMap((s) => s.columns.map((c) => c.name)).join(" ");
  return {
    fileName,
    category: detectCategory(fileName + " " + allHeaders),
    year: detectYear(sheets, fileName),
    signature: sheets[0].columns.map((c) => normalize(c.name)).sort(),
    sheets,
    analyzedAt: Date.now(),
  };
}

// Une varios archivos con columnas parecidas en una sola tabla.
function combineWorkbooks(parts, name) {
  const tables = parts.map(({ buffer, name: partName }) => {
    const wb = readBook(buffer, partName);
    let best = null;
    for (const sn of wb.SheetNames.slice(0, MAX_SHEETS)) {
      const t = extractTable(XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: true, defval: null, blankrows: false }));
      if (t && (!best || t.rows.length > best.rows.length)) best = t;
    }
    return best && { ...best, partName };
  }).filter(Boolean);
  if (!tables.length) throw new Error("No hay datos para combinar.");

  const headers = [...tables[0].headers];
  const keyIdx = new Map(headers.map((h, i) => [normalize(h), i]));
  for (const t of tables.slice(1)) for (const h of t.headers) {
    if (!keyIdx.has(normalize(h))) { keyIdx.set(normalize(h), headers.length); headers.push(h); }
  }
  headers.push("Archivo");
  const rows = [];
  const cleaning = { emptyRemoved: 0, repeatedHeaders: 0, emptyColumns: 0 };
  for (const t of tables) {
    const map = t.headers.map((h) => keyIdx.get(normalize(h)));
    for (const r of t.rows) {
      const out = new Array(headers.length).fill(null);
      map.forEach((dst, src) => (out[dst] = r[src]));
      out[headers.length - 1] = t.partName;
      rows.push(out);
    }
    cleaning.emptyRemoved += t.cleaning.emptyRemoved;
    cleaning.repeatedHeaders += t.cleaning.repeatedHeaders;
  }
  const sheet = analyzeTable("Combinado", { headers, rows, cleaning });
  return {
    fileName: name,
    category: detectCategory(name + " " + headers.join(" ")),
    year: detectYear([sheet], name),
    signature: sheet.columns.filter((c) => c.name !== "Archivo").map((c) => normalize(c.name)).sort(),
    sheets: [sheet],
    analyzedAt: Date.now(),
  };
}

function extractTable(rawRows) {
  // La fila de encabezados es la primera con al menos 2 celdas y mayoría de texto.
  let headerIdx = -1;
  for (let i = 0; i < Math.min(rawRows.length, 30); i++) {
    const r = rawRows[i] || [];
    const filled = r.filter((v) => v !== null && v !== "");
    const texts = filled.filter((v) => typeof v === "string");
    if (filled.length >= 2 && texts.length / filled.length >= 0.6) { headerIdx = i; break; }
  }
  if (headerIdx === -1) return null;

  const rawHeader = rawRows[headerIdx];
  const width = Math.max(...rawRows.slice(headerIdx, headerIdx + 200).map((r) => (r ? r.length : 0)));
  const seen = new Map();
  const headers = [];
  for (let c = 0; c < width; c++) {
    let h = rawHeader[c] == null || rawHeader[c] === "" ? `Columna ${c + 1}` : String(rawHeader[c]).trim();
    const k = h.toLowerCase();
    if (seen.has(k)) { seen.set(k, seen.get(k) + 1); h = `${h} (${seen.get(k)})`; } else seen.set(k, 1);
    headers.push(h);
  }

  const headerKey = headers.join("|").toLowerCase();
  let emptyRemoved = 0, repeatedHeaders = 0;
  const rows = [];
  for (let i = headerIdx + 1; i < rawRows.length; i++) {
    const r = rawRows[i];
    if (!r || r.every((v) => v === null || v === "")) { emptyRemoved++; continue; }
    if (r.map((v) => (v == null ? "" : String(v).trim())).join("|").toLowerCase() === headerKey) { repeatedHeaders++; continue; }
    rows.push(r);
  }

  // Quita columnas totalmente vacías.
  const keep = headers.map((_, c) => rows.some((r) => r[c] !== null && r[c] !== undefined && r[c] !== ""));
  return {
    headers: headers.filter((_, c) => keep[c]),
    rows: rows.map((r) => r.filter((_, c) => keep[c])),
    cleaning: { emptyRemoved, repeatedHeaders, emptyColumns: keep.filter((k) => !k).length },
  };
}

// ---------- Conversión de valores ----------

function toNum(v) {
  if (typeof v === "number") return Number.isFinite(v) ? v : null;
  if (typeof v !== "string") return null;
  let s = v.trim().replace(/^(USD|EUR|S\/\.?|COP|MXN|ARS|CLP|PEN|Bs\.?)/i, "").replace(/[\s$€£%]/g, "");
  if (!/^[-+]?[\d.,]+$/.test(s) || !/\d/.test(s)) return null;
  const lc = s.lastIndexOf(","), ld = s.lastIndexOf(".");
  if (lc > -1 && ld > -1) {
    s = lc > ld ? s.replace(/\./g, "").replace(",", ".") : s.replace(/,/g, "");
  } else if (lc > -1) {
    const parts = s.split(",");
    s = parts.length === 2 && parts[1].length !== 3 ? s.replace(",", ".") : s.replace(/,/g, "");
  } else if (ld > -1 && s.split(".").length > 2) {
    s = s.replace(/\./g, "");
  }
  const n = parseFloat(s);
  return Number.isFinite(n) ? n : null;
}

function toDate(v) {
  if (v instanceof Date) return Number.isNaN(v.getTime()) ? null : v;
  if (typeof v !== "string") return null;
  const s = v.trim();
  let m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (m) return mkDate(+m[1], +m[2], +m[3]);
  m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})/);
  if (m) {
    let y = +m[3]; if (y < 100) y += 2000;
    let d = +m[1], mo = +m[2];
    if (mo > 12 && d <= 12) [d, mo] = [mo, d]; // formato mes/día
    return mkDate(y, mo, d);
  }
  m = s.match(/^(\d{4})[-/](\d{1,2})$/);
  if (m) return mkDate(+m[1], +m[2], 1);
  return null;
}

function mkDate(y, mo, d) {
  if (mo < 1 || mo > 12 || d < 1 || d > 31 || y < 1900 || y > 2200) return null;
  return new Date(Date.UTC(y, mo - 1, d));
}

function normalize(s) {
  return String(s).toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/[^a-z0-9]+/g, " ").trim();
}

// ---------- Perfil de columnas ----------

function profileColumn(name, values) {
  let nonNull = 0, nums = 0, dates = 0, ints = 0;
  const distinct = new Set();
  for (const v of values) {
    if (v === null || v === undefined || v === "") continue;
    nonNull++;
    if (distinct.size < 5000) distinct.add(v instanceof Date ? v.getTime() : String(v).trim().toLowerCase());
    if (toDate(v)) { dates++; continue; }
    const n = toNum(v);
    if (n !== null) { nums++; if (Number.isInteger(n)) ints++; }
  }
  const ratio = (x) => (nonNull ? x / nonNull : 0);
  const n = nonNull;
  let type = "text";
  if (n === 0) type = "empty";
  else if (ratio(dates) >= 0.8) type = "date";
  else if (ratio(nums) >= 0.8) {
    if (SEMANTIC.year.test(name) && ints === nums) type = "category";
    else if (SEMANTIC.identifier.test(name) && !SEMANTIC.measure.test(name)) type = "id";
    else type = "number";
  } else if (SEMANTIC.identifier.test(name)) type = "id";
  else if (distinct.size <= 60 || distinct.size / n <= 0.3) type = "category";

  return { name, type, distinct: distinct.size, nulls: values.length - nonNull, filled: n };
}

function analyzeTable(sheetName, { headers, rows, cleaning }) {
  const cols = headers.map((h, c) => profileColumn(h, rows.map((r) => r[c])));
  cols.forEach((c, i) => (c.index = i));

  const measures = cols.filter((c) => c.type === "number")
    .sort((a, b) => measureScore(b) - measureScore(a));
  const summable = measures.filter((m) => !isAverageLike(m.name));
  const dimensions = cols.filter((c) => c.type === "category" && c.distinct >= 2)
    .sort((a, b) => rank(b, SEMANTIC.dimension) - rank(a, SEMANTIC.dimension) || a.distinct - b.distinct);
  const dateCol = cols.find((c) => c.type === "date");

  const numCache = new Map();
  const numsOf = (col) => {
    if (!numCache.has(col.index)) numCache.set(col.index, rows.map((r) => toNum(r[col.index])));
    return numCache.get(col.index);
  };

  // Duplicados exactos
  const seenRows = new Set(); let duplicates = 0;
  for (const r of rows) { const k = JSON.stringify(r); if (seenRows.has(k)) duplicates++; else seenRows.add(k); }

  // KPIs
  const kpis = [{ label: "Registros", value: rows.length, format: "int" }];
  for (const m of measures.slice(0, 3)) {
    const vals = numsOf(m).filter((v) => v !== null);
    const sum = vals.reduce((a, b) => a + b, 0);
    kpis.push(isAverageLike(m.name)
      ? { label: `Promedio ${m.name}`, value: vals.length ? sum / vals.length : 0, format: "num" }
      : { label: /^total/i.test(m.name) ? m.name : `Total ${m.name}`, value: sum, format: "num" });
  }

  const charts = [];
  const primary = summable[0];
  const valueLabel = primary ? primary.name : "Registros";
  const valueOf = (i) => (primary ? numsOf(primary)[i] || 0 : 1);

  // 1) Tendencia en el tiempo
  if (dateCol) {
    const dates = rows.map((r) => toDate(r[dateCol.index]));
    const range = dateRangeOf(rows, dateCol.index);
    const span = range ? (range.to - range.from) / 864e5 : 0;
    const byDay = span <= 62;
    const series = primary ? [primary] : [];
    const buckets = new Map();
    dates.forEach((d, i) => {
      if (!d) return;
      const key = byDay ? d.toISOString().slice(0, 10) : d.toISOString().slice(0, 7);
      const b = buckets.get(key) || new Array(Math.max(series.length, 1)).fill(0);
      if (series.length) series.forEach((s, j) => (b[j] += numsOf(s)[i] || 0)); else b[0] += 1;
      buckets.set(key, b);
    });
    const keys = [...buckets.keys()].sort();
    if (keys.length >= 2) {
      charts.push({
        type: "line",
        title: `${series.length ? series.map((s) => s.name).join(" y ") : "Registros"} por ${byDay ? "día" : "mes"}`,
        why: `Detecté la columna de fecha "${dateCol.name}", así que muestro cómo cambia en el tiempo.`,
        labels: keys.map((k) => formatPeriod(k, byDay)),
        datasets: (series.length ? series : [{ name: "Registros" }]).map((s, j) => ({ label: s.name, data: keys.map((k) => round(buckets.get(k)[j])) })),
      });
    }
  }

  // 2) Comparación por categoría
  const groupBy = (dim) => {
    const m = new Map();
    rows.forEach((r, i) => {
      const raw = r[dim.index];
      const k = raw === null || raw === "" ? "(vacío)" : raw instanceof Date ? raw.toISOString().slice(0, 10) : String(raw).trim();
      m.set(k, (m.get(k) || 0) + valueOf(i));
    });
    return [...m.entries()].sort((a, b) => b[1] - a[1]);
  };
  const topWithOthers = (entries, n) => {
    if (entries.length <= n) return entries;
    const rest = entries.slice(n - 1).reduce((a, e) => a + e[1], 0);
    return [...entries.slice(0, n - 1), ["Otros", rest]];
  };

  const barDims = dimensions.filter((d) => d.distinct <= 2000).slice(0, 2);
  barDims.forEach((dim, idx) => {
    const entries = topWithOthers(groupBy(dim), TOP_N);
    charts.push({
      type: "bar",
      horizontal: entries.some(([k]) => k.length > 12),
      title: `${valueLabel} por ${dim.name}`,
      why: idx === 0
        ? `"${dim.name}" agrupa bien los datos (${dim.distinct} valores distintos); las barras comparan cuál aporta más.`
        : `Otra forma de cortar los datos: por "${dim.name}".`,
      labels: entries.map((e) => e[0]),
      datasets: [{ label: valueLabel, data: entries.map((e) => round(e[1])) }],
    });
  });

  // 3) Participación (pocas categorías)
  const pieDim = dimensions.find((d) => d.distinct >= 2 && d.distinct <= 6);
  if (pieDim) {
    const entries = groupBy(pieDim);
    charts.push({
      type: "doughnut",
      title: `Participación de ${pieDim.name}`,
      why: `"${pieDim.name}" tiene solo ${pieDim.distinct} opciones: ideal para ver qué porcentaje es cada una.`,
      labels: entries.map((e) => e[0]),
      datasets: [{ label: valueLabel, data: entries.map((e) => round(e[1])) }],
    });
  }

  // 4) Distribución si no hay con qué agrupar
  if (charts.length === 0 && primary) {
    const vals = numsOf(primary).filter((v) => v !== null);
    if (vals.length) {
      let min = Infinity, max = -Infinity;
      for (const v of vals) { if (v < min) min = v; if (v > max) max = v; }
      const bins = 10, step = (max - min) / bins || 1;
      const counts = new Array(bins).fill(0);
      vals.forEach((v) => counts[Math.min(bins - 1, Math.floor((v - min) / step))]++);
      charts.push({
        type: "bar",
        title: `Distribución de ${primary.name}`,
        why: "No encontré fechas ni categorías, así que muestro cómo se reparten los valores.",
        labels: counts.map((_, i) => `${shortNum(min + i * step)}–${shortNum(min + (i + 1) * step)}`),
        datasets: [{ label: "Cantidad", data: counts }],
      });
    }
  }

  // Tabla de muestra
  const preview = rows.slice(0, 50).map((r) => r.map((v) => (v instanceof Date ? v.toISOString().slice(0, 10) : v)));

  const notes = [];
  if (cleaning.emptyRemoved) notes.push(`Quité ${cleaning.emptyRemoved} filas vacías.`);
  if (cleaning.repeatedHeaders) notes.push(`Quité ${cleaning.repeatedHeaders} encabezados repetidos dentro de los datos.`);
  if (cleaning.emptyColumns) notes.push(`Ignoré ${cleaning.emptyColumns} columnas vacías.`);
  if (duplicates) notes.push(`Hay ${duplicates} filas duplicadas exactas (no las borré; revísalas).`);
  const sparse = cols.filter((c) => c.nulls / rows.length > 0.3);
  if (sparse.length) notes.push(`Columnas con muchos vacíos: ${sparse.map((c) => c.name).join(", ")}.`);
  if (!dateCol) notes.push("Agrega una columna de fecha para ver tendencias en el tiempo.");

  return {
    sheet: sheetName,
    rowCount: rows.length,
    columns: cols.map(({ name, type, distinct, nulls }) => ({ name, type, distinct, nulls })),
    kpis,
    charts,
    preview: { headers, rows: preview },
    notes,
    dateRange: dateCol ? dateRangeOf(rows, dateCol.index) : null,
    score: rows.length * (1 + charts.length),
  };
}

function rank(col, re) { return re.test(col.name) ? 1 : 0; }

// Qué columna numérica es la "principal": montos totales primero, luego cantidades;
// precios y porcentajes van al final porque sumarlos no tiene sentido.
function measureScore(col) {
  const n = col.name;
  if (isAverageLike(n)) return 0;
  if (/(total|monto|importe|venta|ingreso|gasto|revenue|sales|amount|utilidad|ganancia|profit|saldo|presupuesto)/i.test(n)) return 3;
  if (/(cantidad|unidades|qty|quantity|stock|existencia)/i.test(n)) return 2;
  return 1;
}
function isAverageLike(name) {
  return /(precio|price|unitari|tasa|rate|promedio|average|%|porcentaje|edad|age|rating|puntaje|nota)/i.test(name);
}
function round(v) { return Math.round(v * 100) / 100; }
function shortNum(v) { return Math.abs(v) >= 1000 ? Math.round(v).toLocaleString("es") : String(Math.round(v * 10) / 10); }

const MONTHS = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];
function formatPeriod(key, byDay) {
  const [y, m, d] = key.split("-");
  return byDay ? `${+d} ${MONTHS[+m - 1]}` : `${MONTHS[+m - 1]} ${y}`;
}

function dateRangeOf(rows, idx) {
  let min = Infinity, max = -Infinity;
  for (const r of rows) { const d = toDate(r[idx]); if (d) { const t = d.getTime(); if (t < min) min = t; if (t > max) max = t; } }
  return Number.isFinite(min) ? { from: min, to: max } : null;
}

function detectCategory(text) {
  let best = ["General", 0];
  for (const [name, re] of FILE_CATEGORIES) {
    const hits = (text.match(new RegExp(re.source, "gi")) || []).length;
    if (hits > best[1]) best = [name, hits];
  }
  return best[0];
}

function detectYear(sheets, fileName) {
  const withDates = sheets.find((s) => s.dateRange);
  if (withDates) return new Date(withDates.dateRange.to).getUTCFullYear();
  const m = String(fileName).match(/(19|20)\d{2}/);
  return m ? +m[0] : new Date().getFullYear();
}
