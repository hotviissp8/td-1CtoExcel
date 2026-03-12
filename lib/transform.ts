import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';

const DEFAULT_KG   = 1.0;
const MAIN_COEF    = 3.5;
const SEC_COEF     = 1.1;
const RETAIL_MUL   = 1.14;

// ---- Helpers ----
function parseRu(s: string) { return parseFloat(s.replace(',', '.')); }
function L(idx: number)     { return XLSX.utils.encode_col(idx); }

/** Safe numeric parse — works for both raw numbers and formatted strings ("3,478.00", "6 175", "293,00") */
function safeNum(v: unknown): number | null {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return isNaN(v) ? null : v;
  let s = String(v).trim().replace(/\s+/g, ''); // strip spaces (RU thousands sep)
  if (s.includes(',') && s.includes('.')) {
    // both separators: whichever comes LAST is the decimal
    const lc = s.lastIndexOf(','), lp = s.lastIndexOf('.');
    s = lp > lc ? s.replace(/,/g, '') : s.replace(/\./g, '').replace(',', '.');
  } else if (s.includes(',')) {
    // comma only: if exactly 3 decimal digits → thousands sep, else decimal
    const [int, dec] = s.split(',');
    s = /^\d{3}$/.test(dec) ? int + dec : int + '.' + dec;
  }
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}

function norm(text: unknown): string {
  return String(text).toLowerCase().replace(/\n/g, ' ').trim();
}

function capFirst(s: string): string {
  return s ? s[0].toUpperCase() + s.slice(1) : s;
}

// ---- Dimension extraction ----
interface DimInfo {
  sortKey:  number;  // thickness / diameter (for sorting + hidden column)
  dim1:     number;  // width  (mm) — 0 if unknown
  dim2:     number;  // height / length (mm) — 0 if unknown
  dimLabel: string;  // Размер column: width×height or length only (NO thickness)
}

// Returns true if name indicates a rod/bar (sold by meter, round cross-section)
function isRod(name: string): boolean {
  return /[øØ°]/.test(name) || /\bд(?:иам\w*)?\s*\d/i.test(name) ||
    /\bстерж/i.test(name) || /\bпруток/i.test(name);
}

// Returns true if sold by the meter (', м' or similar at end)
function isByMeter(name: string): boolean {
  return /,\s*м\.?\s*$/i.test(name);
}

// Returns true if sold by the piece (', шт' at end)
function isByPiece(name: string): boolean {
  return /,\s*шт\.?\s*$/i.test(name);
}

function extractDimInfo(name: string): DimInfo {
  const byMeter = isByMeter(name);
  const byPiece = isByPiece(name);
  const rod     = isRod(name);

  // P1 – explicit Th×W×H     e.g. "2x1000x2000", "0,5x600x1200"
  let m = name.match(/\b(\d+(?:[,\.]\d+)?)\s*[xхXХ]\s*(\d{2,4})\s*[xхXХ]\s*(\d{3,5})\b/i);
  if (m) {
    const th = parseRu(m[1]), w = +m[2], h = +m[3];
    if (byMeter && !byPiece) {
      // sheet by meter: only width matters, length = 1000
      return { sortKey: th, dim1: w, dim2: 1000, dimLabel: `${w}x1000` };
    }
    return { sortKey: th, dim1: w, dim2: h, dimLabel: `${w}x${h}` };
  }

  // P1b – decimal×W  e.g. "0,40х1250мм …х1000мм"
  m = name.match(/\b(\d+[,\.]\d+)\s*[xхXХ]\s*(\d{3,5})/i);
  if (m) {
    const th = parseRu(m[1]), w = +m[2];
    if (byMeter && !byPiece) {
      return { sortKey: th, dim1: w, dim2: 1000, dimLabel: `${w}x1000` };
    }
    const rest = name.slice(name.indexOf(m[0]) + m[0].length);
    const hm = rest.match(/[xхXХ]\s*(\d{3,5})/i);
    const h = hm ? +hm[1] : 0;
    return { sortKey: th, dim1: w, dim2: h, dimLabel: h ? `${w}x${h}` : `${w}` };
  }

  // P2 – ø / °   e.g. "ø20x6000"
  m = name.match(/[øØ°]\s*(\d+(?:[,\.]\d+)?)\s*(?:[xхXХ]\s*(\d{3,6}))?/i);
  if (m) {
    const d = parseRu(m[1]), l = m[2] ? +m[2] : 0;
    // Rod by meter: only length (dim2), dim1=0
    return { sortKey: d, dim1: 0, dim2: l, dimLabel: l ? `${l}` : '' };
  }

  // P3 – "диаметр / д"   e.g. "диаметр 59,97мм"
  m = name.match(/\bд(?:иам\w*)?\s*(\d+(?:[,\.]\d+)?)/i);
  if (m) {
    const d = parseRu(m[1]);
    const lm = name.match(/[xхXХ]\s*(\d{3,6})/i);
    const l = lm ? +lm[1] : 0;
    return { sortKey: d, dim1: 0, dim2: l, dimLabel: l ? `${l}` : '' };
  }

  // P4 – thickness from first "NNмкм" or "NNмм",  then cross-dimension(s)
  const thM = name.match(/\b(\d+(?:[,\.]\d+)?)\s*(мкм|мм)/i);
  const whPairs = [...name.matchAll(/(\d{3,4})\s*[xхXХ]\s*(\d{3,5})/gi)];
  const allX    = [...name.matchAll(/[xхXХ]\s*(\d{3,5})/gi)].map(x => +x[1]);

  if (thM) {
    const rawTh = parseRu(thM[1]);
    const isMicron = thM[2].toLowerCase() === 'мкм';
    const th = isMicron ? rawTh / 1000 : rawTh;

    if (byMeter && !byPiece) {
      // Sold by meter
      if (rod) {
        // Rod: single dim = length (stock length in name)
        const l = allX.length > 0 ? allX[allX.length - 1] : 0;
        return { sortKey: th, dim1: 0, dim2: l, dimLabel: l ? `${l}` : '' };
      } else {
        // Sheet/roll by meter: extract width only, dim2=1000
        const w = allX.length > 0 ? allX[0] : 0;
        return { sortKey: th, dim1: w, dim2: 1000, dimLabel: w ? `${w}x1000` : '' };
      }
    }

    // byPiece or plain (no suffix)
    if (whPairs.length > 0) {
      const w = +whPairs[0][1], h = +whPairs[0][2];
      return { sortKey: th, dim1: w, dim2: h, dimLabel: `${w}x${h}` };
    }
    if (allX.length >= 2) {
      const [w, h] = allX;
      return { sortKey: th, dim1: w, dim2: h, dimLabel: `${w}x${h}` };
    }
    if (allX.length === 1) {
      const d = allX[0];
      if (byPiece) return { sortKey: th, dim1: 0, dim2: d, dimLabel: `${d}` };
      // No suffix, single dim → roll/film: width × 1000
      return { sortKey: th, dim1: d, dim2: 1000, dimLabel: `${d}x1000` };
    }
    return { sortKey: th, dim1: 0, dim2: 0, dimLabel: '' };
  }

  // P5 – plain W×H fallback (no thickness info)
  if (whPairs.length > 0) {
    const w = +whPairs[0][1], h = +whPairs[0][2];
    return { sortKey: w, dim1: w, dim2: h, dimLabel: `${w}x${h}` };
  }

  return { sortKey: 0, dim1: 0, dim2: 0, dimLabel: '' };
}

// ---- Weight extraction ----
function extractWeight(name: string): [number, boolean] {
  const ms = [...name.matchAll(/(\d+(?:,\d+)?)\s*кг/g)];
  if (ms.length > 0) return [parseFloat(ms[ms.length - 1][1].replace(',', '.')), false];
  return [DEFAULT_KG, true];
}

// ---- Column finding ----
function findColumn(
  ch: string[], kws: Array<string | [string, number]>, used: Set<number>
): number | null {
  const scored: Array<[string, number]> = kws.map(k => typeof k === 'string' ? [k, 10] : k);
  let best: number | null = null, bestS = 0;
  for (let i = 0; i < ch.length; i++) {
    if (used.has(i)) continue;
    const s = scored.reduce((a, [kw, sc]) => ch[i].includes(kw) ? a + sc : a, 0);
    if (s > bestS) { bestS = s; best = i; }
  }
  if (best !== null) used.add(best);
  return best;
}

function findAllStockCols(
  ch: string[], used: Set<number>
): Array<{ col: number; label: string }> {
  const res: Array<{ col: number; label: string }> = [];
  for (let i = 0; i < ch.length; i++) {
    if (used.has(i)) continue;
    const h = ch[i];
    if (!h.includes('остат')) continue;
    if (h.includes('налич') || h.includes('под заказ')) continue;
    let label = h
      .replace(/остат\w*/g, '').replace(/\brub\b/gi, '').replace(/руб\b/gi, '')
      .replace(/[,:;\s]+$/, '').replace(/^[,:;\s]+/, '').replace(/\s{2,}/g, ' ').trim();
    if (!label) label = `Склад ${res.length + 1}`;
    res.push({ col: i, label });
    used.add(i);
  }
  return res;
}

// ---- Main transform ----
export interface TransformResult { buffer: Buffer; warnings: string[]; rowCount: number; }

export async function transformExcel(buf: Buffer): Promise<TransformResult> {
  const warnings: string[] = [];

  // ---- Read with raw:true so numbers arrive as numbers, not formatted strings ----
  const wb = XLSX.read(buf, { type: 'buffer' });
  const raw: unknown[][] = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {
    header: 1, defval: '', raw: true,
  }) as unknown[][];

  // header row
  let headerRow: number | null = null;
  for (let i = 0; i < raw.length; i++) {
    if (raw[i].map(v => String(v)).join(' ').toLowerCase().includes('номенклатур')) {
      headerRow = i; break;
    }
  }
  if (headerRow === null) throw new Error('Строка заголовков не найдена');

  // data start row
  let dataStart = headerRow + 1;
  for (let j = headerRow + 1; j < Math.min(headerRow + 20, raw.length); j++) {
    if (/\b\d{3,}\b/.test(raw[j].map(v => String(v)).join(' '))) { dataStart = j; break; }
  }

  // header block start (scan backward for multi-row 1C headers)
  let hBlockStart = headerRow;
  for (let i = headerRow - 1; i >= Math.max(-1, headerRow - 8); i--) {
    const ne = raw[i].map(v => String(v).trim()).filter(Boolean);
    if (!ne.length) break;
    if (ne.some(v => /\b\d{3,}\b/.test(v))) break;
    hBlockStart = i;
  }

  // combined header (one string per column)
  const ncols = Math.max(...raw.slice(0, dataStart).map(r => r.length), 1);
  const ch: string[] = new Array(ncols).fill('');
  for (let ri = hBlockStart; ri < dataStart; ri++) {
    for (let ci = 0; ci < (raw[ri]?.length ?? 0); ci++) {
      const v = norm(raw[ri][ci]);
      if (v) ch[ci] = (ch[ci] + ' ' + v).trim();
    }
  }

  // ---- Detect input columns ----
  const used = new Set<number>();
  const nameCol = findColumn(ch, ['номенклатур', 'товар', 'продукт'], used);
  if (nameCol === null) throw new Error('Колонка с названием не обнаружена');

  // Retail: розничная → мелк/малый ×1.14 → оптов direct
  let retailSrc: number | null = null, retailMarkup = false;
  retailSrc = findColumn(ch, [['розниц', 20], ['розничная', 20]], used);
  if (retailSrc === null) {
    const sb = findColumn(ch, [['мелкийопт', 20], ['малый', 15], ['мелк', 10]], used);
    if (sb !== null) { retailSrc = sb; retailMarkup = true; }
  }
  if (retailSrc === null) retailSrc = findColumn(ch, [['оптов', 10], ['опт', 5]], used);

  if (retailSrc === null) warnings.push('Ни розничная, ни оптовая цена не найдена — столбец Розн пуст');
  else if (retailMarkup) warnings.push('Розничная цена не найдена — вычислена как Малый опт × 1,14');

  const stockCols = findAllStockCols(ch, used);
  if (stockCols.length === 0) warnings.push('Колонки остатка не найдены');

  // ---- Output column layout ----
  //
  //  A0   Номенклатура
  //  B1   Вес, кг
  //  C2   Розн
  //  D3   %
  //  E4   МП  = ROUNDUP(C*D, 0)
  //  F5   коэффициент
  //  G6   МП+ = ROUNDUP(E*F, 0)
  //  H7   Размер  (width × height only, no thickness)
  //  I8..I8+N-1    warehouse stock columns  (N ≥ 1)
  //  8+N           Цена за м²
  //  8+N+1  [hid]  dim1 (width, mm)
  //  8+N+2  [hid]  dim2 (height/length, mm)
  //  8+N+3  [hid]  Д/Т мм  (thickness / diameter — sort key)
  //  8+N+4         Ширина  (user input, blue)
  //  8+N+5         Длина   (user input, blue)
  //  8+N+6         цена    (formula)
  //  8+N+7..8+2N+6  шт_1 .. шт_N  (one per warehouse)

  const N    = Math.max(stockCols.length, 1);
  const BASE = 8;

  const ci = {
    name:    0, kg: 1, retail: 2, pct: 3, mp: 4, coef: 5, mpPlus: 6, dimStr: 7,
    // stock cols: BASE + 0 .. BASE + N-1
    sqm:       BASE + N,
    dim1:      BASE + N + 1,
    dim2:      BASE + N + 2,
    sortKey:   BASE + N + 3,
    width:     BASE + N + 4,
    length:    BASE + N + 5,
    price:     BASE + N + 6,
    // pieces: BASE + N + 7 + s  (s = 0..N-1)
  };
  const piecesBase = BASE + N + 7;
  const totalCols  = piecesBase + N;

  // header row
  const hdr: (string | null)[] = new Array(totalCols).fill(null);
  hdr[ci.name]   = 'Номенклатура';
  hdr[ci.kg]     = 'Вес, кг';
  hdr[ci.retail] = 'Розн';
  hdr[ci.pct]    = '%';
  hdr[ci.mp]     = 'МП';
  hdr[ci.coef]   = 'коэффициент';
  hdr[ci.mpPlus] = 'МП+';
  hdr[ci.dimStr] = 'Размер';
  for (let s = 0; s < N; s++) {
    hdr[BASE + s] = `${capFirst(stockCols[s]?.label ?? `Склад ${s+1}`)} ост-к`;
  }
  hdr[ci.sqm]     = 'Цена за м²';
  hdr[ci.dim1]    = 'dim1';
  hdr[ci.dim2]    = 'dim2';
  hdr[ci.sortKey] = 'Д/Т мм';
  hdr[ci.width]   = 'Ширина';
  hdr[ci.length]  = 'Длина';
  hdr[ci.price]   = 'цена';
  for (let s = 0; s < N; s++) {
    hdr[piecesBase + s] = `${capFirst(stockCols[s]?.label ?? `Склад ${s+1}`)} шт`;
  }

  // ---- Build data rows ----
  interface OutRow { data: (string | number | null)[]; kgDefaulted: boolean; dimInfo: DimInfo; }
  const outRows: OutRow[] = [];

  for (let ri = dataStart; ri < raw.length; ri++) {
    const row = raw[ri];
    const name = String(row[nameCol] ?? '').trim();
    if (!name) continue;

    let retail = retailSrc !== null ? safeNum(row[retailSrc]) : null;
    if (retail === null) continue; // category/separator row
    if (retailMarkup) retail = Math.ceil(retail * RETAIL_MUL);

    const stocks    = stockCols.map(({ col }) => safeNum(row[col]));
    const [kg, kgD] = extractWeight(name);
    const dimInfo   = extractDimInfo(name);

    const d: (string | number | null)[] = new Array(totalCols).fill(null);
    d[ci.name]    = name;
    d[ci.kg]      = kg;
    d[ci.retail]  = retail;
    d[ci.pct]     = MAIN_COEF;
    d[ci.coef]    = SEC_COEF;
    d[ci.dimStr]  = dimInfo.dimLabel || null;
    for (let s = 0; s < N; s++) d[BASE + s] = stocks[s] ?? null;
    d[ci.dim1]    = dimInfo.dim1 || null;
    d[ci.dim2]    = dimInfo.dim2 || null;
    d[ci.sortKey] = dimInfo.sortKey || null;

    outRows.push({ data: d, kgDefaulted: kgD, dimInfo });
  }

  // ---- Sort by thickness / diameter (ascending) ----
  outRows.sort((a, b) => a.dimInfo.sortKey - b.dimInfo.sortKey);

  // ---- Excel output ----
  const excelWb = new ExcelJS.Workbook();
  const ws = excelWb.addWorksheet('Sheet1');
  // Freeze row 1 (header) and column A (name)
  ws.views = [{ state: 'frozen', xSplit: 1, ySplit: 1 }];
  ws.addRow(hdr);
  for (const { data } of outRows) ws.addRow(data);

  // ---- Formulas ----
  for (let r = 2; r <= outRows.length + 1; r++) {
    const cC  = `${L(ci.retail)}${r}`;
    const cD  = `${L(ci.pct)}${r}`;
    const cE  = `${L(ci.mp)}${r}`;
    const cF  = `${L(ci.coef)}${r}`;
    const cG  = `${L(ci.mpPlus)}${r}`;
    const cB  = `B${r}`;
    const cI1 = `${L(ci.dim1)}${r}`;
    const cI2 = `${L(ci.dim2)}${r}`;
    const cW  = `${L(ci.width)}${r}`;
    const cL  = `${L(ci.length)}${r}`;

    // МП
    ws.getCell(cE).value = { formula: `ROUNDUP(${cC}*${cD},0)` } as ExcelJS.CellFormulaValue;
    // МП+
    ws.getCell(cG).value = { formula: `ROUNDUP(${cE}*${cF},0)` } as ExcelJS.CellFormulaValue;

    // Цена за м²:
    //   sheet/roll (dim1>0, dim2>0) → per m²
    //   rod        (dim1=0, dim2>0) → per п.м. (length only)
    //   neither                     → 0
    ws.getCell(`${L(ci.sqm)}${r}`).value = { formula:
      `IF(AND(${cI1}>0,${cI2}>0),` +
        `ROUNDUP(${cG}/((${cI1}/1000)*(${cI2}/1000)),0),` +
      `IF(AND(${cI1}=0,${cI2}>0),ROUNDUP(${cG}/(${cI2}/1000),0),0))`
    } as ExcelJS.CellFormulaValue;

    // цена:
    //   cut piece (both dims, user fills W+L): area ratio
    //   sheet by meter (dim1>0, dim2=1000, user fills W only): W/dim1 ratio
    //   rod (dim1=0, user fills L only): L/dim2 ratio
    ws.getCell(`${L(ci.price)}${r}`).value = { formula:
      `IF(AND(${cW}>0,${cL}>0,${cI1}>0,${cI2}>0),` +
        `ROUNDUP(${cG}*${cW}*${cL}/(${cI1}*${cI2}),0),` +
      `IF(AND(${cW}>0,${cI1}>0,${cI2}=1000),` +
        `ROUNDUP(${cG}*${cW}/${cI1},0),` +
      `IF(AND(${cL}>0,${cI1}=0,${cI2}>0),` +
        `ROUNDUP(${cG}*${cL}/${cI2},0),"")))`
    } as ExcelJS.CellFormulaValue;

    // шт per warehouse:
    //   cut piece (both dims): INT(stock/kg) × INT(dim1/W) × INT(dim2/L)
    //   sheet by meter (dim2=1000, W only): INT(stock/kg) × INT(dim1/W)
    //   rod (dim1=0, L only): INT(stock/kg) × INT(dim2/L)
    for (let s = 0; s < N; s++) {
      const cSt = `${L(BASE + s)}${r}`;
      ws.getCell(`${L(piecesBase + s)}${r}`).value = { formula:
        `IF(AND(${cSt}>0,${cB}>0,${cW}>0,${cL}>0,${cI1}>0,${cI2}>0),` +
          `INT(${cSt}/${cB})*INT(${cI1}/${cW})*INT(${cI2}/${cL}),` +
        `IF(AND(${cSt}>0,${cB}>0,${cW}>0,${cI1}>0,${cI2}=1000),` +
          `INT(${cSt}/${cB})*INT(${cI1}/${cW}),` +
        `IF(AND(${cSt}>0,${cB}>0,${cL}>0,${cI1}=0,${cI2}>0),` +
          `INT(${cSt}/${cB})*INT(${cI2}/${cL}),"")))`
      } as ExcelJS.CellFormulaValue;
    }
  }

  // ---- Styling ----
  const fillWarn   = { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'FFFFF2CC' } };
  const fillRed    = { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'FFFF0000' } };
  const fillMarkup = { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'FFE2EFDA' } };
  const fillInput  = { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'FFDCE6F1' } };

  outRows.forEach(({ kgDefaulted }, idx) => {
    if (kgDefaulted) ws.getCell(`B${idx + 2}`).fill = fillWarn;
  });
  if (retailMarkup) {
    ws.getRow(1).getCell(ci.retail + 1).fill = fillMarkup;
    for (let r = 2; r <= outRows.length + 1; r++) ws.getCell(`${L(ci.retail)}${r}`).fill = fillMarkup;
  }
  if (stockCols.length === 0) {
    for (let r = 2; r <= outRows.length + 1; r++) ws.getCell(`${L(BASE)}${r}`).fill = fillRed;
  }
  for (const idx of [ci.width, ci.length]) {
    ws.getRow(1).getCell(idx + 1).fill = fillInput;
    for (let r = 2; r <= outRows.length + 1; r++) ws.getCell(`${L(idx)}${r}`).fill = fillInput;
  }

  // font
  const font: Partial<ExcelJS.Font> = { name: 'Arial Black', size: 12 };
  ws.eachRow(row => row.eachCell({ includeEmpty: true }, cell => { cell.font = font; }));

  // number format on price columns
  for (const idx of [ci.retail, ci.mp, ci.mpPlus, ci.sqm, ci.price]) {
    for (let r = 2; r <= outRows.length + 1; r++) ws.getCell(`${L(idx)}${r}`).numFmt = '#,##0';
  }

  // hide internal columns
  for (const idx of [ci.dim1, ci.dim2, ci.sortKey]) {
    ws.getColumn(idx + 1).hidden = true;
  }

  // ---- Column widths: iterate explicitly to avoid ExcelJS iterator issues ----
  const widths: number[] = new Array(totalCols).fill(6);

  // header row (row 1)
  for (let c = 0; c < totalCols; c++) {
    const v = ws.getRow(1).getCell(c + 1).value;
    if (typeof v === 'string') widths[c] = Math.max(widths[c], v.length);
  }
  // data rows
  for (let r = 2; r <= outRows.length + 1; r++) {
    const row = ws.getRow(r);
    for (let c = 0; c < totalCols; c++) {
      const v = row.getCell(c + 1).value;
      if (v === null || v === undefined) continue;
      if (typeof v === 'object' && v !== null && 'formula' in v) continue;
      let len: number;
      if (typeof v === 'number') {
        const abs = Math.abs(Math.round(v)).toString();
        len = abs.length + Math.floor((abs.length - 1) / 3);
      } else {
        len = String(v).length;
      }
      widths[c] = Math.max(widths[c], len);
    }
  }
  for (let c = 0; c < totalCols; c++) {
    const col = ws.getColumn(c + 1);
    if (!col.hidden) col.width = Math.min(widths[c] + 2, 55);
  }

  const outBuf = await excelWb.xlsx.writeBuffer();
  return { buffer: Buffer.from(outBuf), warnings, rowCount: outRows.length };
}
