'use strict';

const ExcelJS = require('exceljs');
const path    = require('path');
const fs      = require('fs');

// ─── Константы ───────────────────────────────────────────────────────────────

function getZone(ratio) {
  if (ratio <= 1.1)  return 'auto';
  if (ratio < 1.5)   return 'orange';
  if (ratio < 2.0)   return 'red';
  return 'burgundy';
}

const ZONE_COLORS  = { orange: 'FFFFC000', red: 'FFFF0000', burgundy: 'FF800020' };
const ZONE_LABELS  = {
  auto:     'Авто (≤10%)',
  orange:   'Оранжевая (10–50%)',
  red:      'Красная (50–100%)',
  burgundy: 'Бордовая (>100%)',
};
const LIGHT_COLORS = { orange: 'FFFFF2CC', red: 'FFFFD7D7', burgundy: 'FFEDD5D5' };

const MR_HEADER = 'FFC6EFCE';
const MR_GREEN  = 'FFC6EFCE';
const MR_ORANGE = 'FFFFEB9C';
const MR_RED    = 'FFFFC7CE';
const MR_FONT   = 'FF375623';

const SECTIONS_ORDER = ['Трубы', 'Арматура', 'Изоляция', 'Насосы', 'Оборудование'];
const SECTION_PATTERNS = [
  { name: 'Трубы',    pattern: /труб/i },
  { name: 'Арматура', pattern: /арматур/i },
  { name: 'Изоляция', pattern: /изоляц/i },
  { name: 'Насосы',   pattern: /насос/i },
];

// ─── Вспомогательные функции ─────────────────────────────────────────────────

/** Извлекает текст из ячейки любого типа: plain, formula, RichText, Date */
function cellText(cell) {
  if (!cell || cell.value == null) return '';
  const v = cell.value;
  if (typeof v === 'object') {
    if (Array.isArray(v.richText)) return v.richText.map(r => r.text || '').join('');
    if (v.result   !== undefined)  return v.result  != null ? String(v.result)  : '';
    if (v.text     !== undefined)  return String(v.text);
    if (v instanceof Date)         return v.toISOString();
  }
  return String(v);
}

/** Нормализует строку числа: убирает пробел-разрядник, меняет запятую на точку */
function parseNumStr(s) {
  const cleaned = String(s).trim().replace(/\s/g, '').replace(',', '.');
  const n = parseFloat(cleaned);
  return isNaN(n) ? null : n;
}

function cellNum(cell) {
  if (!cell || cell.value == null) return null;
  const v = cell.value;
  if (typeof v === 'number')  return isNaN(v) ? null : v;
  if (typeof v === 'object') {
    if (Array.isArray(v.richText))
      return parseNumStr(v.richText.map(r => r.text || '').join(''));
    if (v.result !== undefined) {
      if (typeof v.result === 'number') return isNaN(v.result) ? null : v.result;
      return v.result != null ? parseNumStr(String(v.result)) : null;
    }
  }
  return parseNumStr(String(v));
}

function cellStr(cell) { return cellText(cell).trim(); }

function fillCell(cell, argb) {
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb } };
}

function borderCell(cell) {
  const side = { style: 'thin', color: { argb: MR_FONT } };
  cell.border = { top: side, left: side, bottom: side, right: side };
}

function detectSection(classification, name) {
  for (const src of [classification, name]) {
    if (!src) continue;
    for (const s of SECTION_PATTERNS)
      if (s.pattern.test(src.toString())) return s.name;
  }
  return 'Оборудование';
}

function emptyGroups() {
  const g = {};
  for (const s of SECTIONS_ORDER) g[s] = [];
  return g;
}

// ─── Поиск строки заголовков ─────────────────────────────────────────────────

function findHeaderRow(sheet) {
  const PASSES = [
    /цена за единицу материала\s*\d{4}/i,
    /цена за единицу\s*\d{4}/i,
    /цена за ед\.?\s*\d{4}/i,
    /цена\s+20\d{2}/i,
    /сумма\s*разница/i,
    /разница/i,
  ];

  for (let pass = 0; pass < PASSES.length; pass++) {
    const p = PASSES[pass];
    for (let r = 1; r <= 25; r++) {
      const vals = [];
      sheet.getRow(r).eachCell({ includeEmpty: true }, (cell, col) => {
        vals[col - 1] = cellText(cell);
      });
      if (vals.some(v => p.test(v))) {
        console.log(`  findHeaderRow: строка ${r} (pass ${pass + 1}) по ${p}`);
        return { rowNum: r, headers: vals };
      }
    }
  }

  console.error(`  [findHeaderRow] ОШИБКА: заголовки НЕ найдены на листе "${sheet.name}"`);
  for (let r = 1; r <= Math.min(10, sheet.rowCount); r++) {
    const parts = [];
    sheet.getRow(r).eachCell({ includeEmpty: false }, (c, col) => {
      const t = cellText(c).substring(0, 40);
      if (t) parts.push(`кол.${col}: "${t}"`);
    });
    if (parts.length) console.error(`  Строка ${r}: ${parts.join(' | ')}`);
  }

  return null;
}

// ─── Поиск ценовых столбцов ──────────────────────────────────────────────────

function detectPriceCols(headers, maxCol = Infinity) {
  const TIERS = [
    /цена\s+(?:за\s+)?(?:единицу\s+)?(?:материала\s+)?(\d{4})/i,
    /цена[^0-9]{0,25}(\d{4})/i,
    /(\d{4})[^0-9]{0,15}цена/i,
  ];

  const found = [];
  const seen  = new Set();

  headers.forEach((h, i) => {
    if (!h) return;
    const colNum = i + 1;
    if (colNum >= maxCol) return;
    const str = String(h);
    for (const p of TIERS) {
      const m = str.match(p);
      if (m) {
        const year = parseInt(m[1]);
        if (year >= 2015 && year <= 2030 && !seen.has(colNum)) {
          seen.add(colNum);
          found.push({ colNum, year, header: str });
        }
        break;
      }
    }
  });

  found.sort((a, b) => a.year - b.year);
  return found;
}

// ─── Поиск вспомогательных столбцов ──────────────────────────────────────────

function findColInSheet(sheet, patterns, maxRow = 15) {
  for (let r = 1; r <= maxRow; r++) {
    let found = null;
    sheet.getRow(r).eachCell({ includeEmpty: true }, (cell, col) => {
      if (found) return;
      const str = cellText(cell).toLowerCase().replace(/\s+/g, ' ').trim();
      if (str && patterns.some(p => str.includes(p))) found = col;
    });
    if (found) return found;
  }
  return null;
}

// ─── Защита от повторного анализа ────────────────────────────────────────────

function findExistingMRBlock(sheet, headerRowNum) {
  // Ищем "Анализ MR Group" во всех строках от 1 до headerRowNum включительно.
  // Не используем "Цена за единицу материала 2025" как признак MR-блока —
  // этот текст может присутствовать в оригинальном файле как колонка "Стало".
  for (let r = 1; r <= headerRowNum; r++) {
    let found = null;
    sheet.getRow(r).eachCell({ includeEmpty: false }, (cell, col) => {
      if (!found && cellText(cell).includes('Анализ MR Group')) found = col;
    });
    if (found) return found;
  }
  return null;
}

// ─── Последняя заполненная колонка ───────────────────────────────────────────

function findLastFilledCol(sheet, headerRowNum, maxCol = Infinity) {
  let last = 0;
  sheet.getRow(headerRowNum).eachCell({ includeEmpty: false }, (_, col) => {
    if (col > last && col < maxCol) last = col;
  });
  try {
    const merges = sheet._merges;
    if (merges) {
      Object.values(merges).forEach(m => {
        const top    = m.top    ?? m.model?.top;
        const bottom = m.bottom ?? m.model?.bottom;
        const right  = m.right  ?? m.model?.right;
        if (top <= headerRowNum && bottom >= headerRowNum && right > last && right < maxCol)
          last = right;
      });
    }
  } catch (e) {
    console.warn('  findLastFilledCol: ошибка слияний:', e.message);
  }
  return last || sheet.columnCount || 1;
}

// ─── Исправление shared-формул ───────────────────────────────────────────────

function resolveSharedFormulas(workbook) {
  workbook.worksheets.forEach(sheet => {
    sheet.eachRow({ includeEmpty: false }, row => {
      row.eachCell({ includeEmpty: false }, cell => {
        const v = cell.value;
        if (v && typeof v === 'object') {
          if (v.sharedFormula !== undefined)
            cell.value = v.result !== undefined ? v.result : null;
          if (v.shareType === 'shared' && v.formula)
            cell.value = { formula: v.formula, result: v.result };
        }
      });
    });
  });
}

// ─── Обработка одного листа ──────────────────────────────────────────────────

async function processSheet(sheet) {
  console.log(`\n  === Лист: "${sheet.name}" (строк: ${sheet.rowCount}, колонок: ${sheet.columnCount}) ===`);

  const headerInfo = findHeaderRow(sheet);
  if (!headerInfo) throw new Error('не найдена строка с заголовками цен');

  const { rowNum: headerRowNum, headers } = headerInfo;
  console.log(`  Строка заголовков: ${headerRowNum}`);

  const existingMRCol = findExistingMRBlock(sheet, headerRowNum);
  if (existingMRCol) console.warn(`  MR-блок уже есть (кол.${existingMRCol}) — пересчёт`);
  const colLimit = existingMRCol ?? Infinity;

  const priceCols = detectPriceCols(headers, colLimit);
  console.log(`  Ценовых столбцов: ${priceCols.length}`, priceCols.map(p => `[${p.year}→кол${p.colNum}]`).join(' '));

  if (priceCols.length < 2) {
    console.error(`  ОШИБКА: нужно ≥2 ценовых столбца, найдено ${priceCols.length}`);
    console.error(`  Заголовки строки ${headerRowNum}: ${headers.filter(Boolean).slice(0, 20).join(' | ')}`);
    throw new Error(`найдено менее двух ценовых столбцов (${priceCols.length})`);
  }

  const wasCol    = priceCols[0];
  const becameCol = priceCols[priceCols.length - 1];

  const qtyOrientirCol = findColInSheet(sheet, ['кол-во ориентир', 'кол.ориентир', 'количество ориентир', 'ориентир']);
  const nameCol  = findColInSheet(sheet, ['наименование', 'название материала', 'наим.', 'номенклатура']);
  const classCol = findColInSheet(sheet, ['классификация']);
  const unitCol  = findColInSheet(sheet, ['ед. изм', 'ед.изм', 'единица']);
  const qtyCol   = findColInSheet(sheet, ['кол-во', 'количество', 'кол.']);
  const sumQtyColIdx = qtyOrientirCol || qtyCol || null;

  console.log(`  Было: кол.${wasCol.colNum} [${wasCol.year}]  Стало: кол.${becameCol.colNum} [${becameCol.year}]`);
  console.log(`  Кол-во ориентир: кол.${qtyOrientirCol || '—'}  fallback: кол.${qtyCol || '—'}  Наименование: кол.${nameCol || '—'}`);

  const lastFilledCol = findLastFilledCol(sheet, headerRowNum, colLimit);
  const mrPriceColIdx = lastFilledCol + 1;
  const mrSumColIdx   = lastFilledCol + 2;

  console.log(`  Выделены колонки для Анализа MR: mrPriceCol=${mrPriceColIdx}, mrSumCol=${mrSumColIdx}`);

  const becameLetter  = sheet.getColumn(becameCol.colNum).letter;
  const mrPriceLetter = sheet.getColumn(mrPriceColIdx).letter;
  const sumQtyLetter  = sumQtyColIdx ? sheet.getColumn(sumQtyColIdx).letter : null;

  // ── Шапка MR-блока ──────────────────────────────────────────────────────
  if (headerRowNum > 1) {
    try { sheet.mergeCells(headerRowNum - 1, mrPriceColIdx, headerRowNum - 1, mrSumColIdx); } catch {}
    const groupCell = sheet.getRow(headerRowNum - 1).getCell(mrPriceColIdx);
    groupCell.value     = 'Анализ MR Group';
    groupCell.font      = { bold: true, color: { argb: MR_FONT } };
    groupCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    fillCell(groupCell, MR_HEADER);
    borderCell(groupCell);
  }

  const hRow = sheet.getRow(headerRowNum);

  const priceHdr = hRow.getCell(mrPriceColIdx);
  priceHdr.value     = 'Цена за единицу материала 2025';
  priceHdr.font      = { bold: true, color: { argb: MR_FONT } };
  priceHdr.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  fillCell(priceHdr, MR_HEADER);
  borderCell(priceHdr);

  const sumHdr = hRow.getCell(mrSumColIdx);
  sumHdr.value     = 'Сумма материалов 2025';
  sumHdr.font      = { bold: true, color: { argb: MR_FONT } };
  sumHdr.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  fillCell(sumHdr, MR_HEADER);
  borderCell(sumHdr);

  sheet.getColumn(mrPriceColIdx).width = 22;
  sheet.getColumn(mrSumColIdx).width   = 20;

  // ── Цикл по строкам данных ───────────────────────────────────────────────
  const groups = emptyGroups();
  let processed = 0, autoFilled = 0, toRequest = 0;

  console.log(`  КОЛОНКИ: nameCol=${nameCol || 'null'}  unitCol=${unitCol || 'null'}  qtyCol=${qtyCol || 'null'}  sumQtyColIdx=${sumQtyColIdx || 'null'}  wasCol=${wasCol.colNum}  becameCol=${becameCol.colNum}`);

  for (let rowNum = headerRowNum + 1; rowNum <= sheet.rowCount; rowNum++) {
    const row = sheet.getRow(rowNum);

    // ── ДИАГНОСТИКА: дамп первых 3 строк данных ─────────────────────────
    if (rowNum <= headerRowNum + 3) {
      console.log(`\n  === ДИАГНОСТИКА СТРОКИ ${rowNum} ===`);
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        console.log(`    кол.${colNumber} [type=${cell.type}]: ${JSON.stringify(cell.value)}`);
      });
      console.log(`  raw was=${JSON.stringify(row.getCell(wasCol.colNum).value)}  raw became=${JSON.stringify(row.getCell(becameCol.colNum).value)}`);
      console.log(`  cellNum(was)=${cellNum(row.getCell(wasCol.colNum))}  cellNum(became)=${cellNum(row.getCell(becameCol.colNum))}`);
      console.log(`  ============================`);
    }

    // ── Читаем ключевые поля строки напрямую ─────────────────────────────
    const name        = nameCol ? cellStr(row.getCell(nameCol)) : '';
    const wasPrice    = cellNum(row.getCell(wasCol.colNum));
    const becamePrice = cellNum(row.getCell(becameCol.colNum));

    // Пропускаем строки без названия И без хотя бы одной цены
    if (!name.trim() && wasPrice == null && becamePrice == null) continue;
    // Пропускаем итоговые строки ("ИТОГО", "ВСЕГО", "TOTAL")
    if (/^\s*(итого|всего|total)\s*$/i.test(name)) continue;

    processed++;

    console.log(`  Строка ${rowNum}: "${name.substring(0, 50)}" was=${wasPrice} became=${becamePrice}`);

    const mrPriceCell = row.getCell(mrPriceColIdx);
    const mrSumCell   = row.getCell(mrSumColIdx);

    // ── Формула суммы (для любой значимой строки) ─────────────────────────
    if (sumQtyLetter) {
      mrSumCell.value = { formula: `${mrPriceLetter}${rowNum}*${sumQtyLetter}${rowNum}` };
    }
    mrSumCell.numFmt = '#,##0.00 ₽';
    borderCell(mrSumCell);

    // ── Если цены неизвестны → красная ячейка, строка на запрос ──────────
    if (!wasPrice || !becamePrice || wasPrice <= 0 || becamePrice <= 0) {
      mrPriceCell.value  = null;
      mrPriceCell.numFmt = '#,##0.00 ₽';
      fillCell(mrPriceCell, MR_RED);
      borderCell(mrPriceCell);

      const classification = classCol ? cellStr(row.getCell(classCol)) : '';
      const unit           = unitCol  ? cellStr(row.getCell(unitCol))  : '';
      const qty            = qtyCol   ? cellNum(row.getCell(qtyCol))   : null;
      const section        = detectSection(classification, name);

      groups[section].push({ name, classification, unit, qty,
                             wasPrice: wasPrice || 0, becamePrice: becamePrice || 0,
                             ratio: null, zone: 'red' });
      toRequest++;
      console.log(`  ✗ Строка ${rowNum}: нет цен → на запрос (красная)`);
      continue;
    }

    // ── Есть обе цены → считаем зону ─────────────────────────────────────
    const ratio = becamePrice / wasPrice;
    const pct   = (ratio - 1) * 100;
    const zone  = getZone(ratio);

    if (pct <= 10 + 1e-9) {
      mrPriceCell.value  = { formula: `${becameLetter}${rowNum}` };
      mrPriceCell.numFmt = '#,##0.00 ₽';
      fillCell(mrPriceCell, MR_GREEN);
      borderCell(mrPriceCell);
      autoFilled++;
      console.log(`  ✓ Строка ${rowNum}: авто (${pct.toFixed(1)}%)`);
    } else {
      mrPriceCell.value  = null;
      mrPriceCell.numFmt = '#,##0.00 ₽';
      fillCell(mrPriceCell, pct <= 40 ? MR_ORANGE : MR_RED);
      borderCell(mrPriceCell);

      fillCell(row.getCell(wasCol.colNum),    ZONE_COLORS[zone]);
      fillCell(row.getCell(becameCol.colNum), ZONE_COLORS[zone]);

      const classification = classCol ? cellStr(row.getCell(classCol)) : '';
      const unit           = unitCol  ? cellStr(row.getCell(unitCol))  : '';
      const qty            = qtyCol   ? cellNum(row.getCell(qtyCol))   : null;
      const section        = detectSection(classification, name);

      groups[section].push({ name, classification, unit, qty, wasPrice, becamePrice, ratio, zone });
      toRequest++;
      console.log(`  ⚠ Строка ${rowNum}: на запрос (${pct.toFixed(1)}%, зона: ${zone})`);
    }
  }

  console.log(`  Итого → Обработано: ${processed}  Авто: ${autoFilled}  На запрос: ${toRequest}`);

  // Если ноль — принудительный дамп первых 10 строк для диагностики
  if (processed === 0) {
    console.log(`  ⚠ НУЛЬ СТРОК! Дамп строк ${headerRowNum+1}…${Math.min(headerRowNum+10, sheet.rowCount)}:`);
    for (let r = headerRowNum + 1; r <= Math.min(headerRowNum + 10, sheet.rowCount); r++) {
      const parts = [];
      sheet.getRow(r).eachCell({ includeEmpty: false }, (cell, col) => {
        parts.push(`кол${col}[t=${cell.type}]:${JSON.stringify(cell.value)?.substring(0,40)}`);
      });
      console.log(`  Строка ${r}: ${parts.length ? parts.join(' | ') : '(пустая)'}`);
    }
  }

  return { groups, processed, autoFilled, toRequest };
}

// ─── Сборка файла запроса в workbook ─────────────────────────────────────────

async function buildRequestWorkbook(allGroups) {
  const reqWb = new ExcelJS.Workbook();

  const summary = reqWb.addWorksheet('Сводка');
  const sh = summary.addRow(['Раздел', 'Кол-во позиций']);
  sh.font = { bold: true };
  sh.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } };

  let total = 0;
  for (const sec of SECTIONS_ORDER) {
    const n = allGroups[sec].length;
    if (n > 0) { summary.addRow([sec, n]); total += n; }
  }
  summary.addRow([]);
  const tr = summary.addRow(['Итого на запрос:', total]);
  tr.font = { bold: true };
  summary.columns = [{ width: 25 }, { width: 20 }];

  for (const sec of SECTIONS_ORDER) {
    const items = allGroups[sec];
    if (!items.length) continue;

    const s = reqWb.addWorksheet(sec);
    const hRow = s.addRow(['Наименование', 'Классификация', 'Ед. изм.', 'Кол-во',
                           'Цена было', 'Цена стало', 'Коэф.', 'Зона']);
    hRow.font = { bold: true };
    hRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } };

    for (const item of items) {
      const dr = s.addRow([
        item.name, item.classification, item.unit, item.qty,
        item.wasPrice, item.becamePrice,
        item.ratio != null ? parseFloat(item.ratio.toFixed(3)) : null,
        ZONE_LABELS[item.zone],
      ]);
      dr.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: LIGHT_COLORS[item.zone] } };
    }

    s.columns.forEach(col => {
      let max = 10;
      col.eachCell({ includeEmpty: false }, cell => {
        const len = cellText(cell).length;
        if (len > max) max = len;
      });
      col.width = Math.min(max + 4, 60);
    });
  }

  return reqWb;
}

// ─── БУФЕРНЫЙ АНАЛИЗ (используется сервером) ─────────────────────────────────

/**
 * Принимает буфер исходного xlsx, обрабатывает в памяти,
 * возвращает { analysisBuffer, requestBuffer } — готовые к загрузке в Supabase.
 */
async function analyzeBuffer(srcBuffer) {
  console.log(`\n##MARKER_V7## analyzeBuffer: ${srcBuffer.length} байт`);

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(srcBuffer);                         // ← load из буфера

  if (!wb.worksheets.length) throw new Error('Файл не содержит листов.');
  console.log(`Листов: ${wb.worksheets.length} (${wb.worksheets.map(s => `"${s.name}"`).join(', ')})`);

  const allGroups = emptyGroups();
  let totalProcessed = 0, totalAuto = 0, totalRequest = 0;

  for (const sheet of wb.worksheets) {
    try {
      const { groups, processed, autoFilled, toRequest } = await processSheet(sheet);
      for (const sec of SECTIONS_ORDER) allGroups[sec].push(...groups[sec]);
      totalProcessed += processed;
      totalAuto      += autoFilled;
      totalRequest   += toRequest;
    } catch (e) {
      console.error(`  ⛔ Пропуск листа "${sheet.name}": ${e.message}`);
    }
  }

  resolveSharedFormulas(wb);

  // ── Критический шаг: сериализация в новый буфер ───────────────────────
  const analysisBuffer = Buffer.from(await wb.xlsx.writeBuffer());
  console.log(`Файл успешно модифицирован, размер буфера: ${analysisBuffer.length} байт`);

  const reqWb = await buildRequestWorkbook(allGroups);
  const requestBuffer = Buffer.from(await reqWb.xlsx.writeBuffer());
  console.log(`Файл запроса сгенерирован, размер буфера: ${requestBuffer.length} байт`);

  console.log('\n=== Итого ===');
  console.log(`Обработано строк:       ${totalProcessed}`);
  console.log(`Заполнено авто (≤10%):  ${totalAuto}`);
  console.log(`На запрос уточнения:    ${totalRequest}`);

  return { analysisBuffer, requestBuffer };
}

// ─── ФАЙЛОВЫЙ АНАЛИЗ (CLI) ────────────────────────────────────────────────────

/**
 * Анализ из файла → два файла на диске. Используется только для CLI.
 */
async function analyze(inputPath) {
  const srcBuffer = fs.readFileSync(inputPath);
  const { analysisBuffer, requestBuffer } = await analyzeBuffer(srcBuffer);

  const dir     = path.dirname(inputPath);
  const base    = path.basename(inputPath, path.extname(inputPath));
  const outMain = path.join(dir, `${base}_analysis.xlsx`);
  const outReq  = path.join(dir, `${base}_request.xlsx`);

  fs.writeFileSync(outMain, analysisBuffer);
  fs.writeFileSync(outReq,  requestBuffer);

  console.log(`\nФайл анализа: ${outMain}`);
  console.log(`Файл запроса: ${outReq}`);

  return { analysisPath: outMain, requestPath: outReq };
}

module.exports = analyze;
module.exports.analyzeBuffer = analyzeBuffer;

// ─── Запуск из командной строки ──────────────────────────────────────────────

if (require.main === module) {
  const inputFile = process.argv[2];
  if (!inputFile) {
    console.error('Использование: node analyze.js <путь_к_файлу.xlsx>');
    process.exit(1);
  }
  analyze(path.resolve(inputFile)).catch(err => {
    console.error('\nОшибка:', err.message);
    process.exit(1);
  });
}
