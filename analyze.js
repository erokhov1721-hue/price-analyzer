const ExcelJS = require('exceljs');
const path = require('path');

// ─── Константы ───────────────────────────────────────────────────────────────

/** Зоны удорожания для окраски ИСХОДНЫХ ценовых столбцов */
function getZone(ratio) {
  if (ratio <= 1.1)  return 'auto';
  if (ratio < 1.5)   return 'orange';
  if (ratio < 2.0)   return 'red';
  return 'burgundy';
}

const ZONE_COLORS = {
  orange:   'FFFFC000',
  red:      'FFFF0000',
  burgundy: 'FF800020',
};
const ZONE_LABELS = {
  auto:     'Авто (≤10%)',
  orange:   'Оранжевая (10–50%)',
  red:      'Красная (50–100%)',
  burgundy: 'Бордовая (>100%)',
};
const LIGHT_COLORS = {
  orange:   'FFFFF2CC',
  red:      'FFFFD7D7',
  burgundy: 'FFEDD5D5',
};

/** Цвета блока «Анализ MR Group» */
const MR_HEADER = 'FFC6EFCE'; // шапка (светло-зелёный Portal)
const MR_GREEN  = 'FFC6EFCE'; // ≤10% — ячейка с формулой
const MR_ORANGE = 'FFFFEB9C'; // 10–40% — пусто, менеджер заполняет
const MR_RED    = 'FFFFC7CE'; // >40% — пусто, менеджер заполняет
const MR_FONT   = 'FF375623'; // тёмно-зелёный шрифт шапки

const SECTIONS_ORDER = ['Трубы', 'Арматура', 'Изоляция', 'Насосы', 'Оборудование'];

const SECTION_PATTERNS = [
  { name: 'Трубы',    pattern: /труб/i },
  { name: 'Арматура', pattern: /арматур/i },
  { name: 'Изоляция', pattern: /изоляц/i },
  { name: 'Насосы',   pattern: /насос/i },
];

// ─── Вспомогательные функции ─────────────────────────────────────────────────

function cellNum(cell) {
  if (!cell) return null;
  const v = cell.value;
  if (v === null || v === undefined) return null;
  if (typeof v === 'object' && v.result !== undefined) return parseFloat(v.result) || null;
  const n = parseFloat(v);
  return isNaN(n) ? null : n;
}

function cellStr(cell) {
  if (!cell) return '';
  const v = cell.value;
  if (v === null || v === undefined) return '';
  if (typeof v === 'object' && v.result !== undefined)
    return v.result != null ? String(v.result).trim() : '';
  return String(v).trim();
}

function fillCell(cell, argb) {
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb } };
}

const MR_BORDER_COLOR = 'FF375623';
function borderCell(cell) {
  const side = { style: 'thin', color: { argb: MR_BORDER_COLOR } };
  cell.border = { top: side, left: side, bottom: side, right: side };
}

function detectSection(classification, name) {
  for (const src of [classification, name]) {
    if (!src) continue;
    for (const s of SECTION_PATTERNS) {
      if (s.pattern.test(src.toString())) return s.name;
    }
  }
  return 'Оборудование';
}

function emptyGroups() {
  const g = {};
  for (const s of SECTIONS_ORDER) g[s] = [];
  return g;
}

// ─── Поиск столбцов в листе ──────────────────────────────────────────────────

/** Возвращает номер первого столбца, заголовок которого содержит любой из паттернов */
function findColInSheet(sheet, patterns, maxRow = 10) {
  for (let r = 1; r <= maxRow; r++) {
    const row = sheet.getRow(r);
    let found = null;
    row.eachCell({ includeEmpty: true }, (cell, col) => {
      if (found || !cell.value) return;
      const str = cell.value.toString().toLowerCase();
      if (patterns.some(p => str.includes(p))) found = col;
    });
    if (found) return found;
  }
  return null;
}

/** Ищет строку, содержащую «Цена за единицу материала YYYY» */
function findHeaderRow(sheet) {
  for (let r = 1; r <= 10; r++) {
    const row = sheet.getRow(r);
    const vals = [];
    row.eachCell({ includeEmpty: true }, (cell, col) => {
      vals[col - 1] = cell.value ? cell.value.toString() : '';
    });
    if (vals.some(v => /цена за единицу материала\s+\d{4}/i.test(v))) {
      return { rowNum: r, headers: vals };
    }
  }
  return null;
}

/** Возвращает массив ценовых столбцов [{colNum, year}], отсортированных по году.
 *  @param {number} [maxCol=Infinity] — столбцы >= maxCol игнорируются (защита от re-analysis) */
function detectPriceCols(headers, maxCol = Infinity) {
  const pattern = /цена за единицу материала\s+(\d{4})/i;
  const found = [];
  headers.forEach((h, i) => {
    if (!h) return;
    const colNum = i + 1;
    if (colNum >= maxCol) return;
    const m = h.toString().match(pattern);
    if (m) found.push({ colNum, year: parseInt(m[1]), header: h.toString() });
  });
  found.sort((a, b) => a.year - b.year);
  return found;
}

/** Если лист уже содержит блок «Анализ MR Group», возвращает номер его первого столбца,
 *  иначе — null. Используется для защиты от повторного анализа. */
function findExistingMRBlock(sheet, headerRowNum) {
  if (headerRowNum > 1) {
    let found = null;
    sheet.getRow(headerRowNum - 1).eachCell({ includeEmpty: false }, (cell, col) => {
      if (!found && cell.value && cell.value.toString().includes('Анализ MR Group')) {
        found = col;
      }
    });
    if (found) return found;
  }
  // Также проверим строку заголовков
  let found = null;
  sheet.getRow(headerRowNum).eachCell({ includeEmpty: false }, (cell, col) => {
    if (!found && cell.value && cell.value.toString().includes('Цена за единицу материала 2025')) {
      found = col;
    }
  });
  return found;
}

/** Номер последнего непустого столбца в строке headerRowNum.
 *  @param {number} [maxCol=Infinity] — столбцы >= maxCol игнорируются */
function findLastFilledCol(sheet, headerRowNum, maxCol = Infinity) {
  let last = 0;
  sheet.getRow(headerRowNum).eachCell({ includeEmpty: false }, (_, col) => {
    if (col > last && col < maxCol) last = col;
  });
  // Доп. попытка через диапазоны объединений (защита от merged slave-ячеек)
  try {
    const merges = sheet._merges; // internal ExcelJS map
    if (merges) {
      Object.values(merges).forEach(m => {
        const top    = typeof m === 'object' ? (m.top    || m.model?.top)    : null;
        const bottom = typeof m === 'object' ? (m.bottom || m.model?.bottom) : null;
        const right  = typeof m === 'object' ? (m.right  || m.model?.right)  : null;
        if (top <= headerRowNum && bottom >= headerRowNum && right > last && right < maxCol) last = right;
      });
    }
  } catch {}
  return last || sheet.columnCount || 1;
}

// ─── Исправление shared-формул перед сохранением ─────────────────────────────

function resolveSharedFormulas(workbook) {
  workbook.worksheets.forEach(sheet => {
    sheet.eachRow({ includeEmpty: false }, row => {
      row.eachCell({ includeEmpty: false }, cell => {
        const v = cell.value;
        if (v && typeof v === 'object') {
          if (v.sharedFormula !== undefined) {
            cell.value = v.result !== undefined ? v.result : null;
          }
          if (v.shareType === 'shared' && v.formula) {
            cell.value = { formula: v.formula, result: v.result };
          }
        }
      });
    });
  });
}

// ─── Обработка одного листа ──────────────────────────────────────────────────

async function processSheet(sheet) {
  // 1. Найти строку с заголовками цен
  const headerInfo = findHeaderRow(sheet);
  if (!headerInfo) throw new Error('не найдена строка с заголовками цен');

  const { rowNum: headerRowNum, headers } = headerInfo;

  // 2. Защита от повторного анализа: пропустить уже существующий MR-блок
  const existingMRCol = findExistingMRBlock(sheet, headerRowNum);
  if (existingMRCol) {
    console.warn(`  Лист "${sheet.name}": блок «Анализ MR Group» уже присутствует (кол.${existingMRCol}) — пересчёт`);
  }

  // Ограничиваем поиск столбцами левее существующего MR-блока (если он есть)
  const colLimit = existingMRCol ?? Infinity;

  const priceCols = detectPriceCols(headers, colLimit);
  if (priceCols.length < 2) throw new Error('найдено менее двух ценовых столбцов');

  const wasCol    = priceCols[0];                        // самый ранний год
  const becameCol = priceCols[priceCols.length - 1];    // самый поздний год

  // 3. Найти «Кол-во ориентир» (для формулы суммы в MR-блоке)
  const qtyOrientirCol = findColInSheet(sheet,
    ['кол-во ориентир', 'кол.ориентир', 'количество ориентир', 'ориентир']);

  // Вспомогательные столбцы
  const nameCol  = findColInSheet(sheet, ['наименование', 'название материала']);
  const classCol = findColInSheet(sheet, ['классификация']);
  const unitCol  = findColInSheet(sheet, ['ед. изм', 'ед.изм', 'единица']);
  const qtyCol   = findColInSheet(sheet, ['кол-во', 'количество', 'кол.']);

  // Для суммы используем «Кол-во ориентир»; если не нашли — обычное «Кол-во»
  const sumQtyColIdx = qtyOrientirCol || qtyCol || null;

  console.log(`\nЛист: "${sheet.name}"`);
  console.log(`  Было: кол.${wasCol.colNum} [${wasCol.year}]  Стало: кол.${becameCol.colNum} [${becameCol.year}]`);
  console.log(`  Кол-во ориентир: кол.${qtyOrientirCol || '—'}  (fallback кол-во: кол.${qtyCol || '—'})`);

  // 4. Позиции нового MR-блока — правее последней заполненной колонки (левее любого старого MR-блока)
  const lastFilledCol = findLastFilledCol(sheet, headerRowNum, colLimit);
  const mrPriceColIdx = lastFilledCol + 1;
  const mrSumColIdx   = lastFilledCol + 2;

  // Буквенные обозначения для Excel-формул
  const becameLetter   = sheet.getColumn(becameCol.colNum).letter;
  const mrPriceLetter  = sheet.getColumn(mrPriceColIdx).letter;
  const sumQtyLetter   = sumQtyColIdx ? sheet.getColumn(sumQtyColIdx).letter : null;

  console.log(`  Новый MR-блок → кол.${mrPriceColIdx} («${mrPriceLetter}»), кол.${mrSumColIdx}`);

  // 5. Шапка MR-блока
  // 5a. Объединённый заголовок группы (строка ВЫШЕHeaderRow, если есть)
  if (headerRowNum > 1) {
    try {
      sheet.mergeCells(headerRowNum - 1, mrPriceColIdx, headerRowNum - 1, mrSumColIdx);
    } catch {}
    const groupCell = sheet.getRow(headerRowNum - 1).getCell(mrPriceColIdx);
    groupCell.value     = 'Анализ MR Group';
    groupCell.font      = { bold: true, color: { argb: MR_FONT } };
    groupCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    fillCell(groupCell, MR_HEADER);
    borderCell(groupCell);
  }

  // 5b. Подзаголовки в строке заголовков
  const hRow = sheet.getRow(headerRowNum);

  const priceHdr  = hRow.getCell(mrPriceColIdx);
  priceHdr.value     = 'Цена за единицу материала 2025';
  priceHdr.font      = { bold: true, color: { argb: MR_FONT } };
  priceHdr.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  fillCell(priceHdr, MR_HEADER);
  borderCell(priceHdr);

  const sumHdr    = hRow.getCell(mrSumColIdx);
  sumHdr.value     = 'Сумма материалов 2025';
  sumHdr.font      = { bold: true, color: { argb: MR_FONT } };
  sumHdr.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  fillCell(sumHdr, MR_HEADER);
  borderCell(sumHdr);

  // Ширина новых столбцов
  sheet.getColumn(mrPriceColIdx).width = 22;
  sheet.getColumn(mrSumColIdx).width   = 20;

  // 6. Обход строк данных
  const groups = emptyGroups();
  let processed = 0, autoFilled = 0, toRequest = 0;

  sheet.eachRow({ includeEmpty: false }, (row, rowNum) => {
    if (rowNum <= headerRowNum) return;

    const wasPrice    = cellNum(row.getCell(wasCol.colNum));
    const becamePrice = cellNum(row.getCell(becameCol.colNum));

    // Пропустить строки без числовых цен в обоих столбцах
    if (!wasPrice || !becamePrice || wasPrice <= 0 || becamePrice <= 0) return;

    const ratio = becamePrice / wasPrice;
    const pct   = (ratio - 1) * 100;   // % удорожания
    const zone  = getZone(ratio);       // для окраски исходных столбцов

    processed++;

    const mrPriceCell = row.getCell(mrPriceColIdx);
    const mrSumCell   = row.getCell(mrSumColIdx);

    // ── «Сумма материалов 2025»: всегда формула = MRЦена × Кол-воОриентир ──
    if (sumQtyLetter) {
      mrSumCell.value = { formula: `${mrPriceLetter}${rowNum}*${sumQtyLetter}${rowNum}` };
    }
    // else: оставить пустым — нет столбца с количеством
    mrSumCell.numFmt = '#,##0.00 ₽';

    // ── «Цена за единицу материала 2025» (MR Group) ──
    if (pct <= 10) {
      // Зелёная зона: формула → ссылка на исходную «Цена стало»
      mrPriceCell.value  = { formula: `${becameLetter}${rowNum}` };
      mrPriceCell.numFmt = '#,##0.00 ₽';
      fillCell(mrPriceCell, MR_GREEN);
      autoFilled++;
    } else {
      // Оранжевая / Красная / Бордовая: ячейка ПУСТАЯ, менеджер заполняет вручную
      mrPriceCell.value  = null;
      mrPriceCell.numFmt = '#,##0.00 ₽';

      if (pct <= 40) {
        fillCell(mrPriceCell, MR_ORANGE); // рост 10–40%
      } else {
        fillCell(mrPriceCell, MR_RED);    // рост >40%
      }

      // Окрасить исходные ценовые столбцы (было / стало)
      fillCell(row.getCell(wasCol.colNum),    ZONE_COLORS[zone]);
      fillCell(row.getCell(becameCol.colNum), ZONE_COLORS[zone]);

      // Собрать строку для файла запроса
      const name           = nameCol  ? cellStr(row.getCell(nameCol))  : '';
      const classification = classCol ? cellStr(row.getCell(classCol)) : '';
      const unit           = unitCol  ? cellStr(row.getCell(unitCol))  : '';
      const qty            = qtyCol   ? cellNum(row.getCell(qtyCol))   : null;
      const section        = detectSection(classification, name);

      groups[section].push({ name, classification, unit, qty, wasPrice, becamePrice, ratio, zone });
      toRequest++;
    }
  });

  console.log(`  Обработано строк: ${processed}  |  Авто: ${autoFilled}  |  На запрос: ${toRequest}`);
  return { groups, processed, autoFilled, toRequest };
}

// ─── Создание файла запроса ──────────────────────────────────────────────────

async function writeRequestFile(allGroups, outputPath) {
  const reqWb = new ExcelJS.Workbook();

  // Сводный лист
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

  // Листы по разделам
  for (const sec of SECTIONS_ORDER) {
    const items = allGroups[sec];
    if (!items.length) continue;

    const s = reqWb.addWorksheet(sec);
    const hRow = s.addRow([
      'Наименование', 'Классификация', 'Ед. изм.', 'Кол-во',
      'Цена было', 'Цена стало', 'Коэф.', 'Зона',
    ]);
    hRow.font = { bold: true };
    hRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } };

    for (const item of items) {
      const dr = s.addRow([
        item.name, item.classification, item.unit, item.qty,
        item.wasPrice, item.becamePrice,
        parseFloat(item.ratio.toFixed(3)),
        ZONE_LABELS[item.zone],
      ]);
      dr.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: LIGHT_COLORS[item.zone] } };
    }

    s.columns.forEach(col => {
      let max = 10;
      col.eachCell({ includeEmpty: false }, cell => {
        const len = cell.value ? cell.value.toString().length : 0;
        if (len > max) max = len;
      });
      col.width = Math.min(max + 4, 60);
    });
  }

  await reqWb.xlsx.writeFile(outputPath);
}

// ─── Главная функция ─────────────────────────────────────────────────────────

async function analyze(inputPath) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(inputPath);

  if (!wb.worksheets.length) throw new Error('Файл не содержит листов.');

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
      console.warn(`  Пропуск листа "${sheet.name}": ${e.message}`);
    }
  }

  resolveSharedFormulas(wb);

  const dir     = path.dirname(inputPath);
  const base    = path.basename(inputPath, path.extname(inputPath));
  const outMain = path.join(dir, `${base}_analysis.xlsx`);
  const outReq  = path.join(dir, `${base}_request.xlsx`);

  await wb.xlsx.writeFile(outMain);
  await writeRequestFile(allGroups, outReq);

  console.log('\n=== Итого ===');
  console.log(`Обработано строк:       ${totalProcessed}`);
  console.log(`Заполнено авто (≤10%):  ${totalAuto}`);
  console.log(`На запрос уточнения:    ${totalRequest}`);
  console.log(`\nФайл анализа: ${outMain}`);
  console.log(`Файл запроса: ${outReq}`);

  return { analysisPath: outMain, requestPath: outReq };
}

module.exports = analyze;

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
