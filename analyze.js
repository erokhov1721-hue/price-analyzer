const ExcelJS = require('exceljs');
const path = require('path');

// ─── Константы ───────────────────────────────────────────────────────────────

function getZone(ratio) {
  if (ratio <= 1.1)  return 'auto';
  if (ratio < 1.5)   return 'orange';
  if (ratio < 2.0)   return 'red';
  return 'burgundy';
}

const ZONE_COLORS  = { orange: 'FFFFC000', red: 'FFFF0000', burgundy: 'FF800020' };
const ZONE_LABELS  = {
  auto:    'Авто (≤1.1x)',
  orange:  'Оранжевая (1.1–1.5x)',
  red:     'Красная (1.5–2x)',
  burgundy:'Бордовая (>2x)',
};
const LIGHT_COLORS = { orange: 'FFFFF2CC', red: 'FFFFD7D7', burgundy: 'FFEDD5D5' };

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
  if (typeof v === 'object' && v.result !== undefined) return v.result != null ? String(v.result).trim() : '';
  return String(v).trim();
}

function fillCell(cell, argb) {
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb } };
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

/** Ищет столбец по паттернам в первых maxRow строках */
function findColInSheet(sheet, patterns, maxRow = 10, excludeCol = null) {
  for (let r = 1; r <= maxRow; r++) {
    const row = sheet.getRow(r);
    let found = null;
    row.eachCell({ includeEmpty: true }, (cell, col) => {
      if (found || col === excludeCol || !cell.value) return;
      const str = cell.value.toString().toLowerCase();
      if (patterns.some(p => str.includes(p))) found = col;
    });
    if (found) return found;
  }
  return null;
}

/** Найти "Анализ MR Group" — ищем по всем первым 10 строкам */
function findMRCol(sheet) {
  return findColInSheet(sheet, ['анализ mr group', 'анализ mr']);
}

/** Найти строку заголовков (содержит "цена за единицу материала YEAR") */
function findHeaderRow(sheet) {
  for (let r = 1; r <= 10; r++) {
    const row = sheet.getRow(r);
    const vals = [];
    row.eachCell({ includeEmpty: true }, (cell, col) => { vals[col - 1] = cell.value ? cell.value.toString() : ''; });
    if (vals.some(v => /цена за единицу материала\s+\d{4}/i.test(v))) {
      return { rowNum: r, headers: vals };
    }
  }
  return null;
}

/** Из массива заголовков выбрать ценовые столбцы (исключая mrCol) */
function detectPriceCols(headers, mrCol) {
  const pattern = /цена за единицу материала\s+(\d{4})/i;
  const found = [];
  headers.forEach((h, i) => {
    const colNum = i + 1;
    if (colNum === mrCol || !h) return;
    const m = h.toString().match(pattern);
    if (m) found.push({ colNum, year: parseInt(m[1]), header: h.toString() });
  });
  found.sort((a, b) => a.year - b.year);
  return found;
}

// ─── Исправление shared formulas перед записью ───────────────────────────────

function resolveSharedFormulas(workbook) {
  workbook.worksheets.forEach(sheet => {
    sheet.eachRow({ includeEmpty: false }, row => {
      row.eachCell({ includeEmpty: false }, cell => {
        const v = cell.value;
        if (v && typeof v === 'object' && v !== null) {
          // Клоны shared formula → заменить числом
          if (v.sharedFormula !== undefined) {
            cell.value = v.result !== undefined ? v.result : null;
          }
          // Мастер shared formula → сделать обычной формулой
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
  const headerInfo = findHeaderRow(sheet);
  if (!headerInfo) throw new Error(`не найдена строка с заголовками цен`);

  const { rowNum: headerRowNum, headers } = headerInfo;
  const mrCol     = findMRCol(sheet);
  const priceCols = detectPriceCols(headers, mrCol);

  if (priceCols.length < 2) throw new Error(`найдено менее двух ценовых столбцов`);

  const wasCol    = priceCols[0];
  const becameCol = priceCols[priceCols.length - 1];

  // Ищем столбцы наименования, классификации, ед.изм., кол-ва по всему листу
  const nameCol  = findColInSheet(sheet, ['наименование', 'название материала']);
  const classCol = findColInSheet(sheet, ['классификация']);
  const unitCol  = findColInSheet(sheet, ['ед. изм', 'ед.изм', 'единица']);
  const qtyCol   = findColInSheet(sheet, ['кол-во', 'количество', 'кол.']);

  console.log(`\nЛист: "${sheet.name}"`);
  console.log(`  Было: кол.${wasCol.colNum} [${wasCol.year}]  Стало: кол.${becameCol.colNum} [${becameCol.year}]  MR Group: кол.${mrCol || '—'}`);

  const groups = emptyGroups();
  let processed = 0, autoFilled = 0, toRequest = 0;

  sheet.eachRow({ includeEmpty: false }, (row, rowNum) => {
    if (rowNum <= headerRowNum) return;

    const wasPrice    = cellNum(row.getCell(wasCol.colNum));
    const becamePrice = cellNum(row.getCell(becameCol.colNum));
    if (!wasPrice || !becamePrice || wasPrice <= 0 || becamePrice <= 0) return;

    const ratio = becamePrice / wasPrice;
    const zone  = getZone(ratio);

    processed++;

    if (zone === 'auto') {
      if (mrCol) row.getCell(mrCol).value = becamePrice;
      autoFilled++;
    } else {
      fillCell(row.getCell(wasCol.colNum),    ZONE_COLORS[zone]);
      fillCell(row.getCell(becameCol.colNum), ZONE_COLORS[zone]);

      const name           = nameCol  ? cellStr(row.getCell(nameCol))  : '';
      const classification = classCol ? cellStr(row.getCell(classCol)) : '';
      const unit           = unitCol  ? cellStr(row.getCell(unitCol))  : '';
      const qty            = qtyCol   ? cellNum(row.getCell(qtyCol))   : null;
      const section        = detectSection(classification, name);

      groups[section].push({ name, classification, unit, qty, wasPrice, becamePrice, ratio, zone });
      toRequest++;
    }
  });

  console.log(`  Обработано: ${processed}  |  Авто: ${autoFilled}  |  На запрос: ${toRequest}`);
  return { groups, processed, autoFilled, toRequest };
}

// ─── Создание файла запроса ───────────────────────────────────────────────────

async function writeRequestFile(allGroups, outputPath) {
  const reqWb = new ExcelJS.Workbook();

  // Сводка
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
    const hRow = s.addRow(['Наименование', 'Классификация', 'Ед. изм.', 'Кол-во', 'Цена было', 'Цена стало', 'Коэф.', 'Зона']);
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

  // Исправить shared formulas перед сохранением
  resolveSharedFormulas(wb);

  const dir     = path.dirname(inputPath);
  const base    = path.basename(inputPath, path.extname(inputPath));
  const outMain = path.join(dir, `${base}_analysis.xlsx`);
  const outReq  = path.join(dir, `${base}_request.xlsx`);

  await wb.xlsx.writeFile(outMain);
  await writeRequestFile(allGroups, outReq);

  console.log('\n=== Итого ===');
  console.log(`Обработано строк:       ${totalProcessed}`);
  console.log(`Заполнено авто (≤1.1x): ${totalAuto}`);
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

