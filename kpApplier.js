'use strict';

const ExcelJS       = require('exceljs');
const { matchItems } = require('./matcher');

// ── Вспомогательные функции (идентичны analyze.js) ───────────────────────────

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

function cellStr(cell) { return cellText(cell).trim(); }

/** Ищет строку с заголовками цен — многопроходный поиск, строки 1–25 */
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
      if (vals.some(v => p.test(v)))
        return { rowNum: r, headers: vals };
    }
  }
  return null;
}

/** Возвращает номер первого столбца блока «Анализ MR Group» или null */
function findExistingMRBlock(sheet, headerRowNum) {
  // Ищем "Анализ MR Group" во всех строках от 1 до headerRowNum включительно.
  for (let r = 1; r <= headerRowNum; r++) {
    let found = null;
    sheet.getRow(r).eachCell({ includeEmpty: false }, (cell, col) => {
      if (!found && cellText(cell).includes('Анализ MR Group')) found = col;
    });
    if (found) return found;
  }
  return null;
}

/** Ищет первый столбец, заголовок которого содержит один из паттернов */
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

// ── Основная функция ──────────────────────────────────────────────────────────

const MR_GREEN = 'FFC6EFCE';

/**
 * Применяет цены из КП к файлу анализа, хранящемуся в Supabase.
 *
 * @param {string}                        analysisId  - ID анализа
 * @param {Array<{name:string,price:number}>} kpItems - позиции из КП
 * @param {object}                        db          - модуль lib/db.js
 * @returns {{ matched:number, total:number, details:Array }}
 */
async function applyKP(analysisId, kpItems, db) {
  // 1. Получить файлы анализа
  const files   = await db.getFilesByAnalysis(analysisId);
  const srcFile = files.find(f => f.file_type === 'source');
  const anaFile = files.find(f => f.file_type === 'analysis');

  if (!srcFile) throw new Error('Исходный файл (.xlsx) не найден для этого анализа');
  if (!anaFile) throw new Error('Файл анализа не найден — сначала запустите анализ');

  // 2. Читаем исходный xlsx → собираем позиции {rowNum, name, sheetName}
  const srcBuf = await db.downloadFile(srcFile.storage_path);
  const srcWb  = new ExcelJS.Workbook();
  await srcWb.xlsx.load(srcBuf);

  const sourceItems = [];
  for (const sheet of srcWb.worksheets) {
    const headerInfo = findHeaderRow(sheet);
    if (!headerInfo) continue;
    const { rowNum: headerRowNum } = headerInfo;

    const nameColIdx = findColInSheet(sheet,
      ['наименование', 'название материала', 'наим.', 'номенклатура']);
    if (!nameColIdx) continue;

    sheet.eachRow({ includeEmpty: false }, (row, rowNum) => {
      if (rowNum <= headerRowNum) return;
      const name = cellStr(row.getCell(nameColIdx));
      if (name && name.length > 3)
        sourceItems.push({ rowNum, name, sheetName: sheet.name });
    });
  }

  if (!sourceItems.length)
    throw new Error('В исходном файле не найдены строки с наименованиями материалов');

  // 3. Нечёткое сопоставление
  const matches = matchItems(kpItems, sourceItems);
  if (!matches.length) return { matched: 0, total: kpItems.length, details: [] };

  // 4. Обновляем файл анализа
  const anaBuf = await db.downloadFile(anaFile.storage_path);
  const anaWb  = new ExcelJS.Workbook();
  await anaWb.xlsx.load(anaBuf);

  let applied = 0;

  for (const sheet of anaWb.worksheets) {
    const headerInfo = findHeaderRow(sheet);
    if (!headerInfo) continue;
    const { rowNum: headerRowNum } = headerInfo;

    const mrPriceColIdx = findExistingMRBlock(sheet, headerRowNum);
    if (!mrPriceColIdx) continue;

    const mrSumColIdx   = mrPriceColIdx + 1;
    const mrPriceLetter = sheet.getColumn(mrPriceColIdx).letter;

    const sumQtyColIdx =
      findColInSheet(sheet, ['кол-во ориентир', 'кол.ориентир', 'ориентир']) ||
      findColInSheet(sheet, ['кол-во', 'количество', 'кол.']);
    const sumQtyLetter = sumQtyColIdx ? sheet.getColumn(sumQtyColIdx).letter : null;

    const sheetMatches = matches.filter(m => m.sheetName === sheet.name);

    for (const match of sheetMatches) {
      const row       = sheet.getRow(match.sourceRowNum);
      const priceCell = row.getCell(mrPriceColIdx);
      const sumCell   = row.getCell(mrSumColIdx);

      // Устанавливаем цену из КП → ячейка зелёная (как "авто")
      priceCell.value  = match.kpPrice;
      priceCell.numFmt = '#,##0.00 ₽';
      priceCell.fill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: MR_GREEN } };

      // Формула суммы
      if (sumQtyLetter) {
        sumCell.value  = { formula: `${mrPriceLetter}${match.sourceRowNum}*${sumQtyLetter}${match.sourceRowNum}` };
        sumCell.numFmt = '#,##0.00 ₽';
      }
      applied++;
    }
  }

  // 5. Перезаписываем файл анализа в Supabase
  const updatedBuf = Buffer.from(await anaWb.xlsx.writeBuffer());
  await db.uploadFile(
    anaFile.storage_path,
    updatedBuf,
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  );

  return {
    matched: applied,
    total:   kpItems.length,
    details: matches.slice(0, 30),   // первые 30 совпадений для отладки
  };
}

module.exports = applyKP;
