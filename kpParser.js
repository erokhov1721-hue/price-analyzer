'use strict';

const ExcelJS = require('exceljs');
const fs      = require('fs');
const path    = require('path');

// ── Helpers ───────────────────────────────────────────────────────────────────

function parsePrice(raw) {
  if (raw === null || raw === undefined) return null;
  const s = raw.toString().replace(/[^\d.,]/g, '').replace(',', '.');
  const n = parseFloat(s);
  return (!isNaN(n) && n > 0) ? n : null;
}

// ── Excel-парсер ──────────────────────────────────────────────────────────────

async function parseExcelKP(filePath) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);
  const items = [];

  for (const sheet of wb.worksheets) {
    let nameCol = null, priceCol = null, headerRow = null;

    // Ищем строку заголовков (до 30-й строки)
    for (let r = 1; r <= 30; r++) {
      const row = sheet.getRow(r);
      let tmpName = null, tmpPrice = null;
      row.eachCell({ includeEmpty: false }, (cell, col) => {
        const v = cell.value ? cell.value.toString().toLowerCase().trim() : '';
        if (/^(наименование|материал|товар|позиция|наим\.|описание|номенклатура)/.test(v))
          tmpName = col;
        if (
          /^(цена\s*за\s*ед|стоимость\s*ед|цена\s*за\s*1|цена\s*ед|прайс|price)/.test(v) ||
          v === 'цена' || v === 'цена, руб' || v === 'цена (руб)' || v === 'цена, руб.'
        ) tmpPrice = col;
      });
      if (tmpName && tmpPrice) { nameCol = tmpName; priceCol = tmpPrice; headerRow = r; break; }
    }

    if (!nameCol || !priceCol) continue;

    sheet.eachRow({ includeEmpty: false }, (row, rowNum) => {
      if (rowNum <= headerRow) return;
      const raw   = row.getCell(nameCol).value;
      const name  = raw  ? raw.toString().trim()          : null;
      const price = parsePrice(row.getCell(priceCol).value);
      if (name && name.length > 4 && price) items.push({ name, price });
    });
  }

  return items;
}

// ── PDF-парсер ────────────────────────────────────────────────────────────────

// Число в русском формате: "1 234,56" | "1234.56" | "12345"
const RU_NUM_SRC = '\\d{1,3}(?:[\\s\\u00A0]\\d{3})*(?:[.,]\\d{1,2})?';
const PRICE_EOL  = new RegExp(`(${RU_NUM_SRC})\\s*(?:руб\\.?|₽|rub)?\\s*$`, 'i');

function extractItemsFromText(text) {
  const items = [];
  const lines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean);

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    // Пропустить итоговые строки
    if (/^(итого|всего|ндс|сумма\s*(всего|итого)|подытог|total)/i.test(line)) continue;

    const m = line.match(PRICE_EOL);
    if (!m) continue;

    const price = parseFloat(m[1].replace(/[\s\u00A0]/g, '').replace(',', '.'));
    if (price < 1 || price > 50_000_000) continue;   // граничные значения

    // Наименование — текст до цены в той же строке, или предыдущая строка
    const beforePrice = line.slice(0, line.lastIndexOf(m[0])).replace(/\s+\d+\s*$/, '').trim();
    const name = beforePrice.length > 5
      ? beforePrice
      : (i > 0 && lines[i - 1].length > 5 ? lines[i - 1] : null);

    if (
      name &&
      !/^(наименование|материал|ед\.?\s*изм|кол\.?|цена|сумма)/i.test(name)
    ) {
      items.push({ name, price });
    }
  }

  return items;
}

async function parsePDFKP(filePath) {
  // Ленивая загрузка — ошибка при отсутствии пакета видна только при реальном PDF
  const pdfParse = require('pdf-parse');
  const buffer   = fs.readFileSync(filePath);
  const data     = await pdfParse(buffer);
  return extractItemsFromText(data.text);
}

// ── Главный экспорт ───────────────────────────────────────────────────────────

/**
 * Парсит файл КП (PDF или Excel) и возвращает массив [{name, price}].
 * @param {string} filePath     - путь к временному файлу
 * @param {string} originalName - оригинальное имя файла (для определения расширения)
 */
async function parseKP(filePath, originalName) {
  const ext = path.extname(originalName).toLowerCase();
  if (ext === '.pdf')                    return parsePDFKP(filePath);
  if (ext === '.xlsx' || ext === '.xls') return parseExcelKP(filePath);
  throw new Error(`Неподдерживаемый формат: ${ext}. Используйте PDF, XLSX или XLS.`);
}

module.exports = parseKP;
