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

/**
 * Исправляет случаи, когда pdf-parse склеивает количество и цену без пробела:
 * "80" (кол-во) + "1 096,80" (цена) → "801 096,80" (неверно распознано как 801 096,80).
 *
 * Алгоритм: если цена подозрительно велика по сравнению с медианой всех позиций
 * КП, пробуем отрезать 1–2 «лишние» цифры с начала первой тысячной группы.
 * Результат принимается только если:
 *  - оставшаяся часть начинается с ненулевой цифры (исключает "0 250,00")
 *  - новая цена меньше исходной минимум в 100 раз
 */
function fixConcatenatedPrices(items) {
  if (items.length < 2) return;

  const sorted = items.map(it => it.price).sort((a, b) => a - b);
  const median = sorted[Math.floor(sorted.length / 2)];

  for (const item of items) {
    if (item.price <= 100 * median || !item._raw) continue;

    const parts = item._raw.trim().split(/[\s\u00A0]+/);
    if (parts.length < 2) continue;

    const firstGroup = parts[0];
    if (!/^\d{2,3}$/.test(firstGroup)) continue;   // только 2–3-значный первый блок

    for (let n = 1; n < firstGroup.length; n++) {
      const stripped = firstGroup.slice(n);
      if (!stripped || stripped[0] === '0') continue;   // не должна начинаться с 0

      const altStr   = stripped + ' ' + parts.slice(1).join(' ');
      const altPrice = parseFloat(altStr.replace(/[\s\u00A0]/g, '').replace(',', '.'));

      if (altPrice > 0 && altPrice * 100 < item.price) {
        item.price = altPrice;
        break;
      }
    }
  }
}

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
      items.push({ name, price, _raw: m[1] });
    }
  }

  fixConcatenatedPrices(items);
  items.forEach(it => delete it._raw);

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
