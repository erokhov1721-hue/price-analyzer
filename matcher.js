'use strict';

// ── Аббревиатуры материалов ───────────────────────────────────────────────────

const ABBR = {
  вгп:  'водогазопровод',
  эс:   'электросвар',
  пнд:  'полиэтилен',
  пвх:  'поливинилхлорид',
  ппр:  'полипропилен',
  нж:   'нержавеющ',
  гк:   'горячекатан',
  хк:   'холоднокатан',
  оц:   'оцинкован',
};

function expandAbbr(str) {
  return str.replace(/\b(вгп|эс|пнд|пвх|ппр|нж|гк|хк|оц)\b/gi,
    m => ABBR[m.toLowerCase()] || m);
}

// ── Нормализация строки ───────────────────────────────────────────────────────

function normalize(str) {
  return expandAbbr(str)
    .toLowerCase()
    // Символы диаметра → "ду"
    .replace(/[∅øØ]/g, 'ду')
    // Знаки умножения (включая кириллический «х») → латинский x
    .replace(/[×х✕]/g, 'x')
    // Удалить ссылки на ГОСТ, ТУ вместе с номером
    .replace(/\s*гост[\s\d\-.]+[\w-]*/gi, '')
    .replace(/\s*ту[\s\d\-.]+[\w-]*/gi, '')
    // Удалить стоп-слова
    .replace(/\b(по|из|для|и|в|на|с|от|купить|производство|тип|марка|класс|вид|серия|оцинковк[аи]|оцинкованн\w+)\b/gi, ' ')
    // Оставить буквы, цифры, x, точку, запятую, дефис
    .replace(/[^\wа-яёА-ЯЁ\d.x,\-]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

// ── Числовые параметры ────────────────────────────────────────────────────────

function extractNumbers(str) {
  return (str.match(/\d+(?:[.,]\d+)?/g) || [])
    .map(s => parseFloat(s.replace(',', '.')));
}

/**
 * Все числа из srcName должны присутствовать в kpName (с допуском 2 %).
 * Защищает от ложных совпадений по тексту при расхождении размеров.
 */
function numbersMatch(srcName, kpName) {
  const srcNums = extractNumbers(srcName);
  if (!srcNums.length) return true;               // нет цифр → не блокируем
  const kpNums = extractNumbers(kpName);
  return srcNums.every(n =>
    kpNums.some(m => Math.abs(n - m) / (Math.abs(n) || 1) < 0.02)
  );
}

// ── Триграммное сходство ──────────────────────────────────────────────────────

function trigrams(s) {
  const padded = ` ${s} `;
  const set = new Set();
  for (let i = 0; i < padded.length - 2; i++) set.add(padded.slice(i, i + 3));
  return set;
}

function trigramSim(a, b) {
  const ta = trigrams(a), tb = trigrams(b);
  if (!ta.size || !tb.size) return 0;
  let common = 0;
  for (const t of ta) if (tb.has(t)) common++;
  return (2 * common) / (ta.size + tb.size);
}

// ── Keyword-based matching ────────────────────────────────────────────────────

// Категории: если хотя бы один корень встретился в строке — категория определена
const CATEGORIES = [
  ['труб'],
  ['кабел', 'провод', 'кабель'],
  ['задвижк'],
  ['клапан'],
  ['кран'],
  ['арматур'],
  ['насос'],
  ['фланц', 'фланец'],
  ['муфт'],
  ['угол'],       // уголок
  ['швеллер'],
  ['двутавр'],
  ['профил'],
  ['лист'],
  ['прокат'],
  ['бетон'],
  ['цемент'],
  ['кирпич'],
  ['краск', 'грунт'],
  ['изолят'],
];

function getCategoryIdx(normStr) {
  for (let i = 0; i < CATEGORIES.length; i++) {
    if (CATEGORIES[i].some(stem => normStr.includes(stem))) return i;
  }
  return -1;
}

/**
 * Размерные параметры строки.
 * Приоритет: формат NxM (15x2.8, 100x5x3).
 * Fallback: одиночный диаметр после "ду" (ду15, ду 100).
 */
function extractDims(normStr) {
  const full = (normStr.match(/\d+(?:[.,]\d+)?(?:x\d+(?:[.,]\d+)?)+/g) || [])
    .map(d => d.replace(',', '.'));
  if (full.length) return full;

  const m = normStr.match(/\bду\s*(\d+(?:[.,]\d+)?)/);
  if (m) return [m[1].replace(',', '.')];

  return [];
}

function keywordSim(normSrc, normKp) {
  const ci = getCategoryIdx(normSrc);
  const cj = getCategoryIdx(normKp);
  if (ci === -1 || cj === -1 || ci !== cj) return 0;

  const dSrc = extractDims(normSrc);
  const dKp  = extractDims(normKp);

  if (dSrc.length && dKp.length) {
    return dSrc.some(d => dKp.includes(d)) ? 0.88 : 0.15;
  }
  // Одна из строк совсем без размеров — не делаем вывод по ключевым словам
  return 0;
}

// ── Итоговое сходство двух наименований ──────────────────────────────────────

function similarity(srcName, kpName) {
  if (!numbersMatch(srcName, kpName)) return 0;
  const nSrc = normalize(srcName);
  const nKp  = normalize(kpName);
  return Math.max(trigramSim(nSrc, nKp), keywordSim(nSrc, nKp));
}

// ── Основная функция сопоставления ───────────────────────────────────────────

/**
 * Находит наилучшие совпадения между позициями КП и позициями исходной сметы.
 *
 * @param {Array<{name: string, price: number}>}                         kpItems
 * @param {Array<{rowNum: number, name: string, sheetName: string}>}     sourceItems
 * @param {number} [threshold=0.62]  минимальное сходство (0..1)
 * @returns {Array<{sourceRowNum, sheetName, sourceName, kpName, kpPrice, score}>}
 */
function matchItems(kpItems, sourceItems, threshold = 0.62) {
  const results = [];

  for (const src of sourceItems) {
    let bestScore = 0, bestKp = null;

    kpItems.forEach(kp => {
      const s = similarity(src.name, kp.name);
      if (s > bestScore) { bestScore = s; bestKp = kp; }
    });

    if (bestKp && bestScore >= threshold) {
      results.push({
        sourceRowNum: src.rowNum,
        sheetName:    src.sheetName,
        sourceName:   src.name,
        kpName:       bestKp.name,
        kpPrice:      bestKp.price,
        score:        Math.round(bestScore * 100),
      });
    }
  }

  return results;
}

module.exports = { matchItems, similarity };
