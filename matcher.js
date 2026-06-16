'use strict';

// ── Нормализация строки ───────────────────────────────────────────────────────

function normalize(str) {
  return str
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

// ── Итоговое сходство двух наименований ──────────────────────────────────────

function similarity(srcName, kpName) {
  if (!numbersMatch(srcName, kpName)) return 0;
  return trigramSim(normalize(srcName), normalize(kpName));
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
  const results  = [];
  const usedKp   = new Set();   // каждая позиция КП используется не более одного раза

  for (const src of sourceItems) {
    let bestScore = 0, bestKp = null, bestIdx = -1;

    kpItems.forEach((kp, idx) => {
      if (usedKp.has(idx)) return;
      const s = similarity(src.name, kp.name);
      if (s > bestScore) { bestScore = s; bestKp = kp; bestIdx = idx; }
    });

    if (bestKp && bestScore >= threshold) {
      usedKp.add(bestIdx);
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
