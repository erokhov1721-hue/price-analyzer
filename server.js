const express = require('express');
const multer  = require('multer');
const path    = require('path');
const fs      = require('fs');
const analyze = require('./analyze');

const app    = express();
const PORT   = 3000;
const UPLOAD = path.join(__dirname, 'uploads');

if (!fs.existsSync(UPLOAD)) fs.mkdirSync(UPLOAD);

const storage = multer.diskStorage({
  destination: UPLOAD,
  filename: (req, file, cb) => {
    cb(null, Buffer.from(file.originalname, 'latin1').toString('utf8'));
  },
});
const upload = multer({ storage, fileFilter: (req, file, cb) => {
  cb(null, path.extname(file.originalname).toLowerCase() === '.xlsx');
}});

// ─── HTML ─────────────────────────────────────────────────────────────────────

app.get('/', (req, res) => res.send(`<!DOCTYPE html>
<html lang="ru">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Анализ удорожания материалов</title>
  <style>
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

    body {
      font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
      background: #0f1117;
      min-height: 100vh;
      display: flex;
      flex-direction: column;
      align-items: center;
      padding: 48px 16px 80px;
      color: #e2e8f0;
    }

    /* ─── Header ─── */
    .header {
      text-align: center;
      margin-bottom: 40px;
    }
    .header .logo {
      display: inline-flex;
      align-items: center;
      gap: 10px;
      background: linear-gradient(135deg, #667eea, #764ba2);
      -webkit-background-clip: text;
      -webkit-text-fill-color: transparent;
      background-clip: text;
      font-size: 28px;
      font-weight: 700;
      letter-spacing: -0.5px;
      margin-bottom: 6px;
    }
    .header .logo svg { flex-shrink: 0; }
    .header p {
      color: #64748b;
      font-size: 14px;
      max-width: 420px;
      line-height: 1.6;
    }

    /* ─── Card ─── */
    .card {
      background: #1e2130;
      border: 1px solid #2d3248;
      border-radius: 16px;
      padding: 32px;
      width: 100%;
      max-width: 540px;
      box-shadow: 0 8px 40px rgba(0,0,0,.4);
    }

    /* ─── Legend ─── */
    .legend {
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
      margin-bottom: 24px;
    }
    .legend-item {
      display: flex;
      align-items: center;
      gap: 6px;
      font-size: 12px;
      color: #94a3b8;
      background: #262b3d;
      border-radius: 20px;
      padding: 4px 10px;
    }
    .dot { width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; }
    .dot-green   { background: #22c55e; }
    .dot-orange  { background: #f97316; }
    .dot-red     { background: #ef4444; }
    .dot-burgundy{ background: #9b1c4c; }

    /* ─── Drop zone ─── */
    .drop-zone {
      border: 2px dashed #3d4460;
      border-radius: 12px;
      padding: 44px 24px;
      text-align: center;
      cursor: pointer;
      transition: border-color .2s, background .2s;
      background: #161925;
      position: relative;
    }
    .drop-zone:hover { border-color: #667eea; background: #1a1e30; }
    .drop-zone.dragover { border-color: #667eea; background: #1c2040; }
    .drop-zone input[type=file] { display: none; }

    .drop-icon {
      width: 48px; height: 48px;
      background: linear-gradient(135deg, #667eea22, #764ba222);
      border: 1px solid #667eea44;
      border-radius: 12px;
      display: flex;
      align-items: center;
      justify-content: center;
      margin: 0 auto 16px;
    }
    .drop-icon svg { width: 24px; height: 24px; }

    .drop-zone h3 { font-size: 15px; color: #cbd5e1; margin-bottom: 6px; }
    .drop-zone p  { font-size: 13px; color: #475569; }
    .drop-zone p span { color: #818cf8; cursor: pointer; text-decoration: underline; }

    .file-selected {
      display: none;
      align-items: center;
      gap: 10px;
      background: #1f2d3d;
      border: 1px solid #2e4d6a;
      border-radius: 8px;
      padding: 10px 14px;
      margin-top: 14px;
      font-size: 13px;
      color: #93c5fd;
    }
    .file-selected svg { flex-shrink: 0; }
    .file-selected .fname { flex: 1; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .file-selected .remove { cursor: pointer; color: #475569; flex-shrink: 0; }
    .file-selected .remove:hover { color: #ef4444; }

    /* ─── Button ─── */
    .btn {
      margin-top: 20px;
      width: 100%;
      padding: 14px;
      background: linear-gradient(135deg, #667eea, #764ba2);
      color: #fff;
      border: none;
      border-radius: 10px;
      font-size: 15px;
      font-weight: 600;
      cursor: pointer;
      transition: opacity .2s, transform .1s;
      letter-spacing: .2px;
    }
    .btn:hover:not(:disabled) { opacity: .9; transform: translateY(-1px); }
    .btn:active:not(:disabled) { transform: translateY(0); }
    .btn:disabled { opacity: .35; cursor: not-allowed; }

    /* ─── Progress ─── */
    .progress-wrap { display: none; margin-top: 24px; }
    .progress-bar-bg {
      height: 4px; background: #2d3248; border-radius: 4px; overflow: hidden;
    }
    .progress-bar {
      height: 100%;
      background: linear-gradient(90deg, #667eea, #764ba2);
      border-radius: 4px;
      width: 0%;
      transition: width .4s ease;
    }
    .progress-label {
      margin-top: 10px;
      font-size: 13px;
      color: #64748b;
      text-align: center;
    }

    /* ─── Result ─── */
    .result { display: none; margin-top: 28px; }

    .result-header {
      display: flex;
      align-items: center;
      gap: 8px;
      margin-bottom: 20px;
      font-size: 15px;
      font-weight: 600;
      color: #22c55e;
    }

    .stats {
      display: grid;
      grid-template-columns: 1fr 1fr 1fr;
      gap: 10px;
      margin-bottom: 20px;
    }
    .stat-card {
      background: #161925;
      border: 1px solid #2d3248;
      border-radius: 10px;
      padding: 12px;
      text-align: center;
    }
    .stat-card .val { font-size: 22px; font-weight: 700; color: #e2e8f0; }
    .stat-card .lbl { font-size: 11px; color: #475569; margin-top: 4px; }

    .divider {
      height: 1px;
      background: #2d3248;
      margin: 20px 0;
    }

    .dl-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }

    .dl-card {
      display: flex;
      flex-direction: column;
      align-items: center;
      gap: 10px;
      background: #161925;
      border: 1px solid #2d3248;
      border-radius: 12px;
      padding: 20px 16px;
      text-decoration: none;
      transition: border-color .2s, background .2s, transform .1s;
      cursor: pointer;
    }
    .dl-card:hover { border-color: #667eea; background: #1a1e30; transform: translateY(-2px); }
    .dl-card:active { transform: translateY(0); }

    .dl-card .dl-icon {
      width: 44px; height: 44px;
      border-radius: 10px;
      display: flex; align-items: center; justify-content: center;
    }
    .dl-card.blue  .dl-icon { background: #1e3a5f; }
    .dl-card.green .dl-icon { background: #14352b; }
    .dl-card .dl-name { font-size: 13px; font-weight: 600; color: #cbd5e1; text-align: center; }
    .dl-card .dl-desc { font-size: 11px; color: #475569; text-align: center; }

    /* ─── Log ─── */
    .log-toggle {
      margin-top: 16px;
      font-size: 12px;
      color: #475569;
      cursor: pointer;
      display: flex;
      align-items: center;
      gap: 6px;
    }
    .log-toggle:hover { color: #94a3b8; }
    .log-box {
      display: none;
      margin-top: 10px;
      background: #0d1117;
      border: 1px solid #2d3248;
      border-radius: 8px;
      padding: 12px 14px;
      font-family: 'Cascadia Code', 'Consolas', monospace;
      font-size: 12px;
      color: #94a3b8;
      white-space: pre-wrap;
      max-height: 180px;
      overflow-y: auto;
      line-height: 1.6;
    }

    /* ─── Error ─── */
    .error-box {
      display: none;
      margin-top: 20px;
      background: #2d1515;
      border: 1px solid #7f1d1d;
      border-radius: 10px;
      padding: 14px 16px;
      font-size: 13px;
      color: #fca5a5;
      display: none;
    }
  </style>
</head>
<body>

<div class="header">
  <div class="logo">
    <svg width="28" height="28" viewBox="0 0 28 28" fill="none" style="-webkit-text-fill-color:initial">
      <rect width="28" height="28" rx="7" fill="url(#lg)"/>
      <path d="M7 20L11 13L15 17L19 9L21 12" stroke="white" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
      <defs><linearGradient id="lg" x1="0" y1="0" x2="28" y2="28"><stop stop-color="#667eea"/><stop offset="1" stop-color="#764ba2"/></linearGradient></defs>
    </svg>
    Анализ удорожания
  </div>
  <p>Автоматический анализ изменения цен на материалы по Excel-файлу</p>
</div>

<div class="card">

  <!-- Легенда зон -->
  <div class="legend">
    <div class="legend-item"><span class="dot dot-green"></span>≤1.1x — авто</div>
    <div class="legend-item"><span class="dot dot-orange"></span>1.1–1.5x — оранжевая</div>
    <div class="legend-item"><span class="dot dot-red"></span>1.5–2x — красная</div>
    <div class="legend-item"><span class="dot dot-burgundy"></span>>2x — бордовая</div>
  </div>

  <!-- Зона загрузки -->
  <div class="drop-zone" id="dropZone">
    <input type="file" id="fileInput" accept=".xlsx">
    <div class="drop-icon">
      <svg viewBox="0 0 24 24" fill="none" stroke="#818cf8" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
        <path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/>
        <polyline points="17 8 12 3 7 8"/>
        <line x1="12" y1="3" x2="12" y2="15"/>
      </svg>
    </div>
    <h3>Перетащите .xlsx файл сюда</h3>
    <p>или <span id="chooseLink">выберите файл</span> на компьютере</p>
  </div>

  <div class="file-selected" id="fileSelected">
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#93c5fd" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
      <path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/>
    </svg>
    <span class="fname" id="fName"></span>
    <span class="remove" id="removeFile" title="Удалить">
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
    </span>
  </div>

  <button class="btn" id="btn" disabled>Запустить анализ</button>

  <!-- Прогресс -->
  <div class="progress-wrap" id="progressWrap">
    <div class="progress-bar-bg"><div class="progress-bar" id="progressBar"></div></div>
    <div class="progress-label" id="progressLabel">Читаем файл…</div>
  </div>

  <!-- Результат -->
  <div class="result" id="result">
    <div class="result-header">
      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg>
      Анализ завершён
    </div>

    <div class="stats">
      <div class="stat-card">
        <div class="val" id="statTotal">—</div>
        <div class="lbl">строк обработано</div>
      </div>
      <div class="stat-card">
        <div class="val" id="statAuto" style="color:#22c55e">—</div>
        <div class="lbl">заполнено авто</div>
      </div>
      <div class="stat-card">
        <div class="val" id="statReq" style="color:#f97316">—</div>
        <div class="lbl">на запрос</div>
      </div>
    </div>

    <div class="divider"></div>

    <div class="dl-grid">
      <a class="dl-card blue" id="dlAnalysis" href="#" download>
        <div class="dl-icon">
          <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#60a5fa" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
            <rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M9 21V9"/>
          </svg>
        </div>
        <span class="dl-name">Файл анализа</span>
        <span class="dl-desc">С цветами и Анализ MR Group</span>
      </a>
      <a class="dl-card green" id="dlRequest" href="#" download>
        <div class="dl-icon">
          <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#4ade80" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
            <path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/><polyline points="10 9 9 9 8 9"/>
          </svg>
        </div>
        <span class="dl-name">Файл запроса</span>
        <span class="dl-desc">По разделам для уточнения</span>
      </a>
    </div>

    <div class="log-toggle" id="logToggle">
      <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"/></svg>
      Показать лог
    </div>
    <div class="log-box" id="logBox"></div>
  </div>

  <!-- Ошибка -->
  <div class="error-box" id="errorBox"></div>

</div>

<script>
  const dropZone   = document.getElementById('dropZone');
  const fileInput  = document.getElementById('fileInput');
  const fileSelected = document.getElementById('fileSelected');
  const fName      = document.getElementById('fName');
  const removeFile = document.getElementById('removeFile');
  const btn        = document.getElementById('btn');
  const chooseLink = document.getElementById('chooseLink');

  let selectedFile = null;

  chooseLink.onclick = (e) => { e.stopPropagation(); fileInput.click(); };
  dropZone.onclick   = () => fileInput.click();

  dropZone.addEventListener('dragover',  e => { e.preventDefault(); dropZone.classList.add('dragover'); });
  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
  dropZone.addEventListener('drop', e => {
    e.preventDefault(); dropZone.classList.remove('dragover');
    const f = e.dataTransfer.files[0];
    if (f && f.name.endsWith('.xlsx')) setFile(f);
  });

  fileInput.onchange = () => { if (fileInput.files[0]) setFile(fileInput.files[0]); };

  removeFile.onclick = (e) => {
    e.stopPropagation();
    selectedFile = null;
    fileInput.value = '';
    fileSelected.style.display = 'none';
    btn.disabled = true;
  };

  function setFile(f) {
    selectedFile = f;
    fName.textContent = f.name;
    fileSelected.style.display = 'flex';
    btn.disabled = false;
    document.getElementById('result').style.display = 'none';
    document.getElementById('errorBox').style.display = 'none';
  }

  // Анимация прогресс-бара (имитация, т.к. сервер не стримит прогресс)
  let progressTimer = null;
  function startProgress() {
    const bar    = document.getElementById('progressBar');
    const label  = document.getElementById('progressLabel');
    const stages = [
      [0,  15, 800,  'Читаем файл…'],
      [15, 40, 1200, 'Определяем столбцы…'],
      [40, 75, 2000, 'Анализируем строки…'],
      [75, 90, 1500, 'Формируем файлы…'],
    ];
    let i = 0;
    bar.style.width = '0%';
    const run = () => {
      if (i >= stages.length) return;
      const [from, to, dur, text] = stages[i++];
      label.textContent = text;
      bar.style.transition = 'width ' + dur + 'ms ease';
      bar.style.width = to + '%';
      progressTimer = setTimeout(run, dur + 100);
    };
    run();
  }

  function stopProgress() {
    clearTimeout(progressTimer);
    const bar = document.getElementById('progressBar');
    bar.style.transition = 'width .3s ease';
    bar.style.width = '100%';
  }

  btn.onclick = async () => {
    if (!selectedFile) return;

    btn.disabled = true;
    document.getElementById('result').style.display     = 'none';
    document.getElementById('errorBox').style.display   = 'none';
    document.getElementById('progressWrap').style.display = 'block';
    startProgress();

    const fd = new FormData();
    fd.append('file', selectedFile);

    try {
      const res  = await fetch('/analyze', { method: 'POST', body: fd });
      const data = await res.json();
      stopProgress();
      setTimeout(() => {
        document.getElementById('progressWrap').style.display = 'none';
        if (!res.ok) throw new Error(data.error || 'Неизвестная ошибка');

        // Статистика
        const m = data.log.match(/Обработано строк:\\s+(\\d+)[\\s\\S]*?авто[^:]+:\\s+(\\d+)[\\s\\S]*?уточнения:\\s+(\\d+)/);
        if (m) {
          document.getElementById('statTotal').textContent = m[1];
          document.getElementById('statAuto').textContent  = m[2];
          document.getElementById('statReq').textContent   = m[3];
        }

        document.getElementById('dlAnalysis').href     = '/download/' + encodeURIComponent(data.analysisFile);
        document.getElementById('dlAnalysis').download = data.analysisFile;
        document.getElementById('dlRequest').href      = '/download/' + encodeURIComponent(data.requestFile);
        document.getElementById('dlRequest').download  = data.requestFile;
        document.getElementById('logBox').textContent  = data.log;
        document.getElementById('result').style.display = 'block';
        btn.disabled = false;
      }, 400);
    } catch (err) {
      stopProgress();
      document.getElementById('progressWrap').style.display = 'none';
      const box = document.getElementById('errorBox');
      box.style.display = 'block';
      box.textContent   = '❌ ' + err.message;
      btn.disabled = false;
    }
  };

  // Лог-тоггл
  document.getElementById('logToggle').onclick = () => {
    const box     = document.getElementById('logBox');
    const toggle  = document.getElementById('logToggle');
    const open    = box.style.display === 'block';
    box.style.display   = open ? 'none' : 'block';
    toggle.innerHTML    = open
      ? '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"/></svg> Показать лог'
      : '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="18 15 12 9 6 15"/></svg> Скрыть лог';
  };
</script>
</body>
</html>`));

// ─── API ──────────────────────────────────────────────────────────────────────

app.post('/analyze', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Файл не передан или не является .xlsx' });

  const logs = [];
  const origLog  = console.log;
  const origWarn = console.warn;
  console.log  = (...a) => { origLog(...a);  logs.push(a.join(' ')); };
  console.warn = (...a) => { origWarn(...a); logs.push(a.join(' ')); };

  try {
    const { analysisPath, requestPath } = await analyze(req.file.path);
    console.log  = origLog;
    console.warn = origWarn;

    // Переместить результаты в папку uploads
    const aName = path.basename(analysisPath);
    const rName = path.basename(requestPath);
    const aDest = path.join(UPLOAD, aName);
    const rDest = path.join(UPLOAD, rName);
    if (analysisPath !== aDest) fs.renameSync(analysisPath, aDest);
    if (requestPath  !== rDest) fs.renameSync(requestPath,  rDest);

    res.json({ analysisFile: aName, requestFile: rName, log: logs.join('\n') });
  } catch (err) {
    console.log  = origLog;
    console.warn = origWarn;
    res.status(500).json({ error: err.message });
  }
});

app.get('/download/:filename', (req, res) => {
  const file = path.join(UPLOAD, req.params.filename);
  if (!fs.existsSync(file)) return res.status(404).send('Файл не найден');
  res.download(file);
});

// ─── Старт ───────────────────────────────────────────────────────────────────

app.listen(PORT, () => console.log(`Сервер: http://localhost:${PORT}`));
