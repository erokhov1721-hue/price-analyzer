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

const HTML = `<!DOCTYPE html>
<html lang="ru">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Анализ удорожания материалов</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=DM+Sans:ital,opsz,wght@0,9..40,400;0,9..40,500;0,9..40,600;0,9..40,700;1,9..40,400&display=swap">
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.css">
  <style>
    /* ── CSS переменные (Portal design system) ── */
    :root {
      --font-ui: 'DM Sans', system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
      --bg-page:      #F5F5F7;
      --bg-card:      #FFFFFF;
      --bg-secondary: #F9FAFB;
      --bg-tertiary:  #F3F4F6;
      --gray-200: #E5E5E7;
      --gray-300: #D1D1D6;
      --gray-400: #A1A1A6;
      --gray-500: #8E8E93;
      --gray-600: #636366;
      --gray-700: #48484A;
      --gray-900: #1D1D1F;
      --accent-blue:         #007AFF;
      --accent-blue-light:   rgba(0,122,255,0.08);
      --accent-green:        #34C759;
      --accent-green-light:  rgba(52,199,89,0.08);
      --accent-orange:       #FB8C00;
      --accent-orange-light: rgba(251,140,0,0.08);
      --accent-red:          #E53935;
      --accent-red-light:    rgba(229,57,53,0.08);
      --accent-burgundy:     #9B1C4C;
      --accent-burgundy-light: rgba(155,28,76,0.08);
      --shadow-sm: 0 2px 8px rgba(0,0,0,0.06);
      --shadow-md: 0 4px 16px rgba(0,0,0,0.08);
      --shadow-lg: 0 8px 32px rgba(0,0,0,0.10);
      --radius-sm: 8px;
      --radius-md: 12px;
      --radius-lg: 16px;
      --transition: 0.18s ease;
    }

    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

    body {
      font-family: var(--font-ui);
      font-size: 14px;
      line-height: 1.5;
      background: var(--bg-page);
      color: var(--gray-900);
      min-height: 100vh;
      -webkit-font-smoothing: antialiased;
    }

    /* ── Layout ── */
    .app-layout {
      display: flex;
      min-height: 100vh;
    }

    /* ── Sidebar (тёмный, как в Portal) ── */
    .sidebar {
      width: 280px;
      flex-shrink: 0;
      background: linear-gradient(180deg, #1C1C1E 0%, #2C2C2E 100%);
      display: flex;
      flex-direction: column;
      padding: 32px 20px 24px;
    }

    .sidebar-logo {
      display: flex;
      align-items: center;
      gap: 10px;
      margin-bottom: 32px;
    }
    .sidebar-logo .logo-icon {
      width: 36px; height: 36px;
      background: var(--accent-blue);
      border-radius: 10px;
      display: flex; align-items: center; justify-content: center;
      flex-shrink: 0;
    }
    .sidebar-logo .logo-icon i { font-size: 18px; color: #fff; }
    .sidebar-logo .logo-text h1 {
      font-size: 15px; font-weight: 700; color: #fff; letter-spacing: -0.2px;
    }
    .sidebar-logo .logo-text p {
      font-size: 12px; color: rgba(255,255,255,0.4); margin-top: 1px;
    }

    .sidebar-section-title {
      font-size: 11px;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 0.6px;
      color: rgba(255,255,255,0.35);
      margin-bottom: 10px;
      padding-left: 4px;
    }

    /* Зоны в сайдбаре */
    .zone-list { display: flex; flex-direction: column; gap: 6px; margin-bottom: 32px; }
    .zone-item {
      display: flex; align-items: center; gap: 10px;
      background: rgba(255,255,255,0.06);
      border-radius: var(--radius-sm);
      padding: 10px 12px;
    }
    .zone-dot {
      width: 10px; height: 10px; border-radius: 50%; flex-shrink: 0;
    }
    .zone-item .zone-info .zone-name {
      font-size: 13px; font-weight: 500; color: rgba(255,255,255,0.85);
    }
    .zone-item .zone-info .zone-range {
      font-size: 11px; color: rgba(255,255,255,0.4); margin-top: 1px;
    }
    .zone-auto   .zone-dot { background: var(--accent-green); }
    .zone-orange .zone-dot { background: var(--accent-orange); }
    .zone-red    .zone-dot { background: var(--accent-red); }
    .zone-burg   .zone-dot { background: var(--accent-burgundy); }

    .sidebar-footer {
      margin-top: auto;
      padding-top: 20px;
      border-top: 1px solid rgba(255,255,255,0.08);
      font-size: 12px;
      color: rgba(255,255,255,0.25);
    }

    /* ── Main content ── */
    .main-content {
      flex: 1;
      overflow-y: auto;
      padding: 28px;
      display: flex;
      flex-direction: column;
      align-items: center;
    }

    /* ── Page header ── */
    .page-header {
      width: 100%;
      max-width: 680px;
      margin-bottom: 20px;
    }
    .page-header h2 {
      font-size: 22px; font-weight: 700;
      color: var(--gray-900); letter-spacing: -0.3px;
    }
    .page-header p {
      margin-top: 4px; font-size: 14px; color: var(--gray-500);
    }

    /* ── Card ── */
    .card {
      background: var(--bg-card);
      border-radius: var(--radius-lg);
      box-shadow: var(--shadow-md);
      width: 100%;
      max-width: 680px;
      overflow: hidden;
      animation: fadeUp 0.3s ease;
    }

    @keyframes fadeUp {
      from { opacity: 0; transform: translateY(10px); }
      to   { opacity: 1; transform: translateY(0); }
    }

    .card-section {
      padding: 24px 28px;
      border-bottom: 1px solid var(--gray-200);
    }
    .card-section:last-child { border-bottom: none; }

    .section-label {
      font-size: 11px; font-weight: 700; text-transform: uppercase;
      letter-spacing: 0.6px; color: var(--gray-500); margin-bottom: 14px;
    }

    /* ── Drop zone ── */
    .drop-zone {
      border: 1.5px dashed var(--gray-300);
      border-radius: var(--radius-md);
      padding: 40px 24px;
      text-align: center;
      cursor: pointer;
      background: var(--bg-secondary);
      transition: var(--transition);
    }
    .drop-zone:hover  { border-color: var(--accent-blue); background: var(--accent-blue-light); }
    .drop-zone.drag   { border-color: var(--accent-blue); background: var(--accent-blue-light); }
    .drop-zone input  { display: none; }

    .drop-icon {
      width: 52px; height: 52px;
      background: var(--accent-blue-light);
      border-radius: var(--radius-md);
      display: flex; align-items: center; justify-content: center;
      margin: 0 auto 14px;
      border: 1px solid rgba(0,122,255,0.15);
    }
    .drop-icon i { font-size: 22px; color: var(--accent-blue); }

    .drop-zone h3 { font-size: 15px; font-weight: 600; color: var(--gray-900); margin-bottom: 5px; }
    .drop-zone p  { font-size: 13px; color: var(--gray-500); }
    .drop-zone p a { color: var(--accent-blue); cursor: pointer; text-decoration: none; font-weight: 500; }
    .drop-zone p a:hover { text-decoration: underline; }

    /* Выбранный файл */
    .file-pill {
      display: none;
      align-items: center;
      gap: 10px;
      margin-top: 12px;
      padding: 10px 14px;
      background: var(--accent-blue-light);
      border: 1px solid rgba(0,122,255,0.2);
      border-radius: var(--radius-sm);
    }
    .file-pill i { color: var(--accent-blue); font-size: 16px; flex-shrink: 0; }
    .file-pill .fname {
      flex: 1; font-size: 13px; font-weight: 500; color: var(--accent-blue);
      white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
    }
    .file-pill .remove {
      cursor: pointer; color: var(--gray-400); flex-shrink: 0;
      width: 20px; height: 20px; display: flex; align-items: center; justify-content: center;
      border-radius: 50%; transition: var(--transition);
    }
    .file-pill .remove:hover { background: var(--accent-red-light); color: var(--accent-red); }

    /* ── Button ── */
    .btn-primary {
      width: 100%; padding: 13px;
      background: var(--accent-blue);
      color: #fff; border: none;
      border-radius: var(--radius-sm);
      font-family: var(--font-ui);
      font-size: 15px; font-weight: 600;
      cursor: pointer;
      transition: var(--transition);
      letter-spacing: -0.1px;
    }
    .btn-primary:hover:not(:disabled) { background: #0071E3; }
    .btn-primary:active:not(:disabled) { transform: scale(0.99); }
    .btn-primary:disabled { opacity: 0.35; cursor: not-allowed; }

    /* ── Progress ── */
    .progress-section { display: none; padding: 20px 28px; border-bottom: 1px solid var(--gray-200); }
    .progress-header {
      display: flex; justify-content: space-between; align-items: center;
      margin-bottom: 10px;
    }
    .progress-header span { font-size: 13px; font-weight: 500; color: var(--gray-700); }
    .progress-pct { font-size: 13px; color: var(--accent-blue); font-weight: 600; }
    .progress-track {
      height: 4px; background: var(--bg-tertiary);
      border-radius: 4px; overflow: hidden;
    }
    .progress-fill {
      height: 100%; background: var(--accent-blue);
      border-radius: 4px; width: 0%; transition: width 0.5s ease;
    }

    /* ── Result ── */
    .result-section { display: none; }

    /* Stats row */
    .stats-row {
      display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 12px;
      padding: 20px 28px; border-bottom: 1px solid var(--gray-200);
    }
    .stat-box {
      background: var(--bg-secondary);
      border: 1px solid var(--gray-200);
      border-radius: var(--radius-md); padding: 14px 16px; text-align: center;
    }
    .stat-box .num {
      font-size: 24px; font-weight: 700; letter-spacing: -0.5px; color: var(--gray-900);
      line-height: 1;
    }
    .stat-box .num.green { color: var(--accent-green); }
    .stat-box .num.orange { color: var(--accent-orange); }
    .stat-box .lbl { font-size: 11px; color: var(--gray-500); margin-top: 5px; }

    /* Download cards */
    .dl-section { padding: 20px 28px; border-bottom: 1px solid var(--gray-200); }
    .dl-title { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.6px; color: var(--gray-500); margin-bottom: 12px; }
    .dl-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }

    .dl-card {
      display: flex; align-items: center; gap: 14px;
      padding: 14px 16px;
      background: var(--bg-secondary);
      border: 1px solid var(--gray-200);
      border-radius: var(--radius-md);
      text-decoration: none; color: inherit;
      transition: var(--transition); cursor: pointer;
    }
    .dl-card:hover { border-color: var(--accent-blue); background: var(--accent-blue-light); transform: translateY(-1px); box-shadow: var(--shadow-sm); }

    .dl-card .dl-ico {
      width: 40px; height: 40px; border-radius: var(--radius-sm);
      display: flex; align-items: center; justify-content: center; flex-shrink: 0;
    }
    .dl-card.blue  .dl-ico { background: rgba(0,122,255,0.1); }
    .dl-card.green .dl-ico { background: rgba(52,199,89,0.1); }
    .dl-card.blue  .dl-ico i { color: var(--accent-blue); font-size: 18px; }
    .dl-card.green .dl-ico i { color: var(--accent-green); font-size: 18px; }
    .dl-card .dl-info .dl-name { font-size: 13px; font-weight: 600; color: var(--gray-900); }
    .dl-card .dl-info .dl-desc { font-size: 12px; color: var(--gray-500); margin-top: 2px; }

    /* Log */
    .log-section { padding: 0 28px 20px; }
    .log-toggle {
      display: flex; align-items: center; gap: 6px;
      font-size: 12px; color: var(--gray-500); cursor: pointer;
      padding: 10px 0; user-select: none; border: none; background: none;
      font-family: var(--font-ui); transition: color var(--transition);
    }
    .log-toggle:hover { color: var(--gray-700); }
    .log-toggle i { font-size: 13px; transition: transform var(--transition); }
    .log-toggle.open i { transform: rotate(180deg); }
    .log-body {
      display: none; background: var(--bg-page);
      border: 1px solid var(--gray-200); border-radius: var(--radius-sm);
      padding: 12px 14px;
      font-family: 'Cascadia Code', 'Consolas', monospace;
      font-size: 12px; color: var(--gray-600); white-space: pre-wrap;
      max-height: 160px; overflow-y: auto; line-height: 1.6;
    }

    /* Error */
    .error-box {
      display: none; margin: 0 28px 20px;
      background: var(--accent-red-light);
      border: 1px solid rgba(229,57,53,0.2);
      border-radius: var(--radius-sm);
      padding: 12px 16px;
      font-size: 13px; color: var(--accent-red);
      display: none;
    }

    /* ── Responsive ── */
    @media (max-width: 760px) {
      .app-layout { flex-direction: column; }
      .sidebar { width: 100%; flex-direction: row; flex-wrap: wrap; gap: 12px; padding: 16px; }
      .sidebar-logo { margin-bottom: 0; }
      .zone-list { flex-direction: row; flex-wrap: wrap; margin-bottom: 0; }
      .sidebar-footer, .sidebar-section-title { display: none; }
      .main-content { padding: 16px; }
      .stats-row { grid-template-columns: 1fr 1fr 1fr; }
      .dl-grid { grid-template-columns: 1fr; }
    }
  </style>
</head>
<body>
<div class="app-layout">

  <!-- ── Sidebar ── -->
  <aside class="sidebar">
    <div class="sidebar-logo">
      <div class="logo-icon"><i class="bi bi-graph-up-arrow"></i></div>
      <div class="logo-text">
        <h1>Price Analyzer</h1>
        <p>МR Group</p>
      </div>
    </div>

    <div class="sidebar-section-title">Зоны удорожания</div>
    <div class="zone-list">
      <div class="zone-item zone-auto">
        <div class="zone-dot"></div>
        <div class="zone-info">
          <div class="zone-name">Авто-заполнение</div>
          <div class="zone-range">≤ 1.1× — заполняется</div>
        </div>
      </div>
      <div class="zone-item zone-orange">
        <div class="zone-dot"></div>
        <div class="zone-info">
          <div class="zone-name">Оранжевая зона</div>
          <div class="zone-range">1.1× – 1.5× — на запрос</div>
        </div>
      </div>
      <div class="zone-item zone-red">
        <div class="zone-dot"></div>
        <div class="zone-info">
          <div class="zone-name">Красная зона</div>
          <div class="zone-range">1.5× – 2.0× — на запрос</div>
        </div>
      </div>
      <div class="zone-item zone-burg">
        <div class="zone-dot"></div>
        <div class="zone-info">
          <div class="zone-name">Бордовая зона</div>
          <div class="zone-range">> 2.0× — на запрос</div>
        </div>
      </div>
    </div>

    <div class="sidebar-footer">
      Автоматический анализ<br>изменения цен на материалы
    </div>
  </aside>

  <!-- ── Main ── -->
  <main class="main-content">

    <div class="page-header">
      <h2>Анализ удорожания материалов</h2>
      <p>Загрузите Excel-файл — программа автоматически определит зоны, закрасит ячейки и сформирует запрос по разделам.</p>
    </div>

    <div class="card">

      <!-- Загрузка файла -->
      <div class="card-section">
        <div class="section-label">Входной файл</div>
        <div class="drop-zone" id="dropZone">
          <input type="file" id="fileInput" accept=".xlsx">
          <div class="drop-icon"><i class="bi bi-cloud-upload"></i></div>
          <h3>Перетащите .xlsx файл сюда</h3>
          <p>или <a id="chooseLink">выберите файл</a> на компьютере</p>
        </div>
        <div class="file-pill" id="filePill">
          <i class="bi bi-file-earmark-spreadsheet"></i>
          <span class="fname" id="fName"></span>
          <span class="remove" id="removeFile"><i class="bi bi-x"></i></span>
        </div>
      </div>

      <!-- Кнопка -->
      <div class="card-section">
        <button class="btn-primary" id="btn" disabled>
          <i class="bi bi-play-fill"></i>&nbsp; Запустить анализ
        </button>
      </div>

      <!-- Прогресс -->
      <div class="progress-section" id="progressSection">
        <div class="progress-header">
          <span id="progressLabel">Читаем файл…</span>
          <span class="progress-pct" id="progressPct">0%</span>
        </div>
        <div class="progress-track">
          <div class="progress-fill" id="progressFill"></div>
        </div>
      </div>

      <!-- Результат -->
      <div class="result-section" id="resultSection">

        <!-- Статистика -->
        <div class="stats-row">
          <div class="stat-box">
            <div class="num" id="statTotal">—</div>
            <div class="lbl">строк обработано</div>
          </div>
          <div class="stat-box">
            <div class="num green" id="statAuto">—</div>
            <div class="lbl">заполнено авто</div>
          </div>
          <div class="stat-box">
            <div class="num orange" id="statReq">—</div>
            <div class="lbl">на запрос</div>
          </div>
        </div>

        <!-- Скачивание -->
        <div class="dl-section">
          <div class="dl-title">Результаты</div>
          <div class="dl-grid">
            <a class="dl-card blue" id="dlAnalysis" href="#" download>
              <div class="dl-ico"><i class="bi bi-table"></i></div>
              <div class="dl-info">
                <div class="dl-name">Файл анализа</div>
                <div class="dl-desc">Цвета + Анализ MR Group</div>
              </div>
            </a>
            <a class="dl-card green" id="dlRequest" href="#" download>
              <div class="dl-ico"><i class="bi bi-file-earmark-text"></i></div>
              <div class="dl-info">
                <div class="dl-name">Файл запроса</div>
                <div class="dl-desc">По разделам для уточнения</div>
              </div>
            </a>
          </div>
        </div>

        <!-- Лог -->
        <div class="log-section">
          <button class="log-toggle" id="logToggle">
            <i class="bi bi-chevron-down"></i> Подробный лог
          </button>
          <div class="log-body" id="logBody"></div>
        </div>

      </div><!-- /result-section -->

      <!-- Ошибка -->
      <div class="error-box" id="errorBox"></div>

    </div><!-- /card -->
  </main>

</div><!-- /app-layout -->

<script>
  const dropZone   = document.getElementById('dropZone');
  const fileInput  = document.getElementById('fileInput');
  const filePill   = document.getElementById('filePill');
  const fName      = document.getElementById('fName');
  const removeFile = document.getElementById('removeFile');
  const btn        = document.getElementById('btn');
  const chooseLink = document.getElementById('chooseLink');
  let selectedFile = null;

  chooseLink.onclick = e => { e.stopPropagation(); fileInput.click(); };
  dropZone.onclick   = () => fileInput.click();
  dropZone.addEventListener('dragover',  e => { e.preventDefault(); dropZone.classList.add('drag'); });
  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag'));
  dropZone.addEventListener('drop', e => {
    e.preventDefault(); dropZone.classList.remove('drag');
    const f = e.dataTransfer.files[0];
    if (f && f.name.endsWith('.xlsx')) setFile(f);
  });
  fileInput.onchange = () => { if (fileInput.files[0]) setFile(fileInput.files[0]); };
  removeFile.onclick = e => {
    e.stopPropagation();
    selectedFile = null; fileInput.value = '';
    filePill.style.display = 'none'; btn.disabled = true;
  };

  function setFile(f) {
    selectedFile = f;
    fName.textContent = f.name;
    filePill.style.display = 'flex';
    btn.disabled = false;
    document.getElementById('resultSection').style.display = 'none';
    document.getElementById('errorBox').style.display = 'none';
  }

  // Прогресс
  let pTimer = null;
  const STAGES = [
    [0,  12, 700,  'Читаем файл…'],
    [12, 35, 1100, 'Определяем столбцы…'],
    [35, 70, 2200, 'Анализируем строки…'],
    [70, 88, 1400, 'Формируем файлы…'],
  ];
  function startProgress() {
    const fill  = document.getElementById('progressFill');
    const label = document.getElementById('progressLabel');
    const pct   = document.getElementById('progressPct');
    fill.style.width = '0%'; pct.textContent = '0%';
    let i = 0;
    const run = () => {
      if (i >= STAGES.length) return;
      const [, to, dur, text] = STAGES[i++];
      label.textContent = text;
      fill.style.transition = 'width ' + dur + 'ms ease';
      fill.style.width = to + '%';
      pct.textContent = to + '%';
      pTimer = setTimeout(run, dur + 80);
    };
    run();
  }
  function stopProgress() {
    clearTimeout(pTimer);
    const fill = document.getElementById('progressFill');
    const pct  = document.getElementById('progressPct');
    fill.style.transition = 'width .3s ease';
    fill.style.width = '100%'; pct.textContent = '100%';
  }

  btn.onclick = async () => {
    if (!selectedFile) return;
    btn.disabled = true;
    document.getElementById('resultSection').style.display = 'none';
    document.getElementById('errorBox').style.display      = 'none';
    document.getElementById('progressSection').style.display = 'block';
    startProgress();

    const fd = new FormData();
    fd.append('file', selectedFile);

    try {
      const res  = await fetch('/analyze', { method: 'POST', body: fd });
      const data = await res.json();
      stopProgress();
      await new Promise(r => setTimeout(r, 350));
      document.getElementById('progressSection').style.display = 'none';

      if (!res.ok) throw new Error(data.error || 'Неизвестная ошибка');

      const m = data.log.match(/Обработано строк:\\s+(\\d+)[\\s\\S]*?авто[^:]+:\\s+(\\d+)[\\s\\S]*?уточнения:\\s+(\\d+)/);
      if (m) {
        document.getElementById('statTotal').textContent = Number(m[1]).toLocaleString('ru');
        document.getElementById('statAuto').textContent  = Number(m[2]).toLocaleString('ru');
        document.getElementById('statReq').textContent   = Number(m[3]).toLocaleString('ru');
      }

      document.getElementById('dlAnalysis').href     = '/download/' + encodeURIComponent(data.analysisFile);
      document.getElementById('dlAnalysis').download = data.analysisFile;
      document.getElementById('dlRequest').href      = '/download/' + encodeURIComponent(data.requestFile);
      document.getElementById('dlRequest').download  = data.requestFile;
      document.getElementById('logBody').textContent = data.log;
      document.getElementById('resultSection').style.display = 'block';
      btn.disabled = false;
    } catch (err) {
      stopProgress();
      document.getElementById('progressSection').style.display = 'none';
      const box = document.getElementById('errorBox');
      box.textContent = '⚠️  ' + err.message;
      box.style.display = 'block';
      btn.disabled = false;
    }
  };

  document.getElementById('logToggle').onclick = function() {
    const body = document.getElementById('logBody');
    const open = body.style.display === 'block';
    body.style.display = open ? 'none' : 'block';
    this.classList.toggle('open', !open);
  };
</script>
</body>
</html>`;

app.get('/', (req, res) => res.send(HTML));

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
