const express = require('express');
const multer  = require('multer');
const path    = require('path');
const fs      = require('fs');
const analyze = require('./analyze');

const app    = express();
const PORT   = 3000;
const UPLOAD = path.join(__dirname, 'uploads');

app.use(express.static(path.join(__dirname, 'public')));

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
  <link rel="stylesheet" href="/styles.css">
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
