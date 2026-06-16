require('dotenv').config();
const express = require('express');
const multer  = require('multer');
const path    = require('path');
const fs      = require('fs');
const analyze  = require('./analyze');
const parseKP  = require('./kpParser');
const applyKP  = require('./kpApplier');
const db       = require('./lib/db');

const app  = express();
const PORT = process.env.PORT || 3000;
const TMP  = path.join(__dirname, 'uploads');

if (!fs.existsSync(TMP)) fs.mkdirSync(TMP);

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

const upload = multer({
  dest: TMP,
  fileFilter: (req, file, cb) =>
    cb(null, path.extname(file.originalname).toLowerCase() === '.xlsx'),
});

const uploadKP = multer({
  dest: TMP,
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    cb(null, ['.pdf', '.xlsx', '.xls'].includes(ext));
  },
});

// ── Projects ──────────────────────────────────────────────────────────────────

app.get('/api/projects', async (req, res) => {
  try { res.json(await db.getProjects()); }
  catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/projects', async (req, res) => {
  const { name, description } = req.body;
  if (!name?.trim()) return res.status(400).json({ error: 'Название обязательно' });
  try { res.json(await db.createProject({ name: name.trim(), description: description?.trim() || null })); }
  catch (e) { res.status(500).json({ error: e.message }); }
});

app.patch('/api/projects/:id', async (req, res) => {
  const { name, description } = req.body;
  if (name !== undefined && !name?.trim()) return res.status(400).json({ error: 'Название не может быть пустым' });
  try {
    const updates = {};
    if (name        !== undefined) updates.name        = name.trim();
    if (description !== undefined) updates.description = description?.trim() || null;
    res.json(await db.updateProject(req.params.id, updates));
  }
  catch (e) { res.status(500).json({ error: e.message }); }
});

app.delete('/api/projects/:id', async (req, res) => {
  try { await db.deleteProject(req.params.id); res.json({ ok: true }); }
  catch (e) { res.status(500).json({ error: e.message }); }
});

// ── Analyses ──────────────────────────────────────────────────────────────────

app.get('/api/projects/:id/analyses', async (req, res) => {
  try { res.json(await db.getAnalyses(req.params.id)); }
  catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/projects/:id/analyze', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Файл не передан (.xlsx обязателен)' });

  const projectId = req.params.id;
  const tmpPath   = req.file.path;
  const origName  = Buffer.from(req.file.originalname, 'latin1').toString('utf8');
  let analysisId  = null;
  const logs = [];
  const origLog  = console.log;
  const origWarn = console.warn;
  console.log  = (...a) => { origLog(...a);  logs.push(a.join(' ')); };
  console.warn = (...a) => { origWarn(...a); logs.push(a.join(' ')); };

  try {
    const analysis = await db.createAnalysis({ projectId, filename: origName });
    analysisId = analysis.id;

    const srcBuffer = fs.readFileSync(tmpPath);
    const srcPath   = `${projectId}/${analysisId}/source.xlsx`;
    await db.uploadFile(srcPath, srcBuffer, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    await db.createFile({ analysisId, fileType: 'source', storagePath: srcPath });

    const { analysisPath, requestPath } = await analyze(tmpPath);
    console.log = origLog; console.warn = origWarn;

    const anaPath = `${projectId}/${analysisId}/analysis.xlsx`;
    const reqPath = `${projectId}/${analysisId}/request.xlsx`;
    await db.uploadFile(anaPath, fs.readFileSync(analysisPath), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    await db.uploadFile(reqPath, fs.readFileSync(requestPath),  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    await db.createFile({ analysisId, fileType: 'analysis', storagePath: anaPath });
    await db.createFile({ analysisId, fileType: 'request',  storagePath: reqPath });

    const logStr = logs.join('\n');
    const m = logStr.match(/Обработано строк:\s+(\d+)[\s\S]*?авто[^:]+:\s+(\d+)[\s\S]*?уточнения:\s+(\d+)/);
    const completed = await db.completeAnalysis(analysisId, {
      statsTotal:   m ? parseInt(m[1]) : 0,
      statsAuto:    m ? parseInt(m[2]) : 0,
      statsRequest: m ? parseInt(m[3]) : 0,
    });

    [tmpPath, analysisPath, requestPath].forEach(f => { try { fs.unlinkSync(f); } catch {} });
    res.json({ analysis: completed, log: logStr });

  } catch (err) {
    console.log = origLog; console.warn = origWarn;
    if (analysisId) await db.failAnalysis(analysisId, err.message).catch(() => {});
    try { fs.unlinkSync(tmpPath); } catch {}
    res.status(500).json({ error: err.message });
  }
});

// ── KP Upload ─────────────────────────────────────────────────────────────────

app.post('/api/analyses/:id/upload-kp', (req, res, next) => {
  uploadKP.single('file')(req, res, err => {
    if (err) return res.status(400).json({ error: err.message });
    next();
  });
}, async (req, res) => {
  if (!req.file)
    return res.status(400).json({ error: 'Файл не передан (.pdf, .xlsx или .xls)' });

  const tmpPath  = req.file.path;
  const origName = Buffer.from(req.file.originalname, 'latin1').toString('utf8');

  try {
    const kpItems = await parseKP(tmpPath, origName);
    if (!kpItems.length) {
      fs.unlinkSync(tmpPath);
      return res.status(422).json({
        error: 'Не удалось извлечь позиции из КП. ' +
               'Убедитесь, что файл содержит таблицу с наименованиями и ценами за единицу.',
      });
    }

    const result = await applyKP(req.params.id, kpItems, db);
    fs.unlinkSync(tmpPath);
    res.json(result);
  } catch (err) {
    try { fs.unlinkSync(tmpPath); } catch {}
    res.status(500).json({ error: err.message });
  }
});

// ── Download ──────────────────────────────────────────────────────────────────

app.get('/api/files/:id/download', async (req, res) => {
  try {
    const row = await db.getFile(req.params.id);
    if (!row) return res.status(404).json({ error: 'Файл не найден' });

    const buffer = await db.downloadFile(row.storage_path);
    const suffix = { source: '', analysis: '_analysis', request: '_request' }[row.file_type] || '';
    const base   = path.basename(row.analyses.filename, '.xlsx');
    const fname  = encodeURIComponent(base + suffix + '.xlsx');

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${fname}`);
    res.send(buffer);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.listen(PORT, () => console.log(`Сервер: http://localhost:${PORT}`));
