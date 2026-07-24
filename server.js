import express from 'express';
import cors from 'cors';
import multer from 'multer';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { marked } from 'marked';
import mammoth from 'mammoth';
import HTMLtoDOCX from 'html-to-docx';
import archiver from 'archiver';
import AdmZip from 'adm-zip';
import { pdfToDocx, docxToPdf, htmlToPdf } from './pdf_engine/bridge.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const AUTH_TOKEN = process.env.FILE_TOOL_TOKEN || '';
app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true }));

const upload = multer({ dest: path.join(__dirname, 'uploads') });
const outDir = path.join(__dirname, 'outputs');
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'ui.html')));

// multer 上传文件无扩展名，复制一份带正确扩展名供 Python 识别
function withExt(uploadedPath, ext) {
  const dst = uploadedPath + ext;
  fs.copyFileSync(uploadedPath, dst);
  return dst;
}

function requireAuth(req, res, next) {
  if (!AUTH_TOKEN) return next();
  return req.headers.authorization === `Bearer ${AUTH_TOKEN}`
    ? next()
    : res.status(401).json({ error: 'Unauthorized' });
}

// ============================================================
// HTML → PDF   (Word/WPS COM)
// ============================================================
app.post('/convert/html-to-pdf', requireAuth, upload.single('file'), async (req, res) => {
  try {
    const htmlContent = req.file
      ? fs.readFileSync(req.file.path, 'utf8')
      : (req.body.html || '');
    const htmlFile = path.join(outDir, `_html_${Date.now()}.html`);
    const outPath = path.join(outDir, `output_${Date.now()}.pdf`);
    fs.writeFileSync(htmlFile, htmlContent, 'utf8');
    try {
      await htmlToPdf(htmlFile, outPath);
      res.download(outPath, 'output.pdf');
    } finally {
      try { fs.unlinkSync(htmlFile); } catch {}
    }
  } catch (e) {
    res.status(500).json({ error: String(e) });
  }
});

// ============================================================
// HTML → DOCX   (html-to-docx)
// ============================================================
app.post('/convert/html-to-docx', requireAuth, upload.single('file'), async (req, res) => {
  try {
    const htmlContent = req.file
      ? fs.readFileSync(req.file.path, 'utf8')
      : (req.body.html || '');
    const buffer = await HTMLtoDOCX(htmlContent, null, { table: { row: { cantSplit: true } } });
    const outPath = path.join(outDir, `output_${Date.now()}.docx`);
    fs.writeFileSync(outPath, buffer);
    res.download(outPath, 'output.docx');
  } catch (e) {
    res.status(500).json({ error: String(e) });
  }
});

// ============================================================
// DOCX → HTML   (mammoth)
// ============================================================
app.post('/convert/docx-to-html', requireAuth, upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No DOCX uploaded' });
    const buf = fs.readFileSync(req.file.path);
    const { value: html } = await mammoth.convertToHtml({ buffer: buf });
    const outPath = path.join(outDir, `output_${Date.now()}.html`);
    fs.writeFileSync(outPath, html, 'utf8');
    res.download(outPath, 'output.html');
  } catch (e) {
    res.status(500).json({ error: String(e) });
  }
});

// ============================================================
// DOCX → PDF   (Word/WPS COM)
// ============================================================
app.post('/convert/docx-to-pdf', requireAuth, upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No DOCX uploaded' });
    const srcPath = withExt(req.file.path, '.docx');
    const outPath = path.join(outDir, `docx_${Date.now()}.pdf`);
    try {
      await docxToPdf(srcPath, outPath);
      res.download(outPath, 'output.pdf');
    } finally {
      try { fs.unlinkSync(srcPath); } catch {}
    }
  } catch (e) {
    res.status(500).json({ error: String(e) });
  }
});

// ============================================================
// PDF → DOCX   (Python pdf2docx/docling)
// ============================================================
app.post('/convert/pdf-to-docx', requireAuth, upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No PDF uploaded' });
    const srcPath = withExt(req.file.path, '.pdf');
    const outPath = path.join(outDir, `pdf_${Date.now()}.docx`);
    try {
      await pdfToDocx(srcPath, outPath);
      res.download(outPath, 'output.docx');
    } finally {
      try { fs.unlinkSync(srcPath); } catch {}
    }
  } catch (e) {
    res.status(500).json({ error: String(e) });
  }
});

// ============================================================
// Markdown → HTML   (marked)
// ============================================================
app.post('/convert/md-to-html', requireAuth, upload.single('file'), (req, res) => {
  try {
    const md = req.file ? fs.readFileSync(req.file.path, 'utf8') : (req.body.md || '');
    const html = marked.parse(md);
    const outPath = path.join(outDir, `output_${Date.now()}.html`);
    fs.writeFileSync(outPath, html, 'utf8');
    res.download(outPath, 'output.html');
  } catch (e) {
    res.status(500).json({ error: String(e) });
  }
});

// ============================================================
// Markdown → PDF   (marked + Word/WPS COM)
// ============================================================
app.post('/convert/md-to-pdf', requireAuth, upload.single('file'), async (req, res) => {
  try {
    const md = req.file ? fs.readFileSync(req.file.path, 'utf8') : (req.body.md || '');
    const html = marked.parse(md);
    const htmlFile = path.join(outDir, `_mdhtml_${Date.now()}.html`);
    const outPath = path.join(outDir, `md_${Date.now()}.pdf`);
    fs.writeFileSync(htmlFile, html, 'utf8');
    try {
      await htmlToPdf(htmlFile, outPath);
      res.download(outPath, 'markdown.pdf');
    } finally {
      try { fs.unlinkSync(htmlFile); } catch {}
    }
  } catch (e) {
    res.status(500).json({ error: String(e) });
  }
});

// ============================================================
// Batch ZIP conversion
// ============================================================
app.post('/convert/batch-zip', requireAuth, upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Upload a zip file' });
    const zip = new AdmZip(req.file.path);
    const tempDir = path.join(outDir, `batch_${Date.now()}`);
    fs.mkdirSync(tempDir, { recursive: true });
    zip.extractAllTo(tempDir, true);

    const entries = fs.readdirSync(tempDir, { withFileTypes: true });
    const resultsDir = path.join(tempDir, 'results');
    fs.mkdirSync(resultsDir);

    for (const entry of entries) {
      if (!entry.isFile()) continue;
      const p = path.join(tempDir, entry.name);
      const ext = path.extname(entry.name).toLowerCase();
      const base = path.basename(entry.name, ext);

      if (ext === '.html' || ext === '.htm') {
        const html = fs.readFileSync(p, 'utf8');
        // → PDF (Word COM)
        try { await htmlToPdf(p, path.join(resultsDir, base + '.pdf')); } catch (e) { console.error(`[batch] html→pdf: ${e.message}`); }
        // → DOCX
        try { const b = await HTMLtoDOCX(html); fs.writeFileSync(path.join(resultsDir, base + '.docx'), b); } catch (e) { console.error(`[batch] html→docx: ${e.message}`); }

      } else if (ext === '.md' || ext === '.markdown') {
        const md = fs.readFileSync(p, 'utf8');
        const html = marked.parse(md);
        // → HTML
        fs.writeFileSync(path.join(resultsDir, base + '.html'), html, 'utf8');
        // → PDF (Word COM via temp html)
        const tmpHtml = path.join(resultsDir, `_${base}.html`);
        fs.writeFileSync(tmpHtml, html, 'utf8');
        try { await htmlToPdf(tmpHtml, path.join(resultsDir, base + '.pdf')); } catch (e) { console.error(`[batch] md→pdf: ${e.message}`); }
        try { fs.unlinkSync(tmpHtml); } catch {}

      } else if (ext === '.docx') {
        // → HTML
        try {
          const buf = fs.readFileSync(p);
          const { value: html } = await mammoth.convertToHtml({ buffer: buf });
          fs.writeFileSync(path.join(resultsDir, base + '.html'), html, 'utf8');
        } catch (e) { console.error(`[batch] docx→html: ${e.message}`); }
        // → PDF (Word COM)
        try { await docxToPdf(p, path.join(resultsDir, base + '.pdf')); } catch (e) { console.error(`[batch] docx→pdf: ${e.message}`); }

      } else if (ext === '.pdf') {
        // → DOCX (Python)
        try { await pdfToDocx(p, path.join(resultsDir, base + '.docx')); } catch (e) { console.error(`[batch] pdf→docx: ${e.message}`); }
      }
    }

    const outZipPath = path.join(outDir, `batch_out_${Date.now()}.zip`);
    const output = fs.createWriteStream(outZipPath);
    const archive = archiver('zip', { zlib: { level: 9 } });
    archive.pipe(output);
    archive.directory(resultsDir, false);
    await archive.finalize();
    output.on('close', () => res.download(outZipPath, 'batch_results.zip'));
  } catch (e) {
    res.status(500).json({ error: String(e) });
  }
});

// ============================================================
// Find & Replace
// ============================================================
app.post('/edit/replace', requireAuth, upload.single('file'), (req, res) => {
  try {
    const { find, replace } = req.body;
    const content = req.file ? fs.readFileSync(req.file.path, 'utf8') : (req.body.text || '');
    const out = content.replace(new RegExp(find, 'g'), replace ?? '');
    const outPath = path.join(outDir, `edited_${Date.now()}.txt`);
    fs.writeFileSync(outPath, out, 'utf8');
    res.download(outPath, 'edited.txt');
  } catch (e) {
    res.status(500).json({ error: String(e) });
  }
});

const PORT = parseInt(process.env.PORT || '3456', 10);
app.listen(PORT, () => console.log(`DocForge → http://localhost:${PORT}`));
