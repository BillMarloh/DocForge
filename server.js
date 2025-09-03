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

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
// 可选：Bearer Token 鉴权
const AUTH_TOKEN = process.env.FILE_TOOL_TOKEN || '';
app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true }));

const upload = multer({ dest: path.join(__dirname, 'uploads') });
const outDir = path.join(__dirname, 'outputs');
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'ui.html'));
});

// 鉴权中间件（可选）
function requireAuth(req, res, next){
  if(!AUTH_TOKEN) return next();
  const auth = req.headers.authorization || '';
  if(auth === `Bearer ${AUTH_TOKEN}`) return next();
  return res.status(401).json({ error: 'Unauthorized' });
}

// 注入主题、页眉、页脚与页码
function applyTheme(html, { themeCss = '', headerHtml = '', footerHtml = '' } = {}){
  const baseCss = `
    <style>
      @page { size: A4; margin: 20mm 16mm 20mm 16mm; }
      body { font-family: -apple-system, BlinkMacSystemFont, Segoe UI, Roboto, Inter, Arial, sans-serif; color: #111827; }
      ${themeCss}
    </style>
  `;
  // Puppeteer 页眉/页脚使用独立模板，通过 pdf() 的 displayHeaderFooter、headerTemplate、footerTemplate 设置
  return baseCss + html;
}

// HTML -> PDF
app.post('/convert/html-to-pdf', requireAuth, upload.single('file'), async (req, res) => {
  try {
    const htmlPath = req.file ? req.file.path : null;
    let htmlContent = req.body.html || '';
    if (htmlPath) htmlContent = fs.readFileSync(htmlPath, 'utf8');
    const themeCss = req.body.themeCss || '';
    const headerHtml = req.body.headerHtml || '';
    const footerHtml = req.body.footerHtml || '';

    const html = applyTheme(htmlContent, { themeCss, headerHtml, footerHtml });
    const puppeteer = await import('puppeteer');
    const browser = await puppeteer.default.launch({ headless: 'new', args: ['--no-sandbox'] });
    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: 'networkidle0' });
    const outPath = path.join(outDir, `output_${Date.now()}.pdf`);
    await page.pdf({
      path: outPath,
      format: 'A4',
      printBackground: true,
      displayHeaderFooter: Boolean(headerHtml || footerHtml),
      headerTemplate: headerHtml || '<div></div>',
      footerTemplate: footerHtml || '<div style="font-size:10px;width:100%;text-align:center;"><span class="pageNumber"></span>/<span class="totalPages"></span></div>',
      margin: { top: '24mm', right: '16mm', bottom: '24mm', left: '16mm' },
    });
    await browser.close();
    res.download(outPath, 'output.pdf');
  } catch (e) {
    res.status(500).json({ error: String(e) });
  }
});

// HTML -> DOCX（跨平台，基于 html-to-docx）
app.post('/convert/html-to-docx', requireAuth, upload.single('file'), async (req, res) => {
  try {
    const htmlContent = req.file ? fs.readFileSync(req.file.path, 'utf8') : (req.body.html || '');
    const buffer = await HTMLtoDOCX(htmlContent, null, { table: { row: { cantSplit: true } } });
    const outPath = path.join(outDir, `output_${Date.now()}.docx`);
    fs.writeFileSync(outPath, buffer);
    res.download(outPath, 'output.docx');
  } catch (e) {
    res.status(500).json({ error: String(e) });
  }
});

// DOCX -> HTML（跨平台，基于 mammoth）
app.post('/convert/docx-to-html', requireAuth, upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No DOCX uploaded' });
    const arrayBuffer = fs.readFileSync(req.file.path);
    const { value: html } = await mammoth.convertToHtml({ buffer: arrayBuffer });
    const outPath = path.join(outDir, `output_${Date.now()}.html`);
    fs.writeFileSync(outPath, html, 'utf8');
    res.download(outPath, 'output.html');
  } catch (e) {
    res.status(500).json({ error: String(e) });
  }
});

// Markdown -> HTML/PDF
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

// Markdown -> PDF
app.post('/convert/md-to-pdf', requireAuth, upload.single('file'), async (req, res) => {
  try {
    const md = req.file ? fs.readFileSync(req.file.path, 'utf8') : (req.body.md || '');
    const html = marked.parse(md);
    const puppeteer = await import('puppeteer');
    const browser = await puppeteer.default.launch({ headless: 'new', args: ['--no-sandbox'] });
    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: 'networkidle0' });
    const outPath = path.join(outDir, `md_${Date.now()}.pdf`);
    await page.pdf({ path: outPath, format: 'A4', printBackground: true, margin: { top: '20mm', right: '16mm', bottom: '20mm', left: '16mm' } });
    await browser.close();
    res.download(outPath, 'markdown.pdf');
  } catch (e) {
    res.status(500).json({ error: String(e) });
  }
});

// 批量目录转换：上传 ZIP，按规则批量转换后返回打包 ZIP
// 支持：.html → pdf/.docx，.md → html/pdf，.docx → html
app.post('/convert/batch-zip', requireAuth, upload.single('file'), async (req, res) => {
  try {
    if(!req.file) return res.status(400).json({ error: 'Upload a zip file' });
    const zip = new AdmZip(req.file.path);
    const tempDir = path.join(outDir, `batch_${Date.now()}`);
    fs.mkdirSync(tempDir, { recursive: true });
    zip.extractAllTo(tempDir, true);

    const entries = fs.readdirSync(tempDir, { withFileTypes: true });
    const resultsDir = path.join(tempDir, 'results');
    fs.mkdirSync(resultsDir);

    const puppeteer = await import('puppeteer');
    const browser = await puppeteer.default.launch({ headless: 'new', args: ['--no-sandbox'] });
    const page = await browser.newPage();

    for(const entry of entries){
      if(!entry.isFile()) continue;
      const p = path.join(tempDir, entry.name);
      const ext = path.extname(entry.name).toLowerCase();
      const base = path.basename(entry.name, ext);
      if(ext === '.html' || ext === '.htm'){
        const html = fs.readFileSync(p, 'utf8');
        await page.setContent(html, { waitUntil: 'networkidle0' });
        await page.pdf({ path: path.join(resultsDir, base + '.pdf'), format: 'A4', printBackground: true });
        const docxBuf = await HTMLtoDOCX(html);
        fs.writeFileSync(path.join(resultsDir, base + '.docx'), docxBuf);
      } else if(ext === '.md' || ext === '.markdown'){
        const md = fs.readFileSync(p, 'utf8');
        const html = marked.parse(md);
        fs.writeFileSync(path.join(resultsDir, base + '.html'), html, 'utf8');
        await page.setContent(html, { waitUntil: 'networkidle0' });
        await page.pdf({ path: path.join(resultsDir, base + '.pdf'), format: 'A4', printBackground: true });
      } else if(ext === '.docx'){
        const buf = fs.readFileSync(p);
        const { value: html } = await mammoth.convertToHtml({ buffer: buf });
        fs.writeFileSync(path.join(resultsDir, base + '.html'), html, 'utf8');
      }
    }

    await browser.close();

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

// 简单文本查找替换（对上传的文本/HTML）
app.post('/edit/replace', requireAuth, upload.single('file'), (req, res) => {
  try {
    const { find, replace } = req.body;
    const content = req.file ? fs.readFileSync(req.file.path, 'utf8') : (req.body.text || '');
    const re = new RegExp(find, 'g');
    const out = content.replace(re, replace ?? '');
    const outPath = path.join(outDir, `edited_${Date.now()}.txt`);
    fs.writeFileSync(outPath, out, 'utf8');
    res.download(outPath, 'edited.txt');
  } catch (e) {
    res.status(500).json({ error: String(e) });
  }
});

const PORT = parseInt(process.env.PORT || '3456', 10);
const server = app.listen(PORT, () => {
  console.log(`File Converter listening on http://localhost:${PORT}`);
});
server.on('error', (err) => {
  if (err && err.code === 'EADDRINUSE') {
    console.error(`Port ${PORT} is in use. 设置环境变量 PORT 更换端口，或先停止已有进程。`);
  }
});


