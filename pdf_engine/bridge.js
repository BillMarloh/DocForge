/**
 * Node.js → Python Windows Engine Bridge
 *
 * Usage:
 *   import { pdfToDocx, docxToPdf, htmlToPdf } from './pdf_engine/bridge.js';
 */

import { execFile } from 'child_process';
import { promisify } from 'util';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const execFileP = promisify(execFile);
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const CONVERT = path.join(__dirname, 'convert.py');

let _pyReady = undefined;

async function ensurePython() {
  if (_pyReady !== undefined) return _pyReady;
  try {
    const { stdout } = await execFileP('python', ['--version'], { timeout: 5000 });
    const m = stdout.match(/Python (\d+)\.(\d+)/);
    if (!m || +m[1] < 3 || (+m[1] === 3 && +m[2] < 11)) {
      console.log(`[bridge] Python too old: ${stdout.trim()}`);
      _pyReady = false; return false;
    }
    for (const dep of ['fitz', 'pdf2docx']) {
      await execFileP('python', ['-c', `import ${dep}`], { timeout: 10000 });
    }
    console.log(`[bridge] Python ready (${stdout.trim()})`);
    _pyReady = true; return true;
  } catch (e) {
    _pyReady = false; return false;
  }
}

async function run(src, dst, timeout = 120000) {
  if (!(await ensurePython())) throw new Error('Python engine not available');
  if (!fs.existsSync(src)) throw new Error(`File not found: ${src}`);
  const { stdout, stderr } = await execFileP('python', [CONVERT, src, dst], { timeout });
  if (stderr) console.error('[bridge]', stderr);
  const out = stdout.trim();
  console.log('[bridge]', out);
  if (fs.existsSync(out || dst) && fs.statSync(out || dst).size > 0) {
    return out || dst;
  }
  throw new Error('Empty output');
}

export async function pdfToDocx(src, dst)  { return run(src, dst, 300000); }
export async function docxToPdf(src, dst)  { return run(src, dst, 120000); }
export async function htmlToPdf(src, dst)  { return run(src, dst, 60000); }

export async function getEngineStatus() {
  return { python: await ensurePython(), script: CONVERT };
}
