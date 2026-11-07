'use strict';

// Local HTTPS server for the Table Saver Excel add-in.
//  - serves the task pane from ./public
//  - POST /save  { owner, tableName, rows, format: "pdf" | "xlsx" }
//    writes <baseDir>/<owner>/<fileName>.<format> with normal permissions
//    (folders 755, files 644), running as your own user. Existing files are
//    kept: later saves become "<fileName> (1).<format>", "<fileName> (2).<format>", ...

const https = require('https');
const fs = require('fs');
const path = require('path');
const os = require('os');
const PDFDocument = require('pdfkit');
const ExcelJS = require('exceljs');

const PORT = Number(process.env.PORT || 3000);
const HOST = '127.0.0.1'; // never expose outside this machine
const ORIGIN = `https://localhost:${PORT}`;
const PUBLIC_DIR = path.join(__dirname, 'public');

// ---------- config ----------
function expandHome(p) {
  if (p === '~') return os.homedir();
  if (p.startsWith('~/') || p.startsWith('~\\')) return path.join(os.homedir(), p.slice(2));
  return p;
}

const config = JSON.parse(fs.readFileSync(path.join(__dirname, 'config.json'), 'utf8'));
// Relative paths resolve against the project, because a startup service's working
// directory varies by OS (Windows Task Scheduler uses System32).
const BASE_DIR = path.resolve(__dirname, expandHome(process.env.TABLE_SAVER_BASE || config.baseDir || '~/home'));
const FILE_NAME = String(config.fileName || 'table').replace(/[^\w.-]/g, '_');

// Windows can't create files or folders with these names, even with an extension.
// They're rejected on every OS so a sheet behaves the same wherever it's saved.
const WINDOWS_RESERVED_NAME = /^(con|prn|aux|nul|com[0-9]|lpt[0-9])\s*(\..*)?$/i;

if (WINDOWS_RESERVED_NAME.test(FILE_NAME)) {
  throw new Error(`config.json: fileName "${FILE_NAME}" is a reserved name on Windows. Choose another.`);
}

// ---------- TLS (certs created by `npm run certs`) ----------
const certDir = path.join(os.homedir(), '.office-addin-dev-certs');
const tlsOptions = {
  key: fs.readFileSync(path.join(certDir, 'localhost.key')),
  cert: fs.readFileSync(path.join(certDir, 'localhost.crt')),
};

// ---------- helpers ----------
const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.png': 'image/png',
};

function sendJson(res, status, obj) {
  const body = JSON.stringify(obj);
  res.writeHead(status, { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(body) });
  res.end(body);
}

function readBody(req, limit = 5 * 1024 * 1024) {
  return new Promise((resolve, reject) => {
    let size = 0;
    const chunks = [];
    req.on('data', (c) => {
      size += c.length;
      if (size > limit) {
        reject(new Error('Request too large'));
        req.destroy();
        return;
      }
      chunks.push(c);
    });
    req.on('end', () => resolve(Buffer.concat(chunks).toString('utf8')));
    req.on('error', reject);
  });
}

// Owner becomes a folder name, so only allow names that are safe on Windows and Linux:
// no path separators or "..", no trailing dot (Windows strips it), no reserved names.
function safeOwner(owner) {
  const o = String(owner || '').trim();
  if (!/^[\w][\w .-]{0,63}$/.test(o) || o.endsWith('.') || WINDOWS_RESERVED_NAME.test(o)) return null;
  return o;
}

// These modes apply on Linux and macOS. Windows uses ACLs and ignores everything
// except the read-only flag, which 755 and 644 leave off.
function ensureDir(dir) {
  fs.mkdirSync(dir, { recursive: true, mode: 0o755 });
  try { fs.chmodSync(dir, 0o755); } catch (_) { /* e.g. a folder owned by another user */ }
}

function finalizeFile(file) {
  try { fs.chmodSync(file, 0o644); } catch (_) { /* e.g. a file owned by another user */ }
}

// Claims the first free name: table.pdf, table (1).pdf, table (2).pdf, ...
// "wx" fails if the file exists, so simultaneous saves can't claim the same name.
// The writers then fill in the empty file this leaves behind.
function reserveFile(dir, format) {
  for (let n = 0; ; n++) {
    const name = n === 0 ? `${FILE_NAME}.${format}` : `${FILE_NAME} (${n}).${format}`;
    const file = path.join(dir, name);
    try {
      fs.closeSync(fs.openSync(file, 'wx', 0o644));
      return file;
    } catch (e) {
      if (e.code !== 'EEXIST') throw e;
    }
  }
}

function saveErrorMessage(e, dir) {
  if (['EBUSY', 'EPERM', 'EACCES'].includes(e.code)) {
    return `Could not save to ${dir}. Check you have permission to write to that folder.`;
  }
  return e.message;
}

function cell(v) {
  return v === null || v === undefined ? '' : String(v);
}

// ---------- writers ----------
function writePdf(file, tableName, owner, rows) {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ size: 'A4', margin: 40 });
    const out = fs.createWriteStream(file, { mode: 0o644 });
    out.on('finish', resolve);
    out.on('error', reject);
    doc.pipe(out);

    doc.font('Helvetica-Bold').fontSize(16).text(cell(tableName) || 'Table');
    doc.font('Helvetica').fontSize(11).fillColor('#444').text(`Owner: ${owner}`);
    doc.moveDown(0.8).fillColor('#000');

    const cols = Math.max(1, ...rows.map((r) => r.length));
    const left = doc.page.margins.left;
    const width = doc.page.width - left - doc.page.margins.right;
    const colW = width / cols;
    const padX = 4;
    const padY = 4;

    rows.forEach((r, i) => {
      const cells = Array.from({ length: cols }, (_, c) => cell(r[c]));
      doc.font(i === 0 ? 'Helvetica-Bold' : 'Helvetica').fontSize(10);
      const h = Math.max(...cells.map((t) => doc.heightOfString(t, { width: colW - padX * 2 }))) + padY * 2;

      if (doc.y + h > doc.page.height - doc.page.margins.bottom) doc.addPage();
      const y = doc.y;
      if (i === 0) doc.save().rect(left, y, width, h).fill('#EEEEEE').restore();
      cells.forEach((t, c) => {
        const x = left + c * colW;
        doc.rect(x, y, colW, h).strokeColor('#999').lineWidth(0.5).stroke();
        doc.fillColor('#000').text(t, x + padX, y + padY, { width: colW - padX * 2 });
      });
      doc.y = y + h;
      doc.x = left;
    });

    doc.end();
  });
}

async function writeXlsx(file, tableName, owner, rows) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Table');
  ws.addRow([cell(tableName)]).font = { bold: true, size: 14 };
  ws.addRow(['Owner', owner]);
  rows.forEach((r, i) => {
    const row = ws.addRow(r.map((v) => (v === null || v === undefined ? '' : v)));
    if (i === 0) row.font = { bold: true };
  });
  ws.columns.forEach((c) => { c.width = 18; });
  await wb.xlsx.writeFile(file);
}

// ---------- request handlers ----------
async function handleSave(req, res) {
  // same-origin only (blocks other websites from posting to localhost)
  if (req.headers.origin && req.headers.origin !== ORIGIN) {
    return sendJson(res, 403, { error: 'Forbidden origin' });
  }
  if (!String(req.headers['content-type'] || '').startsWith('application/json')) {
    return sendJson(res, 415, { error: 'Expected application/json' });
  }

  let data;
  try {
    data = JSON.parse(await readBody(req));
  } catch (e) {
    return sendJson(res, 400, { error: 'Invalid JSON' });
  }
  if (data === null || typeof data !== 'object') return sendJson(res, 400, { error: 'Expected a JSON object' });

  const owner = safeOwner(data.owner);
  if (!owner) return sendJson(res, 400, { error: 'Owner is empty or not a valid folder name (use letters, numbers, space, . _ -, with no trailing dot and no reserved name such as CON or NUL)' });
  if (data.format !== 'pdf' && data.format !== 'xlsx') return sendJson(res, 400, { error: 'format must be "pdf" or "xlsx"' });
  if (!Array.isArray(data.rows) || data.rows.length === 0 || !data.rows.every(Array.isArray)) {
    return sendJson(res, 400, { error: 'No table rows found' });
  }

  const ownerDir = path.join(BASE_DIR, owner);
  if (path.dirname(ownerDir) !== BASE_DIR) return sendJson(res, 400, { error: 'Bad path' });

  let file;
  try {
    ensureDir(ownerDir);
    file = reserveFile(ownerDir, data.format);
    if (data.format === 'pdf') await writePdf(file, data.tableName, owner, data.rows);
    else await writeXlsx(file, data.tableName, owner, data.rows);
    finalizeFile(file);
    console.log(new Date().toISOString(), 'saved', file);
    return sendJson(res, 200, { ok: true, path: file });
  } catch (e) {
    console.error('save failed:', e);
    // Don't leave an empty or partial file holding this name.
    if (file) fs.rm(file, { force: true }, () => {});
    return sendJson(res, 500, { error: saveErrorMessage(e, ownerDir) });
  }
}

function serveStatic(res, urlPath) {
  const rel = urlPath === '/' ? '/taskpane.html' : urlPath;
  const file = path.normalize(path.join(PUBLIC_DIR, rel));
  if (!file.startsWith(PUBLIC_DIR + path.sep)) {
    res.writeHead(403); return res.end('Forbidden');
  }
  fs.readFile(file, (err, buf) => {
    if (err) { res.writeHead(404); return res.end('Not found'); }
    res.writeHead(200, { 'Content-Type': MIME[path.extname(file)] || 'application/octet-stream', 'Cache-Control': 'no-cache' });
    res.end(buf);
  });
}

function route(req, res) {
  // Any website can make the browser request these URLs, so bad ones must not throw.
  let pathname;
  let decodedPath;
  try {
    pathname = new URL(req.url, ORIGIN).pathname; // throws on "//"
    decodedPath = decodeURIComponent(pathname); // throws on "/%"
  } catch (_) {
    res.writeHead(400); return res.end('Bad request');
  }
  if (decodedPath.includes('\0')) { // file functions throw on null bytes
    res.writeHead(400); return res.end('Bad request');
  }

  if (req.method === 'POST' && pathname === '/save') return handleSave(req, res);
  if (req.method === 'GET' && pathname === '/health') return sendJson(res, 200, { ok: true, baseDir: BASE_DIR });
  if (req.method === 'GET') return serveStatic(res, decodedPath);
  res.writeHead(405); res.end();
}

const server = https.createServer(tlsOptions, (req, res) => {
  // An uncaught error would stop the server, so unexpected failures only end this request.
  Promise.resolve()
    .then(() => route(req, res))
    .catch((e) => {
      console.error('request failed:', e);
      if (res.headersSent) return res.destroy();
      return sendJson(res, 500, { error: 'Internal server error' });
    });
});

server.listen(PORT, HOST, () => {
  console.log(`Table Saver listening on ${ORIGIN}`);
  console.log(`Saving files under ${path.join(BASE_DIR, '<Owner>', `${FILE_NAME}.<pdf|xlsx>`)}`);
});
