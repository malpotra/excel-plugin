#!/usr/bin/env node
'use strict';

// End-to-end check of server.js. Starts its own copy on a spare port, saving to
// a temp folder, so it never touches a running server or real output files.
// Usage: node .claude/skills/verify-server/scripts/smoke-test.js
// Exit code: 0 = all passed, 1 = at least one failure.

const { spawn } = require('child_process');
const fs = require('fs');
const https = require('https');
const os = require('os');
const path = require('path');

const ROOT = path.resolve(__dirname, '../../../..');
const PORT = Number(process.env.SMOKE_PORT || 3987);
const ORIGIN = `https://localhost:${PORT}`;
// BASE is nested so a traversal bug (owner "..") still lands inside TEMP_ROOT and gets cleaned up.
const TEMP_ROOT = fs.mkdtempSync(path.join(os.tmpdir(), 'table-saver-smoke-'));
const BASE = path.join(TEMP_ROOT, 'base');
fs.mkdirSync(BASE);
const FILE_NAME = JSON.parse(fs.readFileSync(path.join(ROOT, 'config.json'), 'utf8')).fileName || 'table';
const ExcelJS = require(path.join(ROOT, 'node_modules/exceljs'));

let failures = 0;
function check(name, ok, detail) {
  if (!ok) failures++;
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${!ok && detail !== undefined ? `  -> ${detail}` : ''}`);
}

function request(method, urlPath, { body, headers, port = PORT } = {}) {
  return new Promise((resolve, reject) => {
    const req = https.request(
      // The dev cert is only trusted through the OS certificate store, which Node ignores.
      { host: '127.0.0.1', port, method, path: urlPath, headers, rejectUnauthorized: false },
      (res) => {
        const chunks = [];
        res.on('data', (c) => chunks.push(c));
        res.on('end', () => resolve({ status: res.statusCode, text: Buffer.concat(chunks).toString('utf8') }));
      });
    req.on('error', reject);
    if (body !== undefined) req.write(body);
    req.end();
  });
}

function postSave(payload, headers = { 'Content-Type': 'application/json' }) {
  const body = typeof payload === 'string' ? payload : JSON.stringify(payload);
  return request('POST', '/save', { body, headers });
}

async function waitForServer(child, port = PORT) {
  for (let i = 0; i < 50; i++) {
    if (child.exitCode !== null) return false;
    try {
      if ((await request('GET', '/health', { port })).status === 200) return true;
    } catch (_) { /* not listening yet */ }
    await new Promise((r) => setTimeout(r, 100));
  }
  return false;
}

function stopServer(child) {
  return new Promise((resolve) => {
    if (child.exitCode !== null) return resolve();
    child.once('exit', resolve);
    child.kill();
  });
}

// A relative output folder must not depend on the working directory: services
// start elsewhere (Windows Task Scheduler uses System32).
async function testRelativeBaseDir() {
  const port = PORT + 1;
  const child = spawn(process.execPath, [path.join(ROOT, 'server.js')], {
    cwd: TEMP_ROOT,
    env: { ...process.env, PORT: String(port), TABLE_SAVER_BASE: 'smoke-relative' },
    stdio: 'ignore',
  });
  try {
    const up = await waitForServer(child, port);
    check('server starts from another working directory', up);
    if (up) {
      const { baseDir } = JSON.parse((await request('GET', '/health', { port })).text);
      check('relative output folder resolves against the project', baseDir === path.join(ROOT, 'smoke-relative'), baseDir);
    }
  } finally {
    await stopServer(child);
  }
}

const SAMPLE = {
  owner: 'Smoke Owner',
  tableName: 'Smoke table',
  rows: [['SNO', 'Column'], [1, 'A'], [2, 'B']],
};

async function testManifest() {
  const manifest = fs.readFileSync(path.join(ROOT, 'manifest.xml'), 'utf8');
  const urls = [...new Set(manifest.match(/https:\/\/localhost:\d+[^"<]*/g) || [])];
  const defaultPort = (fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8').match(/process\.env\.PORT \|\| (\d+)/) || [])[1];
  const wrong = urls.filter((u) => new URL(u).port !== defaultPort);
  check('manifest URLs all use the server default port', urls.length > 0 && wrong.length === 0, wrong.join(', '));

  for (const u of urls) {
    const { pathname } = new URL(u);
    if (pathname === '/') continue;
    const r = await request('GET', pathname);
    check(`manifest resource is served: ${pathname}`, r.status === 200, r.status);
  }
}

async function testStatic() {
  check('GET / serves the task pane', (await request('GET', '/')).text.includes('Table Saver'));
  check('GET /taskpane.js', (await request('GET', '/taskpane.js')).status === 200);

  const traversal = await request('GET', '/..%2f..%2fserver.js');
  check('static path traversal is blocked', [403, 404].includes(traversal.status), traversal.status);
  const backslash = await request('GET', '/..%5c..%5cserver.js');
  check('backslash path traversal is blocked', [403, 404].includes(backslash.status), backslash.status);

  check('unsupported method gets 405', (await request('PUT', '/save')).status === 405);
}

async function testSaves() {
  const dir = path.join(BASE, SAMPLE.owner);

  const pdf = await postSave({ ...SAMPLE, format: 'pdf' });
  const pdfFile = path.join(dir, `${FILE_NAME}.pdf`);
  check('save pdf returns 200 and its path', pdf.status === 200 && JSON.parse(pdf.text).path === pdfFile, pdf.text);
  const pdfBytes = fs.existsSync(pdfFile) ? fs.readFileSync(pdfFile, 'latin1') : '';
  check('pdf file is complete', pdfBytes.startsWith('%PDF-') && pdfBytes.trimEnd().endsWith('%%EOF'));

  const xlsx = await postSave({ ...SAMPLE, format: 'xlsx' });
  const xlsxFile = path.join(dir, `${FILE_NAME}.xlsx`);
  check('save xlsx returns 200 and its path', xlsx.status === 200 && JSON.parse(xlsx.text).path === xlsxFile, xlsx.text);
  if (fs.existsSync(xlsxFile)) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(xlsxFile);
    const ws = wb.getWorksheet('Table');
    check('xlsx has title, owner and table rows',
      ws.getCell('A1').value === SAMPLE.tableName
        && ws.getCell('B2').value === SAMPLE.owner
        && ws.getCell('A3').value === 'SNO'
        && ws.getCell('B5').value === 'B');
  } else {
    check('xlsx file exists', false);
  }

  if (process.platform !== 'win32') {
    const mode = (p) => (fs.statSync(p).mode & 0o777).toString(8);
    check('owner folder is 755', mode(dir) === '755', mode(dir));
    if (fs.existsSync(pdfFile)) check('pdf is 644', mode(pdfFile) === '644', mode(pdfFile));
    if (fs.existsSync(xlsxFile)) check('xlsx is 644', mode(xlsxFile) === '644', mode(xlsxFile));
  }

  // Long tables must paginate rather than fail.
  const long = await postSave({
    ...SAMPLE,
    owner: 'Long Table',
    format: 'pdf',
    rows: [['SNO', 'Text'], ...Array.from({ length: 300 }, (_, i) => [i, 'wrapped text '.repeat(10)])],
  });
  check('multi-page pdf saves', long.status === 200, long.text);
}

// Each of these once stopped the server. Any website can trigger the GET requests.
async function testMalformedRequests() {
  for (const urlPath of ['/%', '/%zz', '//', '///', '/%00', '/a%00b.html']) {
    const r = await request('GET', urlPath).catch((e) => ({ status: e.code }));
    check(`malformed URL gets 400: ${urlPath}`, r.status === 400, r.status);
  }
  for (const body of ['null', '1', '"x"', '[]']) {
    const r = await postSave(body).catch((e) => ({ status: e.code }));
    check(`non-object JSON body gets 400: ${body}`, r.status === 400, r.status);
  }
  check('server still answers after malformed requests', (await request('GET', '/health').catch(() => ({}))).status === 200);
}

async function testCopies() {
  for (const format of ['pdf', 'xlsx']) {
    const owner = `Copies ${format}`;
    const dir = path.join(BASE, owner);
    const names = [`${FILE_NAME}.${format}`, `${FILE_NAME} (1).${format}`, `${FILE_NAME} (2).${format}`];
    const contents = [];
    for (const name of names) {
      const r = await postSave({ ...SAMPLE, owner, format });
      check(`${format} save is named "${name}"`, r.status === 200 && JSON.parse(r.text).path === path.join(dir, name), r.text);
      contents.push(fs.readFileSync(path.join(dir, name)));
    }
    check(`earlier ${format} copies are left unchanged`,
      names.every((name, i) => fs.readFileSync(path.join(dir, name)).equals(contents[i])));
  }

  // A deleted copy's number gets reused, and nothing else is disturbed.
  const dir = path.join(BASE, 'Copies pdf');
  fs.rmSync(path.join(dir, `${FILE_NAME} (1).pdf`));
  const refill = await postSave({ ...SAMPLE, owner: 'Copies pdf', format: 'pdf' });
  check('a freed number is reused', JSON.parse(refill.text).path === path.join(dir, `${FILE_NAME} (1).pdf`), refill.text);

  // Simultaneous saves must each get their own complete file.
  const results = await Promise.all(Array.from({ length: 8 }, () => postSave({ ...SAMPLE, owner: 'Parallel', format: 'pdf' })));
  const paths = results.map((r) => JSON.parse(r.text).path);
  check('simultaneous saves get separate files', new Set(paths).size === 8, paths.join());
  check('simultaneous saves are all complete',
    paths.every((p) => p && fs.readFileSync(p, 'latin1').trimEnd().endsWith('%%EOF')));

  // A read-only file is kept, and the save goes to a new copy.
  const roDir = path.join(BASE, 'Read Only');
  fs.mkdirSync(roDir);
  const roFile = path.join(roDir, `${FILE_NAME}.pdf`);
  fs.writeFileSync(roFile, 'existing');
  fs.chmodSync(roFile, 0o444);
  const ro = await postSave({ ...SAMPLE, owner: 'Read Only', format: 'pdf' });
  check('read-only file gets a new copy', ro.status === 200 && JSON.parse(ro.text).path !== roFile, ro.text);
  check('read-only file is left unchanged', fs.readFileSync(roFile, 'utf8') === 'existing');
  fs.chmodSync(roFile, 0o644); // Windows can't delete a read-only file during cleanup
}

// Skipped on Windows, where modes don't restrict folders, and as root, which ignores them.
async function testUnwritableFolder() {
  if (process.platform === 'win32' || process.getuid() === 0) return;
  fs.chmodSync(BASE, 0o555);
  try {
    const r = await postSave({ ...SAMPLE, owner: 'No Permission', format: 'pdf' });
    check('unwritable folder gets a clear error', r.status === 500 && JSON.parse(r.text).error.includes('permission'), r.text);
  } finally {
    fs.chmodSync(BASE, 0o755);
  }
}

async function testRejections() {
  const before = fs.readdirSync(BASE).sort().join();

  const unsafeOwners = [
    '../evil', '..', '.', '', '   ', 'a/b', 'a\\b', '/etc', 'x'.repeat(65),
    // Invalid on Windows, and rejected everywhere so a sheet behaves the same on every OS.
    'C:', 'a:b', 'a*b', 'a?b', 'a"b', 'a<b', 'a|b', 'trailing.', 'CON', 'con', 'NUL.txt', 'Com1', 'LPT9.pdf',
  ];
  for (const owner of unsafeOwners) {
    const r = await postSave({ ...SAMPLE, owner, format: 'pdf' });
    check(`owner rejected: ${JSON.stringify(owner.slice(0, 12))}`, r.status === 400, r.status);
  }
  const cases = [
    ['unknown format', { ...SAMPLE, format: 'exe' }],
    ['empty rows', { ...SAMPLE, rows: [], format: 'pdf' }],
    ['rows that are not arrays', { ...SAMPLE, rows: ['nope'], format: 'pdf' }],
    ['invalid JSON', '{not json'],
  ];
  for (const [name, payload] of cases) {
    const r = await postSave(payload);
    check(`rejected: ${name}`, r.status === 400, r.status);
  }

  const payload = { ...SAMPLE, owner: 'Blocked', format: 'pdf' };
  const foreign = await postSave(payload, { 'Content-Type': 'application/json', Origin: 'https://evil.example' });
  check('foreign Origin gets 403', foreign.status === 403, foreign.status);
  const sameOrigin = await postSave({ ...SAMPLE, owner: 'Same Origin', format: 'pdf' },
    { 'Content-Type': 'application/json', Origin: ORIGIN });
  check('own Origin is allowed', sameOrigin.status === 200, sameOrigin.status);
  for (const owner of ['Console', 'Company.Inc', 'COM10', 'A. Kumar']) {
    const r = await postSave({ ...SAMPLE, owner, format: 'pdf' });
    check(`similar valid owner is allowed: ${owner}`, r.status === 200, r.text);
    fs.rmSync(path.join(BASE, owner), { recursive: true, force: true });
  }
  const notJson = await postSave(payload, { 'Content-Type': 'text/plain' });
  check('non-JSON content type gets 415', notJson.status === 415, notJson.status);

  const after = fs.readdirSync(BASE).filter((d) => d !== 'Same Origin').sort().join();
  check('rejected requests created no folders', before === after, after);
  const outside = fs.readdirSync(TEMP_ROOT).filter((f) => f !== 'base');
  check('nothing was written outside the base folder', outside.length === 0, outside.join());
}

async function main() {
  const child = spawn(process.execPath, ['server.js'], {
    cwd: ROOT,
    env: { ...process.env, PORT: String(PORT), TABLE_SAVER_BASE: BASE },
    stdio: ['ignore', 'ignore', 'pipe'],
  });
  let stderr = '';
  child.stderr.on('data', (c) => { stderr += c; });

  try {
    if (!(await waitForServer(child))) {
      console.error(`FAIL  server did not start on port ${PORT}\n${stderr}`);
      failures++;
      return;
    }
    const health = JSON.parse((await request('GET', '/health')).text);
    check('/health reports the temp base folder', health.ok === true && health.baseDir === BASE, health.baseDir);

    await testManifest();
    await testStatic();
    await testMalformedRequests();
    await testSaves();
    await testCopies();
    await testUnwritableFolder();
    await testRejections();
    await testRelativeBaseDir();
    check('server stayed up', child.exitCode === null, stderr);
  } finally {
    await stopServer(child);
    // Windows can hold files briefly after a process exits, so allow retries.
    fs.rmSync(TEMP_ROOT, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 });
  }
}

main()
  .catch((e) => { failures++; console.error('FAIL  smoke test crashed:', e); })
  .finally(() => {
    console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll checks passed');
    process.exit(failures ? 1 : 0);
  });
