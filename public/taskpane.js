/* global Office, Excel */
'use strict';

// Sheet layout (edit here if yours differs; indexes are 0-based):
//   A1 = table name | B2 = owner | row 3 = headers | rows 4+ = data
const TABLE_NAME_CELL = { row: 0, col: 0 };
const OWNER_CELL = { row: 1, col: 1 };
const TABLE_START_ROW = 2;

Office.onReady(() => {
  document.getElementById('btn-pdf').onclick = () => save('pdf');
  document.getElementById('btn-xlsx').onclick = () => save('xlsx');
});

function show(msg, ok) {
  const el = document.getElementById('status');
  el.textContent = msg;
  el.className = ok ? 'ok' : 'err';
}

function setBusy(busy) {
  document.querySelectorAll('button').forEach((b) => { b.disabled = busy; });
}

async function readSheet() {
  return Excel.run(async (ctx) => {
    const used = ctx.workbook.worksheets.getActiveWorksheet().getUsedRange();
    used.load(['rowIndex', 'columnIndex', 'rowCount', 'columnCount']);
    await ctx.sync();
    // Re-anchor at A1 so the fixed layout holds even if the used range starts lower/righter.
    const all = ctx.workbook.worksheets.getActiveWorksheet().getRangeByIndexes(
      0, 0, used.rowIndex + used.rowCount, used.columnIndex + used.columnCount);
    all.load('values');
    await ctx.sync();
    return all.values;
  });
}

function trimRows(rows) {
  const isEmpty = (v) => v === '' || v === null || v === undefined;
  const kept = rows.filter((r) => r.some((v) => !isEmpty(v)));
  const width = Math.max(0, ...kept.map((r) => {
    let w = r.length;
    while (w > 0 && isEmpty(r[w - 1])) w--;
    return w;
  }));
  return kept.map((r) => r.slice(0, width));
}

async function save(format) {
  setBusy(true);
  try {
    const values = await readSheet();
    const tableName = String((values[TABLE_NAME_CELL.row] || [])[TABLE_NAME_CELL.col] || '').trim();
    const owner = String((values[OWNER_CELL.row] || [])[OWNER_CELL.col] || '').trim();
    const rows = trimRows(values.slice(TABLE_START_ROW));

    if (!owner) throw new Error('Owner cell (B2) is empty.');
    if (rows.length === 0) throw new Error('No table found from row 3 down.');

    const resp = await fetch('/save', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ owner, tableName, rows, format }),
    });
    const data = await resp.json().catch(() => ({}));
    if (!resp.ok) throw new Error(data.error || `Server returned ${resp.status}`);
    show(`Saved: ${data.path}`, true);
  } catch (e) {
    const msg = e instanceof TypeError ? 'Cannot reach the local server. Is it running?' : e.message;
    show(msg, false);
  } finally {
    setBusy(false);
  }
}
