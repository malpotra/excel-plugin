---
name: update-sheet-layout
description: Change where Table Saver reads the table name, owner or table rows in the Excel sheet, or how cell values are read and trimmed. Use when the user's sheet layout differs, or when moving the owner, title or header cells.
---

# Update the sheet layout

The layout is only known to the task pane. The server just receives `{ owner, tableName, rows, format }`, so `server.js` normally doesn't change.

## Places to keep in sync

The layout is described in several places, and only the first one affects behaviour:

| File | What to update |
|---|---|
| `public/taskpane.js` | `TABLE_NAME_CELL`, `OWNER_CELL`, `TABLE_START_ROW` and the comment above them |
| `public/taskpane.js` | Error messages in `save()` that name cells: `Owner cell (B2)`, `from row 3 down` |
| `public/taskpane.html` | The `hint` paragraph |
| `README.md` | The "Sheet layout" table |

## Rules to preserve

- **Indexes are 0-based**: B2 is `{ row: 1, col: 1 }`, and row 3 is `TABLE_START_ROW = 2`.
- **`readSheet` re-anchors at A1 on purpose.** Keep this, or the fixed cell positions break when the used range doesn't start at A1.
- **The first row from `TABLE_START_ROW` is the header.** Both writers in `server.js` style `rows[0]` as the header.
- **Rows run to the bottom of the used range.** Anything below the table, such as notes or totals, gets exported too.
- **Columns are not limited.** `trimRows` removes empty rows, including blank rows inside the table, and trailing empty columns. It keeps everything else from column A onward.
- **Owner is validated twice.** The pane checks that it isn't empty, and `safeOwner` in `server.js` checks the allowed characters. Don't loosen `safeOwner` when changing the layout, because the owner becomes a folder name.

## Cell values

`readSheet` loads `values`, which are raw values rather than what the cell displays. A date exports as a serial number such as `45992`, and a percentage as `0.5`. Loading `text` instead gives the displayed strings, but numbers then become text in the XLSX output. Confirm which the user wants before switching.

## Verify

The smoke test doesn't run the task pane, so this change needs a check in Excel:

1. Run the `verify-server` skill to confirm nothing else broke.
2. Close and reopen the task pane. `public/` files need no server restart.
3. Ask the user to save from a sheet in the new layout and confirm the output folder, title and rows.
4. Say clearly that the layout change itself was only verified if the user confirmed step 3.
