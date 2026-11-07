---
name: add-export-format
description: Add a new Table Saver export format such as CSV, or change how the PDF or XLSX output is written. Use for new save buttons, new file types, or changes to writePdf and writeXlsx.
---

# Add or change an export format

The format string is also the file extension: `<baseDir>/<owner>/<fileName>.<format>`. If that name is taken, the save becomes `<fileName> (1).<format>`, and so on.

## Adding a format

1. **`server.js`: add a writer** with the same signature as the others: `writeCsv(file, tableName, owner, rows)`. It must return a promise that resolves only after the file is fully written. For streams, resolve on the output stream's `finish` event, as `writePdf` does.
2. **`server.js`: update `handleSave`** in both places:
   - Add the format to the validation check and its error message.
   - Change the dispatch. It currently ends in a bare `else` that calls `writeXlsx`, so if you only update the validation, the new format writes XLSX content under the new extension. Give each format its own branch.
3. **`server.js`: update the comments** at the top of the file that list the formats, and the startup log line (`<pdf|xlsx>`).
4. **`public/taskpane.html`: add a button** with the `secondary` class.
5. **`public/taskpane.js`: bind it** in `Office.onReady`, for example `save('csv')`.
6. **Update the wording** that names the formats:
   - `README.md`: intro and usage line
   - `manifest.xml`: `Description` and `Button.Tip`
   - `package.json`: `description`
7. **`smoke-test.js`: add checks** in `testSaves` in `.claude/skills/verify-server/scripts/`. Confirm that the new file has the right contents, not just that it exists.

## Rules for writers

- Keep the modes: folders 755, files 644. `handleSave` already calls `ensureDir` and `finalizeFile`, so writers don't need to.
- Handle empty values with `cell()`. Sheet values can be null, and numbers should stay numbers where the format supports them.
- Rows can have different lengths, so don't assume they're all as long as the header.
- Don't build the output path inside a writer. `handleSave` owns the path and its safety checks.
- Open the given file for overwriting, not with an exclusive flag such as `wx`. `reserveFile` has already created it empty to claim the name.
- Reject the promise on any failure, so `handleSave` can remove the claimed file.
- For text formats opened in Excel, such as CSV, escape cells starting with `=`, `+`, `-` or `@` to prevent formula injection, and quote commas, quotes and newlines.

## Output layouts to keep consistent

- **PDF:** title, then `Owner: <owner>`, then a grid with a shaded bold header. Rows start on a new page when they don't fit.
- **XLSX:** title in A1, `Owner` in A2 and the owner's name in B2, then the table from row 3 with a bold header. The smoke test checks these cells, so update it if you change them on purpose.

## New dependencies

Use `npm install <pkg>` so `package-lock.json` stays correct. Check first whether `exceljs` already does the job: it can write CSV.

## Verify

1. Restart the server, which is needed for any `server.js` change (see `deploy-changes`).
2. Run the `verify-server` skill.
3. Open a generated sample, or ask the user to. Reload the task pane to see the new button.
