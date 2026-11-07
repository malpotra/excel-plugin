---
name: verify-server
description: Run the Table Saver end-to-end smoke test and report what it does and doesn't prove. Use after any change to server.js, config.json, manifest.xml or package dependencies, or when asked to test or verify this project.
---

# Verify the Table Saver server

The project has no test framework. This smoke test is its regression check:

```bash
node .claude/skills/verify-server/scripts/smoke-test.js
```

It starts its own server on port 3987 (override with `SMOKE_PORT`), saves into a temp folder, then cleans up. It's safe to run while the real server is running on 3000, and it never touches real output files.

## What it covers

- `/health`, the task pane files, and every URL in `manifest.xml`
- The manifest port matching the default port in `server.js`
- PDF and XLSX saves: file contents, folder mode 755, file mode 644, and a multi-page PDF
- Rejections: bad owner names (including names that are only invalid on Windows), unknown formats, bad rows, bad JSON, a foreign `Origin`, a wrong content type, and static path traversal
- Malformed URLs and JSON bodies that aren't objects, which must get a 400 and leave the server running
- Numbered copies: earlier files stay unchanged, simultaneous saves get separate files, and read-only files are kept
- A clear error when the output folder can't be written to, and a relative output folder that doesn't depend on the working directory

## What it doesn't cover

The script is written to run on Windows, Ubuntu and macOS, but it only tests the system it runs on. A pass on one system says nothing about OS-specific behaviour on another, so say which system you ran it on. Saving while an earlier copy is open and locked by another program isn't tested. Folder and file modes are skipped on Windows.

The test never runs `public/taskpane.js` or Excel. After changing either of these, ask the user to check them in Excel, and say that this step wasn't verified automatically:

- Reading cells from the sheet (`readSheet`, `trimRows`, the layout constants)
- The ribbon button and task pane loading

It only checks that a PDF is complete, not how it looks. After changing `writePdf`, render a sample to PNG and look at it, or ask the user to open one.

## Handling results

- **`server did not start`**: read the stderr it prints, then use the `troubleshoot-addin` skill.
- **A failing check**: fix the code, not the check, unless the behaviour was changed on purpose. If it was, update the check in the same change and tell the user.
- **New behaviour**: add a check to `scripts/smoke-test.js` when you add an endpoint, format or validation rule.
