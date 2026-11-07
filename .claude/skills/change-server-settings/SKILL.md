---
name: change-server-settings
description: Change the Table Saver port, output folder or output file name, and update everything that depends on them. Use when port 3000 conflicts, when files should save elsewhere, or when renaming table.pdf or table.xlsx.
---

# Change server settings

## Output folder

Precedence is the `TABLE_SAVER_BASE` env var, then `baseDir` in `config.json`, then `~/home`.

- Prefer editing `config.json`. A background service doesn't inherit shell env vars, so one set in a shell profile silently has no effect under launchd or systemd, and a new Windows variable needs a fresh login.
- `~` is expanded on every system, so `~/...` is the only form that works unchanged on Windows, Ubuntu and macOS. A relative path resolves against the project folder.
- A full path only works on its own system. On another system, `D:/Reports` becomes a folder inside the project. On Windows, write it as `D:/Reports` or `D:\\Reports`, because a single backslash is invalid JSON.
- `%USERPROFILE%` and `$HOME` aren't expanded and become literal folder names.
- Folders are created automatically. Existing files are not moved, so tell the user their old files stay where they are.
- Restart the server, then confirm the new folder in the `baseDir` value from `/health`.

## Output file name

Set `fileName` in `config.json`. Characters other than letters, digits, `_`, `.` and `-` become `_`, and the extension is added automatically. The server refuses to start if the name is reserved on Windows, such as `con` or `nul`.

Saving never overwrites. `reserveFile` in `server.js` picks the first free name, so `report.pdf` is followed by `report (1).pdf`. After a rename, the numbering starts again from the new name, and files under the old name stay where they are.

Also update the places that mention `table.*`:

- `public/taskpane.html`: hint text, including the numbered example
- `README.md`: intro and output folder section
- `package.json`: `description`

## Port

The port must change in two places. Change both, or the add-in stops loading:

1. `server.js`: the default in `process.env.PORT || 3000`
2. `manifest.xml`: **every** `https://localhost:3000` URL (replace all of them)

Then update `README.md`, which mentions the port in its intro and setup sections.

Because the manifest changed, the user must sideload it again. Follow "Manifest changed" in `deploy-changes`.

Prefer changing the default over setting a `PORT` env var. The manifest can't read env vars, so a server running on another port won't match it. The smoke test's manifest check also compares against the default in `server.js`.

The allowed `Origin` is derived from the port, so no separate change is needed there.

## Don't change without asking

These protect the machine, because anything running in a browser can reach localhost. If a request needs one of them changed, explain the risk and get explicit confirmation first:

- `HOST = '127.0.0.1'`: changing it exposes file writing to the network
- The `Origin` check and the JSON content-type check in `handleSave`
- The `safeOwner` character rules

## Verify

Run the `verify-server` skill, then restart the real server and open `/health`.
