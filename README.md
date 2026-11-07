# Table Saver – Excel add-in + local server

Reads the active sheet and saves the table as **PDF or Excel** to
`<baseDir>/<Owner>/table.pdf` (or `.xlsx`).

Saving never overwrites an existing file. If `table.pdf` already exists, the new save becomes
`table (1).pdf`, then `table (2).pdf`, and so on, using the first free number. The task pane
shows the name each save got. Copies are never deleted automatically, so remove old ones
yourself when you no longer need them.

Excel add-ins are sandboxed and can't write to disk, so a small Node server on
`https://localhost:3000` does the file writing. It runs as **your user**. On Ubuntu and
macOS, new folders get mode 755 and files 644. On Windows, they inherit the permissions
of the folder they're created in.

## Supported systems

| System | Server | Excel |
|---|---|---|
| Windows | Node.js | Desktop Excel, or Excel on the web |
| Ubuntu | Node.js | Excel on the web, in a browser on the same computer (there's no desktop Excel for Linux) |
| macOS | Node.js | Desktop Excel, or Excel on the web |

The server and Excel must run on the same computer, because the server only accepts
connections from that machine.

## Sheet layout

|   | A          | B      |
|---|------------|--------|
| 1 | Table name |        |
| 2 | Owner      | ABCD   |
| 3 | SNO        | Column |
| 4 | 1          | A      |
| 5 | 2          | B      |

Change the cell positions at the top of `public/taskpane.js` if yours differs.

## Setup (once)

You need Node.js 20 or newer. On Ubuntu, the version from `apt` may be older, so check with
`node -v`.

The commands are the same on every system:

```bash
cd table-saver-addin
npm install
npx office-addin-dev-certs install --days 365   # trusted localhost HTTPS cert (asks for admin approval)
npm start
```

Without `--days`, the certificate expires after 30 days.

**Ubuntu only:** the command above adds the certificate to the system store, but Chrome reads
its own certificate database. Add it there too, then restart Chrome:

```bash
sudo apt install libnss3-tools
mkdir -p ~/.pki/nssdb
certutil -d sql:$HOME/.pki/nssdb -A -t "C,," -n office-addin-dev-certs -i ~/.office-addin-dev-certs/ca.crt
```

Firefox keeps its own list too. To use it, import `~/.office-addin-dev-certs/ca.crt` under
Settings → Privacy & Security → View Certificates → Authorities.

Check `https://localhost:3000/health` in a browser – it should return `{"ok":true,...}`.

## Choosing the output folder

Edit `baseDir` in `config.json`, or set the env var `TABLE_SAVER_BASE`. The default is `~/home`,
which gives:

| System | Example output file |
|---|---|
| Windows | `C:\Users\you\home\ABCD\table.pdf` |
| Ubuntu | `/home/you/home/ABCD/table.pdf` |
| macOS | `/Users/you/home/ABCD/table.pdf` |

`~` works on every system, so the default `config.json` can be shared between them. A full path
only works on its own system. On Windows, write it with forward slashes or doubled backslashes,
because a single `\` is invalid in JSON:

```json
{ "baseDir": "D:/Reports" }
```

Variables such as `%USERPROFILE%` and `$HOME` aren't expanded. A relative path resolves against
the project folder.

## Load the add-in into Excel (sideload)

**Windows:**
1. Put `manifest.xml` in a folder and share it (right-click → Properties → Sharing), noting the network path `\\PC\folder`.
2. Excel → File → Options → Trust Center → Trust Center Settings → **Trusted Add-in Catalogs**,
   add that path, tick **Show in Menu**, OK, restart Excel.
3. Insert → My Add-ins → **Shared Folder** → Table Saver → Add.

**Ubuntu (Excel on the web):**
1. Open [office.com](https://www.office.com/) in Chrome and open a workbook in Excel.
2. Home → **Add-ins** → **More Settings** → **Upload My Add-in**.
3. Browse to `manifest.xml` and select **Upload**.

The browser stores this upload, so repeat it if you clear the browser's cache or switch
browsers. If the browser asks for permission to access devices on your local network, allow it,
because the add-in can't reach the server otherwise.

**macOS:** copy `manifest.xml` to
`~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/` (create `wef` if missing),
restart Excel, then **Home → Add-ins → Table Saver** (or Insert → My Add-ins → Developer Add-ins).

A **Save Table** button then appears on the Home tab. Click it, then **Save as PDF** / **Save as Excel**.

## Run on startup

**Windows (Task Scheduler, runs at login as you):** run this in a Command Prompt, opened as
administrator if you get "Access is denied":

```bat
schtasks /Create /TN "TableSaver" /SC ONLOGON /RL LIMITED /TR "\"C:\Program Files\nodejs\node.exe\" \"C:\path\to\table-saver-addin\server.js\""
```

The server may run in a console window at login. Closing that window stops the server.

**Ubuntu (systemd, runs at login as you):** edit `service/table-saver.service` (node + project
paths), then

```bash
mkdir -p ~/.config/systemd/user
cp service/table-saver.service ~/.config/systemd/user/
systemctl --user daemon-reload
systemctl --user enable --now table-saver
```

**macOS (launchd):** edit `service/com.tablesaver.plist` (user name + paths), then

```bash
cp service/com.tablesaver.plist ~/Library/LaunchAgents/
launchctl load ~/Library/LaunchAgents/com.tablesaver.plist
```

(`pm2` works too: `pm2 start server.js --name table-saver && pm2 save`.)

## Maintaining

Claude Code picks up the project skills in `.claude/skills/` automatically. Run one with `/<name>`:

| Skill | Use it to |
|---|---|
| `verify-server` | Run the smoke test: `node .claude/skills/verify-server/scripts/smoke-test.js` |
| `update-sheet-layout` | Change which cells hold the table name, owner and rows |
| `add-export-format` | Add a file type or change the PDF/XLSX output |
| `change-server-settings` | Change the port, output folder or file name |
| `deploy-changes` | Restart, re-sideload or install the service after a change |
| `troubleshoot-addin` | Diagnose certificate, connection, service and save errors |

## Security notes

- Listens on `127.0.0.1` only and rejects requests from any origin other than itself.
- The owner value becomes a folder name, so only letters, numbers, space, `.`, `_`, `-` are accepted
  (no `/`, `..`, etc.). A trailing dot and names Windows reserves, such as `CON` and `NUL`, are
  rejected on every system, so a sheet behaves the same wherever it's saved.

## Differences between systems

- **Upper and lower case:** on Windows, owners `abcd` and `ABCD` share one folder, so their files
  are numbered together. On Ubuntu they're separate folders.
