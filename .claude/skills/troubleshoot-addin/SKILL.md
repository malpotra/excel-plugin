---
name: troubleshoot-addin
description: Diagnose Table Saver failures on Windows, Ubuntu or macOS, including "Cannot reach the local server", a blank task pane, a missing ribbon button, certificate errors, save errors, and the server or service not starting. Use when the add-in or server misbehaves.
---

# Troubleshoot Table Saver

## Collect the facts first

These commands only read state, so run them before changing anything. `curl -sk https://localhost:3000/health` works everywhere and shows whether the server is up and which folder it saves to. In Windows PowerShell, type `curl.exe`.

**Windows (PowerShell):**

```powershell
netstat -ano | findstr :3000
schtasks /Query /TN TableSaver
(New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 "$HOME\.office-addin-dev-certs\localhost.crt").NotAfter
```

**Ubuntu:**

```bash
ss -ltnp 'sport = :3000'
systemctl --user status table-saver
journalctl --user -u table-saver -n 30
openssl x509 -in ~/.office-addin-dev-certs/localhost.crt -noout -enddate
```

**macOS:**

```bash
lsof -nP -iTCP:3000 -sTCP:LISTEN
launchctl list | grep tablesaver                          # 2nd column = last exit code
tail -n 30 /tmp/tablesaver.err.log /tmp/tablesaver.log
openssl x509 -in ~/.office-addin-dev-certs/localhost.crt -noout -enddate
ls -l ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
```

## Symptoms on any system

| Symptom | Likely cause | Fix |
|---|---|---|
| Pane says `Cannot reach the local server` | Server is down, or the browser rejected the certificate. The pane shows this for any network-level failure | Check `/health`, then the certificate |
| Pane or ribbon worked before, now blank or failing | **Expired certificate.** Dev certs last 30 days by default | See "Certificates" |
| `ENOENT ... localhost.key` at startup | Certificates were never generated | See "Certificates" |
| `EADDRINUSE` at startup | Another process holds the port, often a second copy of the server | Identify it with the commands above. If it's an old Table Saver process, restart that one. Otherwise see `change-server-settings` |
| JSON error from `config.json` at startup | Often single backslashes in a Windows path | Use `D:/Reports` or `D:\\Reports` |
| `fileName ... is a reserved name` at startup | `fileName` in `config.json` is a name such as `con` | Choose another name |
| Files save in an unexpected folder | Wrong `baseDir`, or a literal `%VAR%` or `$VAR`, which isn't expanded and so becomes a folder inside the project | Check `baseDir` in `/health`, then see `change-server-settings` |
| No **Save Table** button | Manifest isn't sideloaded, or Excel hasn't restarted | See "Manifest changed" in `deploy-changes` |
| Old behaviour after an edit | Server not restarted, or a cached pane | See the table at the top of `deploy-changes` |
| `Owner cell (B2) is empty` | Owner cell is blank, or the sheet layout differs | Check the sheet, then see `update-sheet-layout` |
| `Owner is empty or not a valid folder name` | `safeOwner` allows ASCII letters, digits, `_`, space, `.` and `-`, and at most 64 characters. It must start with a letter, digit or `_`, can't end in `.`, and can't be a Windows reserved name such as `CON` | Correct the sheet value. Don't loosen the check without asking |
| `Forbidden origin` | The request came from somewhere other than `https://localhost:<port>`, such as a manifest pointing at a different host or port | Make the manifest's URLs match the server |
| `Could not save to ...` | The output folder isn't writable | See your system below. Don't run the server as root or administrator to get around it |
| Many `table (N).*` files in a folder | Expected: every save keeps a new copy, and nothing deletes old ones | Ask the user before deleting any |
| Newest save doesn't have the highest number | Expected after deleting a copy: `reserveFile` reuses the first free number | Check the name shown in the pane, or the file's modified time |
| Dates appear as numbers | Task pane reads raw `values` | See "Cell values" in `update-sheet-layout` |

## Windows

| Symptom | Likely cause | Fix |
|---|---|---|
| `Could not save to ...` | Antivirus or Controlled Folder Access can block node from writing to protected folders such as Documents | Choose another `baseDir`, or ask the user to allow node |
| `abcd` and `ABCD` files are numbered together | Windows folder names aren't case-sensitive | Use one spelling in the sheets |
| Server stops after a while | The console window opened at login was closed | Run the task again, and leave the window open or minimized |
| No Shared Folder option in My Add-ins | The catalog isn't registered, or Excel hasn't restarted | Recheck the Trusted Add-in Catalogs steps in the README |

## Ubuntu

| Symptom | Likely cause | Fix |
|---|---|---|
| Certificate warning, or `Cannot reach the local server`, with `/health` fine in curl | Chrome reads its own certificate database, not the system store | Run the `certutil` step in the README, then restart Chrome |
| Same symptom in Firefox | Firefox keeps a separate list | Import `ca.crt` under Authorities (see the README) |
| Add-in gone from Excel on the web | The browser stores the uploaded manifest, and its cache was cleared | Upload it again |
| Pane can't reach the server in the browser, but `/health` opens in a tab | The browser blocked the Excel page from reaching localhost | Allow local network access for the Excel site in the browser's site settings |
| Service fails with `status=203/EXEC` | `ExecStart` has the wrong node path | Use the output of `which node` |
| `Could not save to ...` | The folder is owned by root, often after running the server with `sudo` | Check with `ls -ld <baseDir> <baseDir>/<owner>`, and ask the user before changing ownership |
| `node` errors at startup | The `apt` version may be too old | Check `node -v`, and install Node 20 or newer |

## macOS

| Symptom | Likely cause | Fix |
|---|---|---|
| Service exits immediately | Wrong node path or leftover `YOUR_USER` in the plist. On Apple Silicon, node is often at `/opt/homebrew/bin/node` | Read the error log, then reinstall (see `deploy-changes`) |
| `EPERM` or `Operation not permitted` under launchd only | Privacy controls block background access to `~/Downloads`, `~/Documents` or `~/Desktop` | Move the project, or the output folder, somewhere else |

## Certificates

Generate or renew them with:

```bash
npx office-addin-dev-certs install --days 365
```

- Without `--days`, the certificate lasts 30 days.
- This asks for admin approval (`sudo` on Ubuntu, a security prompt on Windows, the keychain on macOS), so **the user must run it or approve it**. Don't try to work around the prompt.
- If it reports that valid certificates already exist, run `npx office-addin-dev-certs uninstall` first, with the user's approval.
- On Ubuntu, repeat the `certutil` step from the README, because the CA may have changed. If an old entry is in the way, remove it first with `certutil -d sql:$HOME/.pki/nssdb -D -n office-addin-dev-certs`.
- Restart the server afterwards. It only reads the certificate files at startup.

## Look at the task pane itself

For errors that don't appear in the server log, use the pane's developer tools:

- **Windows:** the tools depend on the Office version. See Microsoft's "Debug add-ins on Windows" documentation.
- **Ubuntu:** use the browser's own developer tools (F12), and select the add-in's frame in the console.
- **macOS:** enable them with `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`, restart Excel, then right-click the pane and choose **Inspect Element**.

## After fixing

Run the `verify-server` skill. If the fix touched how the add-in loads, ask the user to confirm it in Excel too.

If you find a new recurring failure, add a row to the matching table.
