---
name: deploy-changes
description: Make a Table Saver change take effect on Windows, Ubuntu or macOS, including restarting the server, reloading the task pane, re-sideloading the manifest, and installing the startup service. Use after editing the project, or when Excel is still showing old behaviour.
---

# Deploy changes locally

Run the `verify-server` skill first. Then do only what the changed files require:

| Changed | Needed |
|---|---|
| `public/*` | Close and reopen the task pane |
| `server.js`, `config.json` | Restart the server |
| `package.json` dependencies | `npm install`, then restart the server |
| `manifest.xml` | Bump the version, sideload again, restart Excel or reload the page |
| `service/*` | Reinstall the service |

Use the commands for the system you're on. Check `process.platform` or `uname` if unsure.

## Restart the server

| How it's running | Windows | Ubuntu | macOS |
|---|---|---|---|
| Startup service | `schtasks /End /TN TableSaver`, then `schtasks /Run /TN TableSaver` | `systemctl --user restart table-saver` | `launchctl kickstart -k gui/$(id -u)/com.tablesaver` |
| pm2 | `pm2 restart table-saver` | same | same |
| Terminal | Ctrl+C, then `npm start` | same | same |

To find out how it's running:

| | Windows | Ubuntu | macOS |
|---|---|---|---|
| Service state | `schtasks /Query /TN TableSaver` | `systemctl --user status table-saver` | `launchctl list \| grep tablesaver` |
| What holds the port | `netstat -ano \| findstr :3000` | `ss -ltnp 'sport = :3000'` | `lsof -nP -iTCP:3000 -sTCP:LISTEN` |

Then confirm it's up. `-k` is needed because curl may not use the store holding the dev certificate. In Windows PowerShell, type `curl.exe`, because `curl` can be an alias for something else.

```bash
curl -sk https://localhost:3000/health
```

## Manifest changed

1. Increase `<Version>` in `manifest.xml`, for example `1.0.0.0` to `1.0.1.0`. Don't change `<Id>`, because Excel would treat it as a different add-in.
2. Sideload it again:
   - **Windows:** replace the copy in the shared catalog folder, then restart Excel.
   - **Ubuntu (Excel on the web):** upload it again under Home → Add-ins → More Settings → Upload My Add-in.
   - **macOS:** copy it to `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/`, then restart Excel.
3. If Excel still shows the old ribbon or pane, clear its cache.

## Clearing the Office cache

Get the user's confirmation first, because this affects **all** add-ins, not just this one.

- **Windows:** close Excel and run `npx office-addin-cache clear`. Or use Trust Center → Trusted Add-in Catalogs → "Next time Office starts, clear all previously-started web add-ins cache", then restart Excel.
- **Ubuntu:** clear the browser's cache. This also removes the uploaded manifest, so upload it again.
- **macOS:** try the web cache first: in the open task pane, use the personality menu and choose **Clear Web Cache**. For a full clear, close Excel and run `npx office-addin-cache clear`. A full clear also empties the `wef` folder and removes every sideloaded add-in, so copy `manifest.xml` back afterwards.

Don't delete individual manifest files to force a refresh, because Microsoft warns this can stop add-ins from loading.

Details: <https://learn.microsoft.com/en-us/office/dev/add-ins/testing/clear-cache>

## Install the startup service

The files in `service/` are templates. Don't overwrite them with machine-specific values. Write the filled-in copy to the install location, and get the node path from `which node` (`where node` on Windows). If that path comes from a version manager such as nvm, warn the user that it can break when the Node version changes.

### Windows

Use the `schtasks` command from the README with the real node and project paths. Keep the quotes around both paths, because either can contain spaces. If it reports "Access is denied", ask the user to run it from an administrator Command Prompt.

Task Scheduler keeps no server logs. To see a startup error, stop the task and run `npm start` in a terminal.

### Ubuntu

1. In `~/.config/systemd/user/table-saver.service`, set `ExecStart` and `WorkingDirectory`. Keep `%h` for the home folder.
2. Load it:
   ```bash
   systemctl --user daemon-reload
   systemctl --user enable --now table-saver
   ```
3. If `/health` fails, read `journalctl --user -u table-saver -n 30`.

The service runs while the user is logged in, which is all the add-in needs, because Excel runs in their browser. Don't enable lingering or install a system-wide unit unless asked.

### macOS

1. In `~/Library/LaunchAgents/com.tablesaver.plist`, replace the node path and both `/Users/YOUR_USER/table-saver-addin` paths.
2. If the project is in `~/Downloads`, `~/Documents` or `~/Desktop`, warn the user that macOS privacy controls may stop a background process reading it. Moving the project elsewhere is the simplest fix.
3. Load it. To reinstall, unload the old one first:
   ```bash
   launchctl unload ~/Library/LaunchAgents/com.tablesaver.plist 2>/dev/null
   launchctl load ~/Library/LaunchAgents/com.tablesaver.plist
   ```
4. If `/health` fails, read `/tmp/tablesaver.err.log`.

## Tell the user

Say exactly which of the steps above you ran and which they still need to do in Excel. You can't restart Excel or check the ribbon from here unless the user asks you to and the tools allow it.
