# Website E2E Test Automation (PowerShell)

## Project layout
```
website-e2e-tester/
├── scripts/
│   └── Run-WebsiteE2ETests.ps1
├── config/
│   ├── Site-Config.csv         ← fill in with your real 6-7 sites
│   └── Site-Config-Demo.csv    ← public sandbox sites, for trial runs
├── TestRun_*/                  ← auto-created per run, gitignored
└── .gitignore
```
Everything here runs locally against your own Windows machine — nothing runs in a cloud sandbox, so test it on your PC before pushing to GitHub. Credentials are never written to disk — the script prompts for username/password fresh every run, for every site (see "Credentials" below).

## One-time setup
```powershell
Install-Module Selenium -Scope CurrentUser -Force
Install-Module ImportExcel -Scope CurrentUser -Force
```
Selenium's module auto-manages a bundled `chromedriver.exe` — but see Troubleshooting below, as the bundled version can be years out of date and may need replacing to match your installed Chrome. **Google Chrome** (not Edge) is the reliably-supported browser for this module — Edge support in the `Selenium` module is limited to the legacy pre-Chromium driver.
Word must already be installed for the `.docx` report step (it drives Word via COM, it doesn't need internet).

**Before your first run**, open `scripts\Run-WebsiteE2ETests.ps1` and set these two lines near the top of the site loop to match your machine:
```powershell
$chromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"   # your Chrome.exe location
$driverDir  = "C:\Code\website-e2e-tester\drivers"                       # folder containing a chromedriver.exe matching your Chrome version
```
See Troubleshooting below for how to find your Chrome path and download a matching `chromedriver.exe` if you don't have one yet.

## Configure your sites
Edit `Site-Config.csv`:
- One `Login` row per site: the login URL and CSS selectors for the username field, password field, submit button, and an element that only appears after a successful login (`PostLoginCheckSelector`) — this is how the script tells Success vs Failure at login.
- Optional on the `Login` row: `LogoutMenuSelector` (a menu/hamburger icon to open first, if logout is behind one) and `LogoutSelector` (the actual logout link/button). Leave both blank to skip logout testing for that site entirely.
- One `Page` row per page you want tested: the URL, and optionally the CSS selector + expected text for the header and subheader. Leave `ExpectedHeader`/`HeaderSelector` blank if you just want a screenshot with no text check.

To find CSS selectors: open the page in Edge/Chrome, right-click the element → Inspect, right-click the highlighted HTML → Copy → Copy selector.

## Credentials
Every run, for every site, the script prompts for username and password right in the PowerShell console (password input is masked). Nothing is saved to disk anywhere — no file, no cache — so you'll be asked again next time, even for the same site. This is deliberate: no credential file to accidentally commit, leave behind, or have picked up by anything else on the machine.

## Run it
From the repo root:
```powershell
.\scripts\Run-WebsiteE2ETests.ps1 -ConfigPath .\config\Site-Config-Demo.csv -Headless:$false
```
Once the demo run looks right, switch to your real config:
```powershell
.\scripts\Run-WebsiteE2ETests.ps1 -ConfigPath .\config\Site-Config.csv
```
Add `-Headless:$false` any time you want to watch the browser while debugging selectors. `TestRun_*\` folders are created next to the repo root automatically and are already gitignored.

## Output
Each run creates a timestamped folder containing:
- `TestResults.xlsx` — Site, Page, URL, Status (Success/Failure/Error), Timestamp, Error message. Rows include LOGIN, one per configured Page, and LOGOUT (if configured).
- `TestReport.docx` — one section per site, one sub-section per row, with the screenshot and any mismatch note
- `Screenshots\` — one PNG per step: pre-login page, post-login landing page, each tested page (headers/subheaders visible in-frame), and the post-logout page if logout is configured

## Status meanings
- **Success** — page loaded and header/subheader (if checked) matched
- **Failure** — page loaded fine, but expected text wasn't found (a real content/UI problem)
- **Error** — something broke the test itself (selector not found, timeout, site unreachable) — treat separately from Failure since it may just mean a selector needs updating, not that the site is broken

## Scheduling
Wrap the `.ps1` call in a Windows Task Scheduler job (Action: `powershell.exe -File "...\Run-WebsiteE2ETests.ps1"`) to run it daily/nightly unattended.

## Troubleshooting

| Symptom | Root cause | Fix |
|---|---|---|
| `running scripts is disabled on this system` (`UnauthorizedAccess` / `PSSecurityException`) | Windows blocks unsigned local scripts by default. Resets every new PowerShell window if set with `-Scope Process`. | Per-session: `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force`<br>Permanent (no admin needed): `Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned` |
| `The term '.\Run-WebsiteE2ETests.ps1' is not recognized...` | Wrong relative path — the script lives in `scripts\`, not the repo root. | Run `.\scripts\Run-WebsiteE2ETests.ps1 ...` from the repo root, or `.\Run-WebsiteE2ETests.ps1` from inside `scripts\`. Check with `Get-ChildItem -Recurse -Filter *.ps1`. |
| Script downloaded/unzipped from elsewhere won't run even after setting execution policy | Windows marks files from a zip/download as "blocked" (Zone.Identifier). | From the project root: `Get-ChildItem -Recurse | Unblock-File` |
| `Install-Module` fails / times out against PSGallery | PowerShell 5.1 defaults to an older TLS version PSGallery no longer accepts. | Run once per session before installing: `[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12` |
| Parser errors mentioning `Missing ')' `, `Unexpected token`, or garbled characters like `â€”` | A non-ASCII character (e.g. an em-dash `—`) got misread — PS 5.1 reads `.ps1` files without a BOM using the system ANSI codepage, not UTF-8. | Replace any non-ASCII punctuation with plain ASCII (e.g. `—` → `-`). Check for others with: `Get-ChildItem *.ps1 -Recurse \| Select-String -Pattern "[^\x00-\x7F]"` |
| `Exception setting "SaveAs": Cannot convert ... value of type "psobject" to type "Object"` | Known PowerShell/COM marshalling quirk — Word's `SaveAs` needs plain typed values, not PowerShell's boxed variable wrapper. | Cast explicitly before wrapping in `[ref]`: `$doc.SaveAs([ref]([string]$docxPath), [ref]([int]16))` |
| Script never returns / doesn't respond to `Ctrl+C` | Word (or the browser) is blocked on a hidden dialog box — COM calls are synchronous and can't be interrupted from the console. | In a **new** window: `Get-Process WINWORD -ErrorAction SilentlyContinue \| Stop-Process -Force` and `Get-Process msedge,msedgedriver -ErrorAction SilentlyContinue \| Stop-Process -Force`. Prevent it going forward by adding `$word.DisplayAlerts = 0` right after `$word.Visible = $false`. If it still hangs, temporarily set `$word.Visible = $true` to see and dismiss the dialog manually, so we know which one is firing. |
| `Retrieving the COM class factory ... failed` for `Word.Application` | Microsoft Word isn't installed (or is installed as a Microsoft 365 "click-to-run" app not registered for COM) on this machine. | Install/repair desktop Word, or run the script on a machine that has it — the `.docx` step specifically needs local Word. |
| `The process cannot access the file '...TestReport.docx' because it is being used by another process` | A previous run's Word document is still open (crashed run left Word running, or you have the file open yourself). | Close Word / run `Get-Process WINWORD \| Stop-Process -Force`, then re-run. |
| `Import-Csv : Could not find a part of the path ...` followed by a `Done.` with empty/near-empty reports | Config path was relative to the wrong folder (e.g. run from `scripts\` with a `.\config\...` path meant for the repo root), so `Import-Csv` failed as a non-fatal error and the rest of the script ran anyway with zero rows. | Use the right relative path for where you're standing (`.\scripts\Run-...ps1 -ConfigPath .\config\...` from repo root, or `..\config\...` from inside `scripts\`). The script now also hard-stops with a clear message if the config file is missing — add this right after the `Import-Csv` line if your copy doesn't have it yet: `if (-not (Test-Path $ConfigPath)) { Write-Error "Config file not found: $ConfigPath"; exit 1 }` |
| No screenshots are produced anywhere, even when Status shows Success | The script called `Save-SeScreenshot -Driver ... -Path ...`, but that cmdlet doesn't accept `-Driver` — screenshots need two steps in this module. | Everywhere the script calls `Save-SeScreenshot -Driver $driver -Path $x`, change it to: `Invoke-SeScreenshot -Driver $driver \| Save-SeScreenshot -Path $x` (3 places: login screenshot, page screenshot, and the one inside the page `catch` block). |
| `A parameter cannot be found that matches parameter name 'Arguments'` (or `'Headless'`) on `Start-SeEdge` | This module's `Start-SeEdge` only wraps the old, legacy (pre-Chromium) Edge driver, which doesn't accept Chromium-style flags at all. | Switch to Chrome instead, which fully supports both: replace `Start-SeEdge -Headless -Arguments @(...)` / `Start-SeEdge -Arguments @(...)` with the equivalent `Start-SeChrome -Headless -Arguments @(...)` / `Start-SeChrome -Arguments @(...)`. (There is a separate `Start-SeNewEdge` cmdlet for Chromium Edge, but it needs its own `msedgedriver.exe` set up separately — Chrome is the simpler path if it's installed.) |
| `unknown error: cannot find Chrome binary` | Chrome is not installed in any of the standard locations the bundled driver checks by default (or it's a per-user install in a non-default folder). | Locate it: `Get-ChildItem "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe","C:\Program Files\Google\Chrome\Application\chrome.exe","C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" -ErrorAction SilentlyContinue`. Then pass it explicitly: `Start-SeChrome -BinaryPath "<path found>" ...` |
| `session not created: This version of ChromeDriver only supports Chrome version 83` (or any specific old version) | The Selenium module bundles a `chromedriver.exe` from when it was published (years old); your actual installed Chrome is much newer, and the two must match. | Check your Chrome version: `(Get-Item "<chrome.exe path>").VersionInfo.ProductVersion`. Download the matching driver from https://googlechromelabs.github.io/chrome-for-testing/#stable (win64 chromedriver), put `chromedriver.exe` in its own folder (e.g. `drivers\`), then add `-WebDriverDirectory "<that folder>"` to the `Start-SeChrome` call so it uses your fresh copy instead of the module's bundled one. Chrome auto-updates periodically, so this may need repeating after a Chrome update. |
| Selenium `Start-SeEdge` fails or can't find a driver | Edge itself isn't installed, or the installed Edge version and the auto-downloaded driver are mismatched (rare, but can happen right after an Edge auto-update). | Confirm Edge opens normally; update the module: `Update-Module Selenium`. Otherwise switch to Chrome (see row above) — it's the more reliably supported browser in this module. |
| A page row always comes back `Error: element not found` even though the page loads fine | The CSS selector for `HeaderSelector`/`SubHeaderSelector` doesn't match the real page (site markup changed, or selector was copied wrong). | Re-run with `-Headless:$false` to watch it, and re-copy the selector from the browser's Inspect panel (right-click element → Inspect → right-click the HTML → Copy → Copy selector). |
| Login always shows `Failure` even with correct credentials | `PostLoginCheckSelector` doesn't match anything on the real post-login page, or the page takes longer than the fixed 2-3 second wait to load. | Verify the selector exists on the real logged-in page; if the site is just slow, increase the `Start-Sleep -Seconds 3` after the login click. |
| First-run credential prompt appears in Alt-Tab but the window never comes to the foreground / won't accept focus | This was a symptom of the older `Get-Credential` popup, which can fail to come forward when another process (the browser) is active in the background. | Already fixed — the script prompts in-console via `Read-Host` now (no separate popup window to lose focus). If you're seeing this, you may be on an older copy of the script. |
| `Exception calling ".ctor" with "2" argument(s): "Cannot process argument because the value of argument "userName" is not valid...` right after entering credentials | Chrome/ChromeDriver write background log lines (GCM/TensorFlow messages) directly to the console; if one of those lines interleaves with an in-progress `Read-Host` prompt, it can corrupt what PowerShell captures as your typed username, producing an empty/invalid value. | Move the credential prompt (`Get-SiteCredential` call) to *before* `Start-SeChrome` runs, so nothing is writing to the console yet while you're typing. Nothing needs cleaning up — the corrupted attempt fails before saving anything, so just re-run. |
| Two Chrome windows open per site, or one Chrome window is left running after the script finishes a site | A `Start-SeChrome` call was pasted in twice for the same site (e.g. while moving the credential prompt earlier) instead of replacing the original, so the browser launches twice and the first instance is never assigned to `$driver` or closed. | Make sure there is exactly one `$driver = if ($Headless) { Start-SeChrome ... } else { Start-SeChrome ... }` block per site, with `Get-SiteCredential` called once, above it. Remove any duplicate block. |
| Chrome opens but is stuck on a blank/unfilled login page for a long time | Not actually stuck — `Find-SeElement` waits up to 15 seconds per element, and every run pauses for the console credential prompt before the browser even opens. | Check whether `chrome.exe`/`chromedriver.exe` are still using CPU in Task Manager (still working) vs fully idle (actually frozen), and check the console for a waiting `Username:`/`Password:` prompt. Give it up to 30–45 seconds before assuming it's hung. |
