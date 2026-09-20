# Website E2E Test Automation (PowerShell)

## Project layout
```
website-e2e-tester/
├── scripts/
│   ├── Run-WebsiteE2ETests.ps1
│   ├── Run-QuickSiteCheck.ps1    ← lightweight variant: Excel only, only the pages you list per site
│   ├── Run-SiteAuditReport.ps1   ← same full run, audit-style Excel: Pass/Fail, Start/End time, duration per website
│   └── Update-ChromeDriver.ps1   ← run this whenever Chrome auto-updates and the driver falls out of sync
├── config/
│   ├── Site-Config.csv         ← fill in with your real 6-7 sites, full page lists (used by the main script AND Run-SiteAuditReport.ps1)
│   ├── Site-Config-Demo.csv    ← public sandbox sites, for trial runs
│   └── QuickCheck-Config.csv   ← same column format, but only the pages you want in a quick check
├── TestRun_*/                  ← auto-created per run of Run-WebsiteE2ETests.ps1, gitignored
├── QuickCheck_*/                ← auto-created per run of Run-QuickSiteCheck.ps1, gitignored
├── AuditRun_*/                  ← auto-created per run of Run-SiteAuditReport.ps1, gitignored
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
Edit `Site-Config.csv` — one `Login` row per site plus one `Page` row per page you want tested. See **Config File Reference** below for the full column list and worked examples covering every login pattern we've handled (unusual field names, two-step logins, iframes, shadow DOM, MFA, logout variants).

## Config File Reference

### Column reference — `Login` rows
| Column | Required? | Purpose |
|---|---|---|
| `SiteName` | Yes | Groups rows together; also used in screenshot filenames and report headings |
| `RowType` | Yes | `Login` for this row |
| `Enabled` | No (defaults to enabled) | Set to `No` to skip this site for a run without deleting its rows |
| `Url` | Yes | The login page URL |
| `UserFieldSelector` | Yes | Selector for the username/email field |
| `ContinueButtonSelector` | No | Only for two-step logins (username → Continue → separate password screen). Leave blank for single-screen logins |
| `PassFieldSelector` | Yes | Selector for the password field |
| `LoginButtonSelector` | Yes | Selector for the submit/Sign In/Log In button |
| `PostLoginCheckSelector` | Recommended | Selector for something that only exists after a successful login — this is what makes Success vs Failure real rather than just "didn't crash" |
| `RequiresMFA` | No | Set to `Yes` to pause after login and wait for you to complete 2FA/OTP manually. Requires `-Headless:$false` |
| `IframeSelector` | No | Selector for an `<iframe>` wrapping the login form, if present |
| `LogoutMenuSelector` | No | Selector for a menu/avatar that must be opened before the logout link is visible |
| `LogoutSelector` | No | Selector for the logout link/button. Leave both this and `LogoutMenuSelector` blank to skip logout testing |
| `LogoutPageUrl` | No | A page known to have the logout control on it. When set, the script navigates there first before attempting logout — needed whenever the last `Page` row might leave the browser somewhere without a logout link (a standalone widget/settings page with no shared nav). Without it, logout is attempted on whatever page the last `Page` row left off on |
| `FooterCaptureSelector` | No | Selector for text to **capture and log** on the post-login landing page (e.g. a footer server tag like "S09"). Not compared to any expected value — see Cookbook below |
| `CaptureFooterScreenshot` | No | Set to `Yes` to take a **dedicated screenshot** of the bottom of the post-login page (scrolls down first), saved as its own image and embedded in Word separately from the normal full-page shot. Use this when the footer *itself* is the test case — see Cookbook below |

### Column reference — `Page` rows
| Column | Required? | Purpose |
|---|---|---|
| `SiteName` / `RowType` | Yes | `RowType` = `Page` |
| `PageName` | Yes | Label used in the report and screenshot filename |
| `PageUrl` | Yes | Direct URL to navigate to (must work once logged in) |
| `HeaderSelector` / `ExpectedHeader` | No | Selector + text to verify a heading matches. Leave both blank for a screenshot-only check with no text validation |
| `SubHeaderSelector` / `ExpectedSubHeader` | No | Same idea, for a secondary heading |
| `FooterCaptureSelector` | No | Selector for text to **capture and log** on this page (e.g. a footer server/instance tag). Not compared to any expected value — see Cookbook below |
| `CaptureFooterScreenshot` | No | Set to `Yes` to take a dedicated screenshot of the bottom of this page, in addition to the normal full-page shot — see Cookbook below |

Every selector column above (on both row types) accepts any of the three syntaxes below — mix and match freely per field, per site.

### Selector syntax cheat sheet
| Syntax | When to use | Example |
|---|---|---|
| Plain CSS | The default — element has a stable `id`, `class`, or attribute | `#username`, `input[name='userId']`, `.login-btn` |
| XPath (starts with `//`) | Element only identifiable by its visible text — common for buttons/links labeled "Sign In", "Continue", "Log out" with no useful id | `//button[contains(text(),'Sign In')]` |
| Shadow DOM chain (contains `>>>`) | Element lives inside a **web component's shadow root** (common in modern component libraries) | `custom-login-widget >>> #username` (chain more `>>>` for nested shadow roots; a **closed** shadow root can't be reached by any tool) |

**A field's wrapping structure (`<form>`, `<table>`, nested `<div>`s) doesn't matter for CSS/XPath** — they match the element itself, not its container. The only structural things that DO need special handling are iframes (`IframeSelector`, above) and shadow DOM (the `>>>` syntax above); everything else — a login form inside a table, a deeply nested div, a React/Angular component's rendered markup — works with an ordinary selector.

### Cookbook — common variations

**Unusually-named fields** (no `username`/`password`/`login` naming at all):
```
UserFieldSelector: #email
PassFieldSelector: input[type='password']    ← works on almost any site: browsers require type="password" to mask input, regardless of what the site calls the field
LoginButtonSelector: //button[contains(text(),'Sign In')]
```

**Two-step login** (username screen → Continue → separate password screen):
```
UserFieldSelector: #email
ContinueButtonSelector: #continueBtn
PassFieldSelector: #password        ← found only after Continue is clicked
LoginButtonSelector: #submitBtn
```

**Login form inside an iframe:**
```
IframeSelector: iframe#login-frame
UserFieldSelector: #username        ← this selector is now relative to inside that iframe
PassFieldSelector: #password
```

**Login field inside a web component's shadow DOM:**
```
UserFieldSelector: app-login >>> #username
PassFieldSelector: app-login >>> #password
```

**Site with MFA:**
```
RequiresMFA: Yes
```
(run with `-Headless:$false` for that site so you can actually see and complete the MFA step)

**Multiple similar inputs on one page** (e.g., a login form embedded in a table with several rows) — a bare `#username` might not be unique enough; scope it with enough of the surrounding structure to be specific:
```
UserFieldSelector: table#loginTable tr:first-child input[name='username']
```

**Logout variants:**
```
Simple link, always visible:      LogoutSelector: a[href='/logout']
Behind a menu/avatar first:       LogoutMenuSelector: .user-avatar   |   LogoutSelector: a[href='/logout']
No id, only visible text:         LogoutSelector: //a[contains(text(),'Log out')]
Only reachable from one page:     LogoutPageUrl: https://example.com/dashboard   |   LogoutSelector: a[href='/logout']
```

**Capturing a footer/server tag that legitimately varies** (e.g. "S09", "S14" — which server or instance answered the request) — this is *different* from `HeaderSelector`/`ExpectedHeader`, which fail the row if the text doesn't match an exact expected value. `FooterCaptureSelector` never causes a Failure; it just records whatever text it finds (or `(not found)` if the selector doesn't match, or blank if the column is empty) into a new `CapturedInfo` field shown in both the Excel report and the Word report for that row:
```
FooterCaptureSelector: footer .server-tag
FooterCaptureSelector: //footer//span[contains(@class,'instance-id')]
FooterCaptureSelector: #footerPanel                 ← captures the whole footer block's text if there's no single dedicated element
```
Set it on a `Login` row to capture it once, right after login; set it on any `Page` row to capture it on that specific page (useful if the tag can differ page to page, e.g. if pages are served by different backend instances). Leave it blank anywhere you don't need it — it's fully optional and off by default.

Whenever `FooterCaptureSelector` finds something, that value is also folded into the **screenshot filenames** for that row, so you can tell which server answered just by looking at the `Screenshots\` folder — no need to open the report: `SiteA_00_Login_S09.png`, `SiteA_Dashboard_S14.png`, `SiteA_Dashboard_Footer_S14.png` (if `CaptureFooterScreenshot` is also set). If nothing is captured (selector blank, or not found), filenames fall back to the plain form (`SiteA_00_Login.png`, `SiteA_Dashboard.png`).

**Footer screenshot as its own test step** (different from the text capture above — use this when the footer's *appearance*, not just its text, is what needs checking, or you just want a dedicated bottom-of-page image without picking apart selectors):
```
CaptureFooterScreenshot: Yes
```
Set it on a `Login` row and/or any `Page` row. When `Yes`, the script scrolls to the bottom of that page and takes a separate screenshot of just that view (`{Site}_{Page}_Footer.png` in the `Screenshots\` folder), on top of the normal full-page screenshot it always takes. In `TestReport.docx`, this image appears right under the main screenshot for that row, labeled "Footer screenshot". In `TestResults.xlsx`, a `FooterScreenshotCaptured` column shows `Yes` for any row where it ran, so you can filter/scan for them quickly. Leave it blank (the default) anywhere you don't need the extra shot.

### Finding selectors via DevTools
Open the page in Chrome, right-click the element → **Inspect**, right-click the highlighted HTML in the Elements panel → **Copy** → **Copy selector**. To check whether a field sits inside an iframe or shadow DOM: look at what wraps it in the Elements panel — an `<iframe>` tag above it means you need `IframeSelector`; a `#shadow-root` node means you need the `>>>` chain syntax instead.

## Credentials
Every run, for every site, the script prompts for username and password right in the PowerShell console (password input is masked). Nothing is saved to disk anywhere — no file, no cache — so you'll be asked again next time, even for the same site. This is deliberate: no credential file to accidentally commit, leave behind, or have picked up by anything else on the machine.

## Run it

**Before you run** (quick checklist — most first-run-of-the-day issues come from one of these):
1. **Execution policy**, once per new PowerShell window:
   ```powershell
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
   ```
2. **Chrome/driver still matching?** If Chrome auto-updated since your last run, `.\scripts\Update-ChromeDriver.ps1` first (see Troubleshooting for what it fixes).
3. **You're in the right folder** — commands below assume the repo root. From inside `scripts\`, use `..\config\...` instead of `.\config\...`.
4. **Have credentials ready** — every site prompts fresh, every run (nothing is cached). For an `MFA`-enabled site, also have your phone/authenticator ready and remember to run with `-Headless:$false` for it.
5. **Watch for corporate security software interference** the first time you run on a new/locked-down machine — errors mentioning "Chrome updater", "Access is denied", or "unable to open pipe" (IPC) are almost always antivirus/EDR blocking the freshly-downloaded `chromedriver.exe` or Chrome's internal sandboxing, not a real script problem. See Troubleshooting below.

From the repo root:
```powershell
.\scripts\Run-WebsiteE2ETests.ps1 -ConfigPath .\config\Site-Config-Demo.csv -Headless:$false
```
Once the demo run looks right, switch to your real config:
```powershell
.\scripts\Run-WebsiteE2ETests.ps1 -ConfigPath .\config\Site-Config.csv
```
Add `-Headless:$false` any time you want to watch the browser while debugging selectors. `TestRun_*\` folders are created next to the repo root automatically and are already gitignored.

**During an incident**, running only the affected site(s) is faster than a full pass: set `Enabled` to `No` for every site except the ones you're checking, run, then flip back. The console's final `Summary: X Success, Y Failure, Z Error` line tells you the outcome at a glance without opening Excel.

## Output
Each run creates a timestamped folder containing:
- `TestResults.xlsx` — Site, Page, URL, Status (Success/Failure/Error), Timestamp, Duration, **CapturedInfo** (footer/server tag text, if `FooterCaptureSelector` is configured for that row — blank otherwise), **FooterScreenshotCaptured** (`Yes` if `CaptureFooterScreenshot` was set for that row — blank otherwise), Error message. Rows include LOGIN, one per configured Page, and LOGOUT (if configured).
- `TestReport.docx` — one section per site, one sub-section per row, with the URL/duration line, a captured-info line (if `FooterCaptureSelector` is configured for that row), the main screenshot, a separate footer screenshot (if `CaptureFooterScreenshot` is configured for that row), and any mismatch note
- `Screenshots\` — one PNG per step: pre-login page, (for two-step logins) the screen right after Continue is clicked, post-login landing page, each tested page (headers/subheaders visible in-frame), and the post-logout page if logout is configured, plus a `_Footer.png` for any row with `CaptureFooterScreenshot: Yes`. Every full-page screenshot is captured with the browser window temporarily resized to the page's full scroll height, so long pages are captured in their entirety, not just what's visible on screen.

## Status meanings
- **Success** — page loaded and header/subheader (if checked) matched
- **Failure** — page loaded fine, but expected text wasn't found (a real content/UI problem)
- **Error** — something broke the test itself (selector not found, timeout, site unreachable) — treat separately from Failure since it may just mean a selector needs updating, not that the site is broken
- **Skipped** — that site's `Enabled` column was set to `No` for this run

## Summary line and exit code
At the end of every run, the console prints a one-line summary (`Summary: 14 Success, 1 Failure, 0 Error`) in green if everything passed or red if anything didn't — no need to open Excel just to know whether the run was clean. The script also exits with code `1` if there was any Failure or Error, and `0` if everything passed (Skipped sites don't count against this). That makes it usable in Windows Task Scheduler (check the task's last run result) or any CI pipeline that wants to alert automatically on a bad run.

## ROI / business-value estimate
Both `Run-WebsiteE2ETests.ps1` and `Run-QuickSiteCheck.ps1` write a second Excel worksheet, **"ROI Summary"**, alongside the results — a rough estimate of the manual effort each run replaced, projected out to a monthly/annual figure. It's meant for showing "here's what this is worth" to a manager or in a status update, not as a precise cost model.

**How it's calculated** — every one of these is an assumption you supply, not something the script measures or knows on its own:
| Parameter | Default | Meaning |
|---|---|---|
| `-ManualMinutesPerCheck` | `3` | How long a person takes to manually open one page/login/logout step, verify it, and note the result |
| `-RunsPerDay` | `4` | How many times a day this suite (or its manual equivalent) actually runs — bump this up for incident days |
| `-HourlyRate` | `22` | Fully-loaded hourly cost of whoever would do this manually. Set to `0` to skip cost figures entirely and show time only |
| `-CurrencySymbol` | `$` | Only used if `-HourlyRate` is set above `0` — change to `₹`, `€`, etc. as needed |

The math: `checks performed × ManualMinutesPerCheck` = manual-equivalent time, compared against the run's actual elapsed time, then multiplied by `RunsPerDay` and projected across ~22 working days/month and ~260 working days/year. If `-HourlyRate` is set, the same time savings are converted to cost savings too.

Cost figures are **on by default** now (`-HourlyRate 22`, in `$`) — every run's Excel "ROI Summary" tab and console line include a dollar estimate unless you override it.

**Example (overriding the defaults):**
```powershell
.\scripts\Run-WebsiteE2ETests.ps1 -ConfigPath .\config\Site-Config.csv -ManualMinutesPerCheck 4 -RunsPerDay 6 -HourlyRate 800 -CurrencySymbol "₹"
```
This says: each manual check would take ~4 minutes, this suite effectively runs ~6 times a day (accounting for incident-driven re-runs), and the fully-loaded hourly cost of doing it by hand is ₹800/hour. The console prints a one-line estimate (e.g. `ROI (est.): ~42 min saved this run (88%) vs. manual - projected ~15.4 hrs/month at 6 runs/day`), and the full breakdown — including cost — lands in the "ROI Summary" tab of the Excel report.

**Treat the numbers as a starting estimate, not a fact**: the defaults are generic guesses. If you actually know how long your manual process takes and how often it really runs, pass your own values — the estimate is only as good as the assumptions behind it, and presenting it as measured savings without saying so would be misleading.

## Scheduling
Wrap the `.ps1` call in a Windows Task Scheduler job so it can run on a timer instead of manually.

**Important limitation first:** since credentials are prompted fresh every run with no caching (by design — see Credentials above), a scheduled run will **pause and wait at the credential prompt** for each site, exactly like a manual run does. This isn't something Task Scheduler can supply an answer to on its own. So scheduling here is really "one click to kick off a run at a set time" rather than "fully unattended" — someone still needs to be at the machine to type credentials when it fires, unless you deliberately revisit the no-caching design later for scheduled use specifically.

**Setting up the task:**
1. Open Task Scheduler (Start menu → type `Task Scheduler`, or run `taskschd.msc`).
2. In the right-hand pane, click **Create Basic Task...**
3. **Name**: something like `Website E2E Tests`. Click Next.
4. **Trigger**: choose Daily (or Weekly, or another cadence) and set the time you want it to fire. Click Next.
5. **Action**: choose **Start a program**. Click Next.
6. Fill in:
   - **Program/script**: `powershell.exe`
   - **Add arguments**:
     ```
     -ExecutionPolicy Bypass -File "C:\Code\website-e2e-tester\scripts\Run-WebsiteE2ETests.ps1" -ConfigPath "C:\Code\website-e2e-tester\config\Site-Config.csv"
     ```
     (adjust the paths to match where your project actually lives)
   - **Start in**: `C:\Code\website-e2e-tester\scripts`
7. Click Next, review the summary, then **Finish**.
8. Right-click the new task in the list → **Properties**:
   - On the **General** tab, consider checking **Run whether user is logged on or not** only if you understand the credential-prompt limitation above — with this checked, there's no visible console for you to type into when it reaches a prompt, so the task will simply hang until it's manually stopped. For most setups here, leave it as **Run only when user is logged on**, so you still get a visible console window to respond to prompts.
   - On the **Conditions** tab, uncheck "Start the task only if the computer is on AC power" if this is a laptop you don't want the schedule skipped on battery.
9. Test it immediately: right-click the task → **Run**, and confirm a PowerShell window opens and behaves the same as running the command manually.

**Checking whether a scheduled run succeeded**, without opening any report:
- In Task Scheduler, select the task and look at **Last Run Result** in the bottom pane. `(0x0)` means the script exited cleanly with no Failure/Error rows (see Summary line and exit code above); any other value means something in that run needs attention — open the corresponding `TestRun_*` folder to see what.

**For a hands-off nightly run of only the sites that don't need MFA or manual attention**, set `Enabled` to `No` on any MFA-required sites in a dedicated scheduled-run copy of the config, and schedule that copy instead of your full `Site-Config.csv` — that way the automatic run doesn't stall waiting for a 2FA code nobody's there to enter.

## Quick Site Check (lightweight run)
`scripts\Run-QuickSiteCheck.ps1` is a trimmed-down sibling of the main script for when you only need a fast spot-check — e.g. during an incident, confirm Login + Account Overview still work and capture the server tag, without running (or waiting for) every page across every site.

**What's different from `Run-WebsiteE2ETests.ps1`:**
- Uses the **same engine** underneath — identical selector handling (CSS/XPath/shadow-DOM), login patterns (two-step, MFA, iframe), full-page and footer screenshots, and fresh-credential-every-run behavior. Nothing about *how* it drives the browser is different.
- Reads a **separate, smaller config file** (`config\QuickCheck-Config.csv` by default) in the exact same 22-column format as `Site-Config.csv` — you just list fewer `Page` rows per site (e.g. only `Login` + `AccountsOverview`, skip the rest). Since it's a different file, your full `Site-Config.csv` is untouched.
- Produces **Excel only** — no `.docx`, so Word isn't even required to run this script. The Excel sheet (`QuickCheckResults.xlsx`) has just: `Site`, `Page` (holds `LOGIN`/`LOGOUT`/`SITE-LEVEL` for those special rows, or the page name), `URL`, `Status`, `TimeTakenSec`, `Error` — same color-coded Success/Failure/Error/Skipped as before.
- Footer/server-tag capture (`FooterCaptureSelector`, `CaptureFooterScreenshot`) works exactly as in the main script, including multiple different tags across different pages of the same site (e.g. Login shows `S09`, Account Overview shows `S14`) — each gets its own tagged screenshot file in `Screenshots\`. The captured text itself isn't a separate Excel column here (keeps the sheet to the essentials above); read it off the screenshot filename or reuse the full script's Excel output (which does include a `CapturedInfo` column) if you need it in a spreadsheet cell too.
- Also writes an **"ROI Summary"** worksheet, same as the main script — see **ROI / business-value estimate** above. Same `-ManualMinutesPerCheck` / `-RunsPerDay` / `-HourlyRate` / `-CurrencySymbol` parameters apply here too.

**Run it:**
```powershell
.\scripts\Run-QuickSiteCheck.ps1 -ConfigPath .\config\QuickCheck-Config.csv -Headless:$false
```
Same flags as the main script (`-ConfigPath`, `-OutputFolder`, `-Headless`), same exit-code behavior (`0` clean, `1` if anything Failed/Errored) for Task Scheduler/CI use.

**Building your `QuickCheck-Config.csv`**: copy just the `Login` row and the specific `Page` row(s) you care about from `Site-Config.csv` for each site — the column layout is identical, so no translation needed. Leave `LogoutSelector` blank on a site if you don't want it to bother logging out for a quick check (or keep it if you'd rather always close the session cleanly).

## Site Audit Report (Pass/Fail, timestamps, per-website duration)
`scripts\Run-SiteAuditReport.ps1` is another sibling of the main script — same automation, same Word report, same footer/server-tag capture, same ROI Summary worksheet — but a differently-shaped **Excel** report, built for an audit-style record rather than a quick scan.

**What's different from `Run-WebsiteE2ETests.ps1`:**
- Uses the **same engine and the same config file** — `config\Site-Config.csv` by default, no separate config needed. Logs into every site, visits every configured page, logs out — full run, same as the main script.
- Still produces the **Word report** (`TestReport.docx`), with the same screenshots, footer screenshots, and captured server-tag info per row.
- The **Excel** report (`TestResults.xlsx`, `Results` sheet) uses a different column set:

  | Column | Meaning |
  |---|---|
  | `Site_Name` | Same as `Site` elsewhere in this project |
  | `RowType` | Holds `LOGIN`, `LOGOUT`, `SITE-LEVEL`, `CONFIG`, or the configured page name — same convention as the `Page` column in the other scripts, just relabeled |
  | `URL` | The URL for that step |
  | `Status` | **`Pass`** / **`Fail`** / `Error` / `Skipped` — note this script uses `Pass`/`Fail` instead of `Success`/`Failure` used elsewhere in this project |
  | `Start_Time` / `End_Time` | Timestamp for when that specific row's step began and ended (not just one timestamp per row, like the other scripts) |
  | `Duration(Per website)` | **The same value repeated on every row for that site** — total elapsed time for the whole site (login through logout), not that individual row's own step time. Look at the Login row's `Start_Time` and the Logout row's `End_Time` for that site to see the window this duration covers |
  | `Error` | Same as elsewhere |

  Color-coding: `Pass` = green, `Fail` = yellow, `Error` = pink, `Skipped` = gray — same scheme, new labels.
- Also writes the **"ROI Summary"** worksheet, same as the main script — see **ROI / business-value estimate** above. Same `-ManualMinutesPerCheck` / `-RunsPerDay` / `-HourlyRate` / `-CurrencySymbol` parameters apply here too.

**Run it:**
```powershell
.\scripts\Run-SiteAuditReport.ps1 -ConfigPath .\config\Site-Config.csv -Headless:$false
```
Same flags as the main script. Output goes to a separate `AuditRun_*\` folder (gitignored), so it never collides with `TestRun_*\` or `QuickCheck_*\` from the other two scripts.

**When to use which script** — quick reference:
- **`Run-WebsiteE2ETests.ps1`** — full daily/incident run, Word + Excel, `Success`/`Failure`/`Error` labels, per-step duration
- **`Run-QuickSiteCheck.ps1`** — fast spot-check of a handful of pages, Excel only, no Word
- **`Run-SiteAuditReport.ps1`** — full run like the main script, but Excel formatted as an audit trail: `Pass`/`Fail` labels, explicit Start/End timestamps per row, and total time-per-website instead of time-per-page

## Troubleshooting

### PowerShell & script execution
| Symptom | Root cause | Fix |
|---|---|---|
| `running scripts is disabled on this system` (`UnauthorizedAccess` / `PSSecurityException`) | Windows blocks unsigned local scripts by default. Resets every new PowerShell window if set with `-Scope Process`. | Per-session: `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force`<br>Permanent (no admin needed): `Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned` |
| `The term '.\Run-WebsiteE2ETests.ps1' is not recognized...` | Wrong relative path — the script lives in `scripts\`, not the repo root. | Run `.\scripts\Run-WebsiteE2ETests.ps1 ...` from the repo root, or `.\Run-WebsiteE2ETests.ps1` from inside `scripts\`. Check with `Get-ChildItem -Recurse -Filter *.ps1`. |
| Script downloaded/unzipped from elsewhere won't run even after setting execution policy | Windows marks files from a zip/download as "blocked" (Zone.Identifier). | From the project root: `Get-ChildItem -Recurse | Unblock-File` |
| Double-clicking the `.ps1` file in File Explorer does nothing, or a window flashes and closes instantly | Windows' default action for `.ps1` files is "open in Notepad" or a security-restricted run, not "execute in a persistent console" — this is intentional (prevents malicious scripts running via double-click). | Always launch from an actual PowerShell window: open PowerShell/Windows Terminal, `cd` to the project, then run the command. Never double-click the file itself. |
| Selenium/COM behavior is inconsistent or COM objects don't release properly | Running from PowerShell ISE, which is deprecated and has known quirks with COM automation and some module output. | Use a regular PowerShell console or Windows Terminal instead of ISE for this script. |
| `Install-Module` fails / times out against PSGallery | PowerShell 5.1 defaults to an older TLS version PSGallery no longer accepts. | Run once per session before installing: `[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12` |
| `Import-Module : The specified module 'Selenium' (or 'ImportExcel') was not loaded because no valid module file was found` | The module was never actually installed (different from the TLS-timeout row above — this means the install never succeeded at all). | `Install-Module Selenium -Scope CurrentUser -Force` and `Install-Module ImportExcel -Scope CurrentUser -Force`, then confirm with `Get-Module -ListAvailable Selenium,ImportExcel`. |
| Parser errors mentioning `Missing ')' `, `Unexpected token`, or garbled characters like `â€”` | A non-ASCII character (e.g. an em-dash `—`) got misread — PS 5.1 reads `.ps1` files without a BOM using the system ANSI codepage, not UTF-8. | Replace any non-ASCII punctuation with plain ASCII (e.g. `—` → `-`). Check for others with: `Get-ChildItem *.ps1 -Recurse \| Select-String -Pattern "[^\x00-\x7F]"` |
| First-run credential prompt appears in Alt-Tab but the window never comes to the foreground / won't accept focus | This was a symptom of the older `Get-Credential` popup, which can fail to come forward when another process (the browser) is active in the background. | Already fixed — the script prompts in-console via `Read-Host` now (no separate popup window to lose focus). If you're seeing this, you may be on an older copy of the script. |
| `Exception calling ".ctor" with "2" argument(s): "Cannot process argument because the value of argument "userName" is not valid...` right after entering credentials | Chrome/ChromeDriver write background log lines (GCM/TensorFlow messages) directly to the console; if one of those lines interleaves with an in-progress `Read-Host` prompt, it can corrupt what PowerShell captures as your typed username, producing an empty/invalid value. | Move the credential prompt (`Get-SiteCredential` call) to *before* `Start-SeChrome` runs, so nothing is writing to the console yet while you're typing. Nothing needs cleaning up — the corrupted attempt fails before saving anything, so just re-run. |

### Config file / CSV
| Symptom | Root cause | Fix |
|---|---|---|
| `Import-Csv : Could not find a part of the path ...` followed by a `Done.` with empty/near-empty reports | Config path was relative to the wrong folder (e.g. run from `scripts\` with a `.\config\...` path meant for the repo root), so `Import-Csv` failed as a non-fatal error and the rest of the script ran anyway with zero rows. | Use the right relative path for where you're standing (`.\scripts\Run-...ps1 -ConfigPath .\config\...` from repo root, or `..\config\...` from inside `scripts\`). The script hard-stops with a clear message if the config file is genuinely missing. |
| A selector column seems to be silently ignored (e.g. `IframeSelector` never triggers, `RequiresMFA` never pauses) even though it looks filled in | A typo or extra trailing space in the **column header** (e.g. `RequiresMFA ` with a trailing space, or wrong case in a hand-edited CSV) means PowerShell reads that property as blank/null rather than erroring — it fails silently, not loudly. | Compare your header row character-for-character against the Config File Reference above. Quick check: `(Import-Csv .\config\Site-Config.csv | Select-Object -First 1).PSObject.Properties.Name` lists the exact column names PowerShell actually sees. |
| Only the first of several intended logins for a site seems to run; the others are ignored | Two `Login` rows accidentally exist for the same `SiteName` (usually a copy-paste mistake) — the script only ever uses the first one (`Select-Object -First 1`) and silently ignores the rest. | Each `SiteName` should have exactly one `Login` row. Check with: `Import-Csv .\config\Site-Config.csv \| Where-Object RowType -eq 'Login' \| Group-Object SiteName \| Where-Object Count -gt 1` |
| A site that should be grouped together shows up as two separate (and broken) entries in the report | A trailing space or inconsistent capitalization in `SiteName` between rows (e.g. `"SauceDemo "` vs `"SauceDemo"`) makes `Group-Object SiteName` treat them as two different sites — one of which then has no Login row and errors with `CONFIG: No Login row found`. | Make sure `SiteName` is spelled identically (same case, no trailing spaces) on every row for that site. |
| A CSS selector containing special characters looks different after editing the CSV in Excel, and stops matching | Excel can silently reformat cell content on save (e.g. altering quote characters, or misinterpreting a leading `=`/`-` as the start of a formula in some locales/versions). | Edit `Site-Config.csv` in a plain text or code editor (Notepad, VS Code) rather than Excel, especially for selector columns with punctuation. If you must use Excel, re-open the saved file afterward and spot-check the affected cells. |

### Chrome & ChromeDriver
| Symptom | Root cause | Fix |
|---|---|---|
| `A parameter cannot be found that matches parameter name 'Arguments'` (or `'Headless'`) on `Start-SeEdge` | This module's `Start-SeEdge` only wraps the old, legacy (pre-Chromium) Edge driver, which doesn't accept Chromium-style flags at all. | Switch to Chrome instead, which fully supports both: replace `Start-SeEdge -Headless -Arguments @(...)` / `Start-SeEdge -Arguments @(...)` with the equivalent `Start-SeChrome -Headless -Arguments @(...)` / `Start-SeChrome -Arguments @(...)`. (There is a separate `Start-SeNewEdge` cmdlet for Chromium Edge, but it needs its own `msedgedriver.exe` set up separately — Chrome is the simpler path if it's installed.) |
| `unknown error: cannot find Chrome binary` | Chrome is not installed in any of the standard locations the bundled driver checks by default (or it's a per-user install in a non-default folder). | Locate it: `Get-ChildItem "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe","C:\Program Files\Google\Chrome\Application\chrome.exe","C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" -ErrorAction SilentlyContinue`. Then pass it explicitly: `Start-SeChrome -BinaryPath "<path found>" ...` |
| Chrome binary still not found, even after checking the three standard paths above | Chrome was installed through the **Microsoft Store**, which uses a sandboxed package location entirely different from a normal install. | Check `Get-ChildItem "$env:LOCALAPPDATA\Packages\*Chrome*" -Recurse -Filter chrome.exe -ErrorAction SilentlyContinue`. If it's a Store install, consider installing the regular desktop version from google.com/chrome instead — it's the more reliably automatable option. |
| `session not created: This version of ChromeDriver only supports Chrome version 83` (or any specific old version) | The Selenium module bundles a `chromedriver.exe` from when it was published (years old); your actual installed Chrome is much newer, and the two must match. | Run `.\scripts\Update-ChromeDriver.ps1` — it auto-detects your Chrome version and downloads a matching driver, skipping the download entirely if you're already in sync. Chrome auto-updates periodically, so re-run this any time you suspect a mismatch. |
| `session not created` mentioning an **architecture** mismatch, or the driver simply won't launch, on a Surface/ARM-based Windows machine | You're on ARM64 Windows, which needs the `arm64` chromedriver build, not `win64`. | In `Update-ChromeDriver.ps1`, change the platform filter from `'win64'` to `'arm64'` when selecting the download URL. |
| `DevToolsActivePort file doesn't exist` | ChromeDriver started the Chrome process but Chrome failed to fully initialize — common in restricted/locked-down environments (permissions, antivirus interference, or resource constraints). | Add `--no-sandbox` and `--disable-dev-shm-usage` to the `Start-SeChrome -Arguments` list. If it persists, check Task Manager for orphaned `chrome.exe`/`chromedriver.exe` processes and kill them first, and confirm antivirus isn't blocking Chrome's helper processes (see the IPC row below). |
| `unable to find chromedriver executable` (or similar), distinct from "cannot find Chrome binary" | `$driverDir` (`-WebDriverDirectory`) points at a folder that doesn't actually contain `chromedriver.exe` yet — usually because `Update-ChromeDriver.ps1` was never run on this machine. | Run `.\scripts\Update-ChromeDriver.ps1` first (creates the folder and downloads a matching driver), then re-run the test script. |
| Selenium `Start-SeEdge` fails or can't find a driver | Edge itself isn't installed, or the installed Edge version and the auto-downloaded driver are mismatched (rare, but can happen right after an Edge auto-update). | Confirm Edge opens normally; update the module: `Update-Module Selenium`. Otherwise switch to Chrome (see row above) — it's the more reliably supported browser in this module. |
| Two Chrome windows open per site, or one Chrome window is left running after the script finishes a site | A `Start-SeChrome` call was pasted in twice for the same site (e.g. while moving the credential prompt earlier) instead of replacing the original, so the browser launches twice and the first instance is never assigned to `$driver` or closed. | Make sure there is exactly one `$driver = if ($Headless) { Start-SeChrome ... } else { Start-SeChrome ... }` block per site, with `Get-SiteCredential` called once, above it. Remove any duplicate block. |
| Errors mentioning "Chrome updater", `Access is denied`, or `unable to open pipe` (IPC) — usually on a new or locked-down corporate machine | Antivirus/EDR software (or Group Policy) intercepting Chrome's internal process-spawning and named-pipe communication, or blocking the freshly-downloaded `chromedriver.exe` from running since it's an unfamiliar/unsigned binary. Not a script or version-mismatch problem. | First, unblock the driver: `Unblock-File "C:\Code\website-e2e-tester\drivers\chromedriver.exe"`. If it persists, add sandbox-disabling flags to both `Start-SeChrome -Arguments` lines in the script: `@('--ignore-certificate-errors','--no-sandbox','--disable-dev-shm-usage')` (safe here since this is a controlled automated-testing browser instance, not one browsing untrusted sites). If it still persists, this is a genuine security-policy block that needs your IT/security team to allowlist `chrome.exe`/`chromedriver.exe`'s automated behavior — not something fixable from inside the script. |

### Automation runtime (selectors, login logic, screenshots)
| Symptom | Root cause | Fix |
|---|---|---|
| No screenshots are produced anywhere, even when Status shows Success | The script called `Save-SeScreenshot -Driver ... -Path ...`, but that cmdlet doesn't accept `-Driver` — screenshots need two steps in this module. | Everywhere the script calls `Save-SeScreenshot -Driver $driver -Path $x`, change it to: `Invoke-SeScreenshot -Driver $driver \| Save-SeScreenshot -Path $x` (already fixed in the current script — if you're seeing this, you're on an old copy). |
| A page row always comes back `Error: element not found` even though the page loads fine | The CSS selector for `HeaderSelector`/`SubHeaderSelector` doesn't match the real page (site markup changed, or selector was copied wrong). | Re-run with `-Headless:$false` to watch it, and re-copy the selector from the browser's Inspect panel (right-click element → Inspect → right-click the HTML → Copy → Copy selector). |
| Login always shows `Failure` even with correct credentials | `PostLoginCheckSelector` doesn't match anything on the real post-login page, or the page takes longer than the fixed 2-3 second wait to load. | Verify the selector exists on the real logged-in page; if the site is just slow, increase the `Start-Sleep -Seconds 3` after the login click. |
| Chrome opens but is stuck on a blank/unfilled login page for a long time | Not actually stuck — `Find-SeElement` waits up to 15 seconds per element, and every run pauses for the console credential prompt before the browser even opens. | Check whether `chrome.exe`/`chromedriver.exe` are still using CPU in Task Manager (still working) vs fully idle (actually frozen), and check the console for a waiting `Username:`/`Password:` prompt. Give it up to 30–45 seconds before assuming it's hung. |
| LOGOUT step times out with `element not found`, even though the site works fine and logout definitely exists somewhere on it | The logout control only lives on certain pages (e.g. a dashboard/account page), but the last `Page` row visited left the browser on a different page (a standalone widget/settings page) with no shared navigation — LOGOUT always runs on whatever page the last `Page` row ended on unless told otherwise. | Set `LogoutPageUrl` on that site's Login row to a page you know has the logout control — the script will navigate there before attempting logout. |

### Microsoft Word (report generation)
| Symptom | Root cause | Fix |
|---|---|---|
| `Exception setting "SaveAs": Cannot convert ... value of type "psobject" to type "Object"` | Known PowerShell/COM marshalling quirk — Word's `SaveAs` needs plain typed values, not PowerShell's boxed variable wrapper. | Cast explicitly before wrapping in `[ref]`: `$doc.SaveAs([ref]([string]$docxPath), [ref]([int]16))` (already fixed in the current script). |
| Script never returns / doesn't respond to `Ctrl+C` | Word (or the browser) is blocked on a hidden dialog box — COM calls are synchronous and can't be interrupted from the console. | In a **new** window: `Get-Process WINWORD -ErrorAction SilentlyContinue \| Stop-Process -Force` and `Get-Process chrome,chromedriver -ErrorAction SilentlyContinue \| Stop-Process -Force`. The script already sets `$word.DisplayAlerts = 0` to prevent most of these. If it still hangs, temporarily set `$word.Visible = $true` to see and dismiss the dialog manually. |
| `Retrieving the COM class factory ... failed` for `Word.Application` | Microsoft Word isn't installed (or is installed as a Microsoft 365 "click-to-run" app not registered for COM) on this machine. | Install/repair desktop Word, or run the script on a machine that has it — the `.docx` step specifically needs local Word. |
| `The process cannot access the file '...TestReport.docx' because it is being used by another process` | A previous run's Word document is still open (crashed run left Word running, or you have the file open yourself). | Close Word / run `Get-Process WINWORD \| Stop-Process -Force`, then re-run. |
| A document-recovery / "we recovered some of your work" pane appears when Word launches, and the script seems to hang despite `DisplayAlerts = 0` | A previous crashed run left Word with unsaved recovery data; `DisplayAlerts` suppresses many prompts but not the recovery pane itself. | Open Word manually once, dismiss/discard the recovery pane, close Word, then re-run the script. |
| Word COM automation fails intermittently — works sometimes, fails other times, with no consistent error | Mixed Office installation types on the machine (e.g. a Click-to-Run Microsoft 365 install alongside leftovers from an old MSI-based Office install) can cause flaky COM registration. | Run `Online Repair` on your Office installation (Settings → Apps → Microsoft Office → Modify → Online Repair). |

### Excel report (ImportExcel)
| Symptom | Root cause | Fix |
|---|---|---|
| `Import-Module : The specified module 'ImportExcel' was not loaded because no valid module file was found` | The module isn't installed yet. | `Install-Module ImportExcel -Scope CurrentUser -Force` (see also the TLS row under PowerShell if the install itself fails). |
| `New-ConditionalText : A parameter cannot be found that matches parameter name 'BackgroundColor'` (or similar) after updating the module | A newer/older `ImportExcel` version changed a parameter name — the module isn't perfectly stable across versions. | Check the installed version with `Get-Module -ListAvailable ImportExcel`, and check that cmdlet's current parameters with `Get-Help New-ConditionalText -Full`. Pin a known-working version if needed: `Install-Module ImportExcel -RequiredVersion <version> -Force`. |
| `TestResults.xlsx` exists but Excel says the file is corrupt / unreadable / format doesn't match the extension when you open it | The script was interrupted (crashed, killed, or the machine slept) partway through writing the Excel file. | Delete the bad file and re-run — each run writes into a fresh timestamped `TestRun_*` folder, so this never affects previous good runs. |
| `The process cannot access the file '...TestResults.xlsx' because it is being used by another process` | The Excel file from this same run is already open (rare, since each run gets a fresh timestamped filename — usually means you manually opened it while the script was still writing). | Close Excel before the script finishes, or just wait for the run to complete before opening the file. |

## Git / GitHub

**Cloning fresh** (setting up on a different machine, or after a reinstall):
```bash
git clone https://github.com/<your-username>/website-e2e-tester.git
cd website-e2e-tester
```
Then follow One-time setup above — modules, Chrome path, `Update-ChromeDriver.ps1` — since none of that is stored in git.

**Getting the latest changes** (before starting work, or if changes were made from another machine):
```bash
git pull
```

**Pushing your own changes** — check what actually changed first, since only the script, README, and config templates should ever be committed:
```bash
git status
```
Confirm the list only shows files you meant to change — nothing from a `TestRun_*` folder should ever appear (already gitignored, so seeing one there would be unusual and worth investigating). Then:
```bash
git add <file1> <file2> ...
git commit -m "Describe what changed"
git push
```
Avoid `git add .` out of habit — naming files explicitly is a second safety check against accidentally committing something that shouldn't be there.

**If `git push` is rejected** ("fetch first" / "Updates were rejected"):
```bash
git pull origin main --allow-unrelated-histories
```
If that reports a conflict (most commonly on `README.md` if it was edited from two places), open the flagged file, resolve the `<<<<<<<` / `=======` / `>>>>>>>` markers by hand, then:
```bash
git add <resolved-file>
git commit -m "Merge remote changes"
git push
```

**If `git push` asks for a password and your normal login doesn't work**: GitHub no longer accepts account passwords over HTTPS. Generate a **Personal Access Token** (GitHub.com → Settings → Developer settings → Personal access tokens) and paste that in place of the password when prompted.
