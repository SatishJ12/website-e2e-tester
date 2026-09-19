<#
.SYNOPSIS
    End-to-end website smoke test: logs into each site, visits a predefined
    list of pages, checks header/subheader text, takes screenshots, and
    writes results to Word (report) + Excel (status matrix).

.REQUIREMENTS
    Install-Module Selenium     -Scope CurrentUser -Force
    Install-Module ImportExcel  -Scope CurrentUser -Force
    Microsoft Word installed (for the .docx report via COM automation)

.USAGE
    .\Run-WebsiteE2ETests.ps1 -ConfigPath .\Site-Config.csv
    Prompts for username/password every run, for every site - nothing is
    saved to disk between runs.
#>

param(
    [string]$ConfigPath   = "$PSScriptRoot\..\config\Site-Config.csv",
    [string]$OutputFolder = "$PSScriptRoot\..\TestRun_$(Get-Date -Format yyyyMMdd_HHmmss)",
    [switch]$Headless = $true,

    # ---- ROI / business-value estimate (see "ROI Summary" sheet in the Excel
    # report) - all four are just assumptions you supply; the script doesn't
    # know your actual manual process or pay rate, so adjust these to match
    # reality rather than trusting the defaults for anything you'd present.
    [double]$ManualMinutesPerCheck = 3,     # how long a person takes to manually open a page, verify it, and note the result
    [int]$RunsPerDay               = 4,     # how many times a day this suite (or a manual equivalent) actually gets run
    [double]$HourlyRate            = 22,    # fully-loaded hourly cost of whoever would do this manually; set to 0 to skip cost figures, time-only
    [string]$CurrencySymbol        = "$"    # only used if HourlyRate > 0
)

$ScriptStartTime = Get-Date
Write-Host "Started: $($ScriptStartTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Cyan

Import-Module Selenium -ErrorAction Stop
Import-Module ImportExcel -ErrorAction Stop
Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
$ScreenshotFolder = Join-Path $OutputFolder "Screenshots"
New-Item -ItemType Directory -Path $ScreenshotFolder -Force | Out-Null

if (-not (Test-Path $ConfigPath)) {
    Write-Error "Config file not found: $ConfigPath"
    exit 1
}
$AllRows = Import-Csv $ConfigPath
$Results = New-Object System.Collections.Generic.List[object]


# Captures (does NOT assert/compare) the text of an element such as a footer
# server tag ("S09", "S14", etc.) that legitimately varies run to run and
# site to site. Returns the trimmed text, or "" if the selector is blank,
# the element isn't found, or anything else goes wrong - this must never
# throw and must never affect Success/Failure/Error status.
function Get-SeCapturedText {
    param($Driver, [string]$Selector, [double]$Timeout = 5)
    if (-not $Selector) { return "" }
    try {
        $el = Find-SeElementAny -Driver $Driver -Selector $Selector -Timeout $Timeout
        if ($el -and $el.Text) { return $el.Text.Trim() }
        return ""
    } catch {
        return "(not found)"
    }
}


# Turns captured text (e.g. "S09") into something safe to put in a filename.
# Returns "" for blank/whitespace-only/"(not found)" input, so callers can
# just check for a non-empty result before adding it to a filename.
function Get-CaptureFileTag {
    param([string]$Text)
    if (-not $Text) { return "" }
    $Text = $Text.Trim()
    if (-not $Text -or $Text -eq '(not found)') { return "" }
    $tag = ($Text -replace '[^\w\-]', '_').Trim('_')
    if ($tag.Length -gt 30) { $tag = $tag.Substring(0, 30) }   # keep filenames sane if the captured text is a long sentence
    return $tag
}

function Get-SiteCredential {
    param([string]$SiteName)
    Write-Host "`nEnter login credentials for $SiteName" -ForegroundColor Cyan
    $username = Read-Host "Username"
    $securePassword = Read-Host "Password" -AsSecureString
    return New-Object System.Management.Automation.PSCredential($username, $securePassword)
}

# Finds an element using a CSS selector, an XPath expression, or a shadow-DOM
# piercing chain. A selector starting with "//" is XPath. A selector containing
# ">>>" is a shadow-DOM chain - e.g. "custom-login-widget >>> #username" means
# "find custom-login-widget by CSS, then find #username inside its shadow root"
# (chain as many ">>>" as there are nested shadow roots). Everything else is a
# normal CSS selector. Note: a CLOSED shadow root cannot be pierced by any
# method, including this one - that's a browser-level restriction, not a
# limitation specific to this script.
function Find-SeElementAny {
    param($Driver, [string]$Selector, [double]$Timeout = 10)
    if ($Selector -like '*>>>*') {
        $parts = $Selector -split '\s*>>>\s*'
        $js = 'let el = document.querySelector(arguments[0]); if (!el) return null;'
        for ($i = 1; $i -lt $parts.Count; $i++) {
            $js += " if (!el.shadowRoot) return null; el = el.shadowRoot.querySelector(arguments[$i]); if (!el) return null;"
        }
        $js += ' return el;'
        $deadline = (Get-Date).AddSeconds($Timeout)
        do {
            $found = $Driver.ExecuteScript($js, $parts)
            if ($found) { return $found }
            Start-Sleep -Milliseconds 300
        } while ((Get-Date) -lt $deadline)
        throw "Shadow DOM element not found (or shadow root is closed/inaccessible) for: $Selector"
    }
    elseif ($Selector -like '//*') {
        return Find-SeElement -Driver $Driver -XPath $Selector -Timeout $Timeout
    } else {
        return Find-SeElement -Driver $Driver -CssSelector $Selector -Timeout $Timeout
    }
}

# Clicks an element, falling back to a JavaScript click if the normal click
# is blocked by an overlay (cookie banner, floating widget, sticky header,
# etc.) - "element click intercepted" errors are common on real sites.
function Invoke-SeClickSafe {
    param($Driver, $Element)
    try {
        Invoke-SeClick -Element $Element
    } catch {
        $Driver.ExecuteScript("arguments[0].click();", $Element) | Out-Null
    }
}

# Takes a screenshot of the ENTIRE page, not just the visible viewport.
# Temporarily resizes the browser window to the page's full scroll height
# (works even headless, since Chrome allows arbitrary window sizes there),
# takes the shot, then restores the original window size. Falls back to a
# normal viewport screenshot if anything about the resize fails.
function Invoke-SeFullPageScreenshot {
    param($Driver, [string]$Path)
    try {
        $originalSize = $Driver.Manage().Window.Size
        $totalHeight = [int]$Driver.ExecuteScript("return Math.max(document.body.scrollHeight, document.documentElement.scrollHeight);")
        $totalWidth  = [int]$Driver.ExecuteScript("return Math.max(document.body.scrollWidth, document.documentElement.scrollWidth, window.innerWidth);")
        $newWidth  = [math]::Max($originalSize.Width, $totalWidth)
        $newHeight = [math]::Max($originalSize.Height, $totalHeight)
        $Driver.Manage().Window.Size = New-Object System.Drawing.Size($newWidth, $newHeight)
        Start-Sleep -Milliseconds 400   # let layout settle after resize
        Invoke-SeScreenshot -Driver $Driver | Save-SeScreenshot -Path $Path
        $Driver.Manage().Window.Size = $originalSize
    } catch {
        # Resize failed for some reason (e.g. a fixed/sticky-position layout
        # quirk) - fall back to a normal viewport screenshot rather than losing
        # the screenshot entirely.
        try { Invoke-SeScreenshot -Driver $Driver | Save-SeScreenshot -Path $Path } catch { }
    }
}

# Takes a screenshot of just the bottom of the page (scrolls to
# document.body.scrollHeight first, then a normal viewport screenshot) - use
# this when the footer itself, not the whole page, is the thing being
# checked. Returns "" (and logs nothing) if CaptureFooterScreenshot isn't
# set to "Yes" for this row, so calling it is safe/no-op by default.
function Invoke-SeFooterScreenshot {
    param($Driver, [string]$Path, [string]$CaptureFlag)
    if ($CaptureFlag -ne 'Yes') { return "" }
    try {
        $Driver.ExecuteScript("window.scrollTo(0, document.body.scrollHeight);") | Out-Null
        Start-Sleep -Milliseconds 500
        Invoke-SeScreenshot -Driver $Driver | Save-SeScreenshot -Path $Path
        $Driver.ExecuteScript("window.scrollTo(0, 0);") | Out-Null
        return $Path
    } catch {
        return ""
    }
}

foreach ($group in $AllRows | Group-Object SiteName) {

    $siteName = $group.Name
    $loginRow = $group.Group | Where-Object { $_.RowType -eq 'Login' } | Select-Object -First 1
    $pageRows = $group.Group | Where-Object { $_.RowType -eq 'Page' }

    if (-not $loginRow) {
        $Results.Add([PSCustomObject]@{
            Site=$siteName; Page="CONFIG"; URL=""; Status="Error"
            Error="No Login row found in config"; Timestamp=(Get-Date); Screenshot=""; DurationSec=0; CapturedInfo=""; FooterScreenshot=""
        })
        continue
    }

    if ($loginRow.Enabled -eq 'No') {
        Write-Host "  [$siteName] SKIPPED (disabled in config)" -ForegroundColor DarkGray
        $Results.Add([PSCustomObject]@{
            Site=$siteName; Page="SITE-LEVEL"; URL=$loginRow.Url; Status="Skipped"
            Error=""; Timestamp=(Get-Date); Screenshot=""; DurationSec=0; CapturedInfo=""; FooterScreenshot=""
        })
        continue
    }

    $driver = $null
        try {
            #Use this if Chrome is installed
            $chromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"    #Update this path based on the system this code is run.
            $driverDir  = "C:\Code\website-e2e-tester\drivers"
            $cred = Get-SiteCredential -SiteName $siteName

            $driver = if ($Headless) {
                Start-SeChrome -BinaryPath $chromePath -WebDriverDirectory $driverDir -Headless -Arguments @('--ignore-certificate-errors','--window-size=1400,1000')
            } else {
                Start-SeChrome -BinaryPath $chromePath -WebDriverDirectory $driverDir -Arguments @('--ignore-certificate-errors')
            }

        #Use this if Edge is installed
        #$driver = if ($Headless) {
        #    Start-SeEdge -Headless -Arguments @('--ignore-certificate-errors','--window-size=1400,1000')
        #} else {
        #    Start-SeEdge -Arguments @('--ignore-certificate-errors')
        #}

        # ---- LOGIN ----
        $loginStepStart = Get-Date
        Enter-SeUrl -Driver $driver -Url $loginRow.Url
        Start-Sleep -Seconds 2
        $preLoginShot = Join-Path $ScreenshotFolder "$siteName`_00_PreLogin.png"
        Invoke-SeFullPageScreenshot -Driver $driver -Path $preLoginShot

        if ($loginRow.IframeSelector) {
            $frameEl = Find-SeElementAny -Driver $driver -Selector $loginRow.IframeSelector -Timeout 15
            Switch-SeFrame -Frame $frameEl
        }

        $userField = Find-SeElementAny -Driver $driver -Selector $loginRow.UserFieldSelector -Timeout 15
        Send-SeKeys -Element $userField -Keys $cred.UserName

        if ($loginRow.ContinueButtonSelector) {
            # Two-step login: username -> Continue -> separate password screen
            $continueBtn = Find-SeElementAny -Driver $driver -Selector $loginRow.ContinueButtonSelector -Timeout 15
            Invoke-SeClickSafe -Driver $driver -Element $continueBtn
            Start-Sleep -Seconds 2
            $afterUserShot = Join-Path $ScreenshotFolder "$siteName`_00b_AfterUsername.png"
            Invoke-SeFullPageScreenshot -Driver $driver -Path $afterUserShot
        }

        $passField = Find-SeElementAny -Driver $driver -Selector $loginRow.PassFieldSelector -Timeout 15
        Send-SeKeys -Element $passField -Keys $cred.GetNetworkCredential().Password

        $loginBtn = Find-SeElementAny -Driver $driver -Selector $loginRow.LoginButtonSelector -Timeout 15
        Invoke-SeClickSafe -Driver $driver -Element $loginBtn
        Start-Sleep -Seconds 3

        if ($loginRow.IframeSelector) {
            Switch-SeFrame -Root   # back to the main document - login normally navigates the outer page away from the iframe
        }

        if ($loginRow.RequiresMFA -eq 'Yes') {
            if ($Headless) {
                Write-Host "  [$siteName] WARNING: RequiresMFA is Yes but running headless - there's no visible browser to complete MFA in. Re-run with -Headless:`$false for this site." -ForegroundColor Yellow
            } else {
                Write-Host "`n  [$siteName] MFA required - complete it now in the browser window." -ForegroundColor Yellow
                Read-Host "  Press Enter once you've completed MFA and are logged in"
            }
        }

        $postLoginEl = $null
        if ($loginRow.PostLoginCheckSelector) {
            $postLoginEl = Find-SeElementAny -Driver $driver -Selector $loginRow.PostLoginCheckSelector -Timeout 10 -ErrorAction SilentlyContinue
        }
        $loginStatus = if ($postLoginEl) { "Success" } else { "Failure" }

        # Optional: capture (not assert) a footer/server tag on the post-login
        # page, e.g. "S09" - value legitimately varies run to run. Captured
        # BEFORE the screenshots below so the tag can be folded into their
        # filenames (e.g. "SiteA_00_Login_S09.png").
        $loginCapturedInfo = Get-SeCapturedText -Driver $driver -Selector $loginRow.FooterCaptureSelector
        $loginTag = Get-CaptureFileTag -Text $loginCapturedInfo
        $loginTagSuffix = if ($loginTag) { "_$loginTag" } else { "" }

        $loginShot = Join-Path $ScreenshotFolder "$siteName`_00_Login$loginTagSuffix.png"
        Invoke-SeFullPageScreenshot -Driver $driver -Path $loginShot

        # Optional: a dedicated footer screenshot as its own test step
        # (separate from the full-page screenshot above), when the footer
        # itself - not the whole page - is what's being checked.
        $loginFooterShotPath = Join-Path $ScreenshotFolder "$siteName`_00_Login_Footer$loginTagSuffix.png"
        $loginFooterShot = Invoke-SeFooterScreenshot -Driver $driver -Path $loginFooterShotPath -CaptureFlag $loginRow.CaptureFooterScreenshot
        $loginDuration = [math]::Round(((Get-Date) - $loginStepStart).TotalSeconds, 1)
        Write-Host "  [$siteName] LOGIN - $loginStatus ($loginDuration sec)"

        $Results.Add([PSCustomObject]@{
            Site=$siteName; Page="LOGIN"; URL=$loginRow.Url; Status=$loginStatus
            Error=$(if ($loginStatus -eq 'Failure') { "Post-login element not found: $($loginRow.PostLoginCheckSelector)" } else { "" })
            Timestamp=(Get-Date); Screenshot=$loginShot; DurationSec=$loginDuration; CapturedInfo=$loginCapturedInfo; FooterScreenshot=$loginFooterShot
        })

        if ($loginStatus -eq "Failure") { continue }   # skip pages if login didn't work

        # ---- PAGES ----
        foreach ($page in $pageRows) {
            $status = "Success"; $errMsg = ""; $capturedInfo = ""; $footerShot = ""
            $safeName = ($page.PageName -replace '[^\w\-]', '_')
            # Default filenames (no captured tag yet - used as a fallback if
            # something throws before the capture below runs, e.g. the page
            # itself fails to load).
            $shotPath = Join-Path $ScreenshotFolder "$siteName`_$safeName.png"
            $footerShotPath = Join-Path $ScreenshotFolder "$siteName`_$safeName`_Footer.png"
            $pageStepStart = Get-Date

            try {
                Enter-SeUrl -Driver $driver -Url $page.PageUrl
                Start-Sleep -Seconds 2

                if ($page.ExpectedHeader -and $page.HeaderSelector) {
                    $headerEl = Find-SeElementAny -Driver $driver -Selector $page.HeaderSelector -Timeout 10
                    if ($headerEl.Text.Trim() -notmatch [regex]::Escape($page.ExpectedHeader.Trim())) {
                        $status = "Failure"
                        $errMsg = "Header mismatch: expected '$($page.ExpectedHeader)', found '$($headerEl.Text)'"
                    }
                }

                if ($status -eq "Success" -and $page.ExpectedSubHeader -and $page.SubHeaderSelector) {
                    $subEl = Find-SeElementAny -Driver $driver -Selector $page.SubHeaderSelector -Timeout 10
                    if ($subEl.Text.Trim() -notmatch [regex]::Escape($page.ExpectedSubHeader.Trim())) {
                        $status = "Failure"
                        $errMsg = "Subheader mismatch: expected '$($page.ExpectedSubHeader)', found '$($subEl.Text)'"
                    }
                }

                # Optional: capture (not assert) footer/server-tag text such as
                # "S09"/"Sxx" - the value legitimately varies run to run and
                # site to site, so it's logged for reference, never pass/failed.
                # Captured BEFORE the screenshots below so the tag can be
                # folded into their filenames (e.g. "SiteA_Dashboard_S09.png").
                $capturedInfo = Get-SeCapturedText -Driver $driver -Selector $page.FooterCaptureSelector
                $pageTag = Get-CaptureFileTag -Text $capturedInfo
                if ($pageTag) {
                    $shotPath = Join-Path $ScreenshotFolder "$siteName`_$safeName`_$pageTag.png"
                    $footerShotPath = Join-Path $ScreenshotFolder "$siteName`_$safeName`_Footer_$pageTag.png"
                }

                Invoke-SeFullPageScreenshot -Driver $driver -Path $shotPath

                # Optional: a dedicated footer screenshot as its own test
                # step (in addition to the full-page shot above), for cases
                # where the footer itself is the thing being checked.
                $footerShot = Invoke-SeFooterScreenshot -Driver $driver -Path $footerShotPath -CaptureFlag $page.CaptureFooterScreenshot
            }
            catch {
                $status = "Error"
                $errMsg = $_.Exception.Message
                Invoke-SeFullPageScreenshot -Driver $driver -Path $shotPath
            }

            $pageDuration = [math]::Round(((Get-Date) - $pageStepStart).TotalSeconds, 1)
            $capturedSuffix = if ($capturedInfo) { " - [$capturedInfo]" } else { "" }
            Write-Host "  [$siteName] $($page.PageName) - $status ($pageDuration sec) - $($page.PageUrl)$capturedSuffix"

            $Results.Add([PSCustomObject]@{
                Site=$siteName; Page=$page.PageName; URL=$page.PageUrl; Status=$status
                Error=$errMsg; Timestamp=(Get-Date); Screenshot=$shotPath; DurationSec=$pageDuration; CapturedInfo=$capturedInfo; FooterScreenshot=$footerShot
            })
        }

        # ---- LOGOUT ----
        if ($loginRow.LogoutSelector) {
            $logoutStepStart = Get-Date
            try {
                if ($loginRow.LogoutPageUrl) {
                    # Navigate to a known page that actually has the logout
                    # control, rather than assuming whatever page the last
                    # "Page" row left the browser on has it too (it might not -
                    # e.g. a standalone widget/settings page with no nav bar).
                    Enter-SeUrl -Driver $driver -Url $loginRow.LogoutPageUrl
                    Start-Sleep -Seconds 2
                }
                if ($loginRow.LogoutMenuSelector) {
                    $menuBtn = Find-SeElementAny -Driver $driver -Selector $loginRow.LogoutMenuSelector -Timeout 10
                    Invoke-SeClickSafe -Driver $driver -Element $menuBtn
                    Start-Sleep -Seconds 1
                }
                $logoutBtn = Find-SeElementAny -Driver $driver -Selector $loginRow.LogoutSelector -Timeout 10
                Invoke-SeClickSafe -Driver $driver -Element $logoutBtn
                Start-Sleep -Seconds 2

                $logoutShot = Join-Path $ScreenshotFolder "$siteName`_99_Logout.png"
                Invoke-SeFullPageScreenshot -Driver $driver -Path $logoutShot
                $logoutDuration = [math]::Round(((Get-Date) - $logoutStepStart).TotalSeconds, 1)
                Write-Host "  [$siteName] LOGOUT - Success ($logoutDuration sec)"

                $Results.Add([PSCustomObject]@{
                    Site=$siteName; Page="LOGOUT"; URL=$driver.Url; Status="Success"
                    Error=""; Timestamp=(Get-Date); Screenshot=$logoutShot; DurationSec=$logoutDuration; CapturedInfo=""; FooterScreenshot=""
                })
            } catch {
                $logoutDuration = [math]::Round(((Get-Date) - $logoutStepStart).TotalSeconds, 1)
                $Results.Add([PSCustomObject]@{
                    Site=$siteName; Page="LOGOUT"; URL=""; Status="Error"
                    Error=$_.Exception.Message; Timestamp=(Get-Date); Screenshot=""; DurationSec=$logoutDuration; CapturedInfo=""; FooterScreenshot=""
                })
            }
        }

    }
    catch {
        $Results.Add([PSCustomObject]@{
            Site=$siteName; Page="SITE-LEVEL"; URL=$loginRow.Url; Status="Error"
            Error=$_.Exception.Message; Timestamp=(Get-Date); Screenshot=""; DurationSec=0; CapturedInfo=""; FooterScreenshot=""
        })
    }
    finally {
        if ($driver) { Stop-SeDriver -Driver $driver }
    }
}

# ================= EXCEL REPORT =================
$ExcelPath = Join-Path $OutputFolder "TestResults.xlsx"
$Results |
    Select-Object Site, Page, URL, Status, Timestamp, DurationSec, CapturedInfo,
        @{Name='FooterScreenshotCaptured'; Expression={ if ($_.FooterScreenshot) { 'Yes' } else { '' } }},
        Error |
    Export-Excel -Path $ExcelPath -WorksheetName "Results" -TableName "Results" `
                  -AutoSize -FreezeTopRow -BoldTopRow `
                  -ConditionalText @(
                      New-ConditionalText -Text "Success" -BackgroundColor LightGreen
                      New-ConditionalText -Text "Failure" -BackgroundColor LightYellow
                      New-ConditionalText -Text "Error"   -BackgroundColor LightPink
                      New-ConditionalText -Text "Skipped" -BackgroundColor LightGray
                  )

# ================= WORD REPORT =================
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0   # wdAlertsNone - suppress any popup that would otherwise hang silently
$doc = $word.Documents.Add()
$sel = $word.Selection

$sel.Style = "Title"
$sel.TypeText("Website End-to-End Test Report")
$sel.TypeParagraph()
$sel.Style = "Normal"
$sel.TypeText("Run: $(Get-Date -Format 'yyyy-MM-dd HH:mm')  |  Total duration: $([math]::Round(((Get-Date) - $ScriptStartTime).TotalMinutes, 1)) min")
$sel.TypeParagraph(); $sel.TypeParagraph()

foreach ($siteGroup in $Results | Group-Object Site) {
    $sel.Style = "Heading 1"
    $sel.TypeText($siteGroup.Name)
    $sel.TypeParagraph()

    foreach ($r in $siteGroup.Group) {
        $sel.Style = "Heading 2"
        $sel.TypeText("$($r.Page)  -  $($r.Status)")
        $sel.TypeParagraph()

        $sel.Style = "Normal"
        $sel.TypeText("URL: $($r.URL)  |  Duration: $($r.DurationSec) sec")
        $sel.TypeParagraph()

        if ($r.CapturedInfo) {
            $sel.TypeText("Captured info (e.g. server/footer tag): $($r.CapturedInfo)")
            $sel.TypeParagraph()
        }

        if ($r.Error) {
            $sel.Font.Color = 255            # red-ish
            $sel.TypeText("Issue: $($r.Error)")
            $sel.Font.Color = 0
            $sel.TypeParagraph()
        }

        if ($r.Screenshot -and (Test-Path $r.Screenshot)) {
            $sel.InlineShapes.AddPicture($r.Screenshot) | Out-Null
            $sel.TypeParagraph()
        }

        if ($r.FooterScreenshot -and (Test-Path $r.FooterScreenshot)) {
            $sel.Style = "Normal"
            $sel.TypeText("Footer screenshot:")
            $sel.TypeParagraph()
            $sel.InlineShapes.AddPicture($r.FooterScreenshot) | Out-Null
            $sel.TypeParagraph()
        }
        $sel.TypeParagraph()
    }
}

$docxPath = Join-Path (Resolve-Path $OutputFolder) "TestReport.docx"
$doc.SaveAs([ref]([string]$docxPath), [ref]([int]16))   # 16 = wdFormatDocumentDefault (.docx)
$doc.Close()
$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null

$ScriptEndTime = Get-Date
$TotalElapsed = $ScriptEndTime - $ScriptStartTime

$successCount = ($Results | Where-Object Status -eq 'Success').Count
$failureCount = ($Results | Where-Object Status -eq 'Failure').Count
$errorCount   = ($Results | Where-Object Status -eq 'Error').Count
$skippedCount = ($Results | Where-Object Status -eq 'Skipped').Count

# ================= ROI SUMMARY (business-value estimate) =================
# Rough, assumption-based estimate of manual effort this run replaced -
# never a measured fact, just $ManualMinutesPerCheck x checks vs. actual
# automated time. Written as its own worksheet so it doesn't clutter the
# Results sheet, and projected out to monthly/annual using $RunsPerDay.
$checksCount          = ($Results | Where-Object { $_.Status -ne 'Skipped' }).Count
$manualEquivMinutes   = [math]::Round($checksCount * $ManualMinutesPerCheck, 1)
$automatedMinutes     = [math]::Round($TotalElapsed.TotalMinutes, 1)
$timeSavedMinutes     = [math]::Round($manualEquivMinutes - $automatedMinutes, 1)
$percentSaved         = if ($manualEquivMinutes -gt 0) { [math]::Round(($timeSavedMinutes / $manualEquivMinutes) * 100, 0) } else { 0 }
$dailySavedMinutes    = [math]::Round($timeSavedMinutes * $RunsPerDay, 1)
$dailySavedHours      = [math]::Round($dailySavedMinutes / 60, 1)
$monthlySavedHours    = [math]::Round($dailySavedMinutes * 22 / 60, 1)    # ~22 working days/month
$annualSavedHours     = [math]::Round($dailySavedMinutes * 260 / 60, 1)   # ~260 working days/year

$roiRows = @(
    [PSCustomObject]@{ Metric = "Checks performed this run (excl. Skipped)";               Value = $checksCount }
    [PSCustomObject]@{ Metric = "Assumed manual time per check (min)";                     Value = $ManualMinutesPerCheck }
    [PSCustomObject]@{ Metric = "Manual-equivalent time for this run (min)";                Value = $manualEquivMinutes }
    [PSCustomObject]@{ Metric = "Actual automated time for this run (min)";                 Value = $automatedMinutes }
    [PSCustomObject]@{ Metric = "Time saved this run (min)";                                Value = $timeSavedMinutes }
    [PSCustomObject]@{ Metric = "Time saved this run (%)";                                  Value = "$percentSaved%" }
    [PSCustomObject]@{ Metric = "Assumed runs per day";                                     Value = $RunsPerDay }
    [PSCustomObject]@{ Metric = "Time saved per day (hrs)";                                 Value = $dailySavedHours }
    [PSCustomObject]@{ Metric = "Time saved per month - approx. 22 working days (hrs)";     Value = $monthlySavedHours }
    [PSCustomObject]@{ Metric = "Time saved per year - approx. 260 working days (hrs)";     Value = $annualSavedHours }
)

if ($HourlyRate -gt 0) {
    $costSavedRun   = [math]::Round(($timeSavedMinutes / 60) * $HourlyRate, 2)
    $costSavedMonth = [math]::Round($monthlySavedHours * $HourlyRate, 2)
    $costSavedYear  = [math]::Round($annualSavedHours * $HourlyRate, 2)
    $roiRows += [PSCustomObject]@{ Metric = "Assumed hourly rate ($CurrencySymbol)";        Value = $HourlyRate }
    $roiRows += [PSCustomObject]@{ Metric = "Cost saved this run ($CurrencySymbol)";        Value = $costSavedRun }
    $roiRows += [PSCustomObject]@{ Metric = "Cost saved per month ($CurrencySymbol)";       Value = $costSavedMonth }
    $roiRows += [PSCustomObject]@{ Metric = "Cost saved per year ($CurrencySymbol)";        Value = $costSavedYear }
}

$roiRows | Export-Excel -Path $ExcelPath -WorksheetName "ROI Summary" -AutoSize -BoldTopRow -FreezeTopRow

Write-Host "Done."
Write-Host "Started      : $($ScriptStartTime.ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Host "Finished     : $($ScriptEndTime.ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Host "Total time   : $($TotalElapsed.ToString('hh\:mm\:ss'))"
Write-Host "Excel report : $ExcelPath"
Write-Host "Word report  : $docxPath"
Write-Host "Screenshots  : $ScreenshotFolder"

if ($timeSavedMinutes -ge 0) {
    Write-Host "ROI (est.)   : ~$timeSavedMinutes min saved this run ($percentSaved%) vs. manual - projected ~$monthlySavedHours hrs/month at $RunsPerDay runs/day (see 'ROI Summary' tab in Excel)" -ForegroundColor Cyan
} else {
    Write-Host "ROI (est.)   : automated run took longer than the assumed manual estimate this time - check -ManualMinutesPerCheck / -RunsPerDay against reality (see 'ROI Summary' tab)" -ForegroundColor Yellow
}

$summaryColor = if ($failureCount -gt 0 -or $errorCount -gt 0) { 'Red' } else { 'Green' }
$summaryLine = "Summary: $successCount Success, $failureCount Failure, $errorCount Error"
if ($skippedCount -gt 0) { $summaryLine += ", $skippedCount Skipped" }
Write-Host $summaryLine -ForegroundColor $summaryColor

# Non-zero exit code whenever anything failed or errored, so a Task Scheduler
# job or CI pipeline can detect a bad run automatically without opening the
# Excel report.
if ($failureCount -gt 0 -or $errorCount -gt 0) {
    exit 1
} else {
    exit 0
}
