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
    First run per site will prompt for credentials once, then reuse them
    (encrypted, tied to this Windows user + machine) on every future run.
#>

param(
    [string]$ConfigPath   = "$PSScriptRoot\..\config\Site-Config.csv",
    [string]$CredFolder   = "$PSScriptRoot\..\Creds",
    [string]$OutputFolder = "$PSScriptRoot\..\TestRun_$(Get-Date -Format yyyyMMdd_HHmmss)",
    [switch]$Headless = $true
)

$ScriptStartTime = Get-Date
Write-Host "Started: $($ScriptStartTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Cyan

Import-Module Selenium -ErrorAction Stop
Import-Module ImportExcel -ErrorAction Stop

New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
New-Item -ItemType Directory -Path $CredFolder   -Force | Out-Null
$ScreenshotFolder = Join-Path $OutputFolder "Screenshots"
New-Item -ItemType Directory -Path $ScreenshotFolder -Force | Out-Null

if (-not (Test-Path $ConfigPath)) {
    Write-Error "Config file not found: $ConfigPath"
    exit 1
}
$AllRows = Import-Csv $ConfigPath
$Results = New-Object System.Collections.Generic.List[object]

function Get-SiteCredential {
    param([string]$SiteName)
    $path = Join-Path $CredFolder "$SiteName.xml"
    if (Test-Path $path) {
        return Import-Clixml $path
    }
    Write-Host "`nEnter login credentials for $SiteName" -ForegroundColor Cyan
    $username = Read-Host "Username"
    $securePassword = Read-Host "Password" -AsSecureString
    $cred = New-Object System.Management.Automation.PSCredential($username, $securePassword)
    $cred | Export-Clixml $path
    return $cred
}

# Finds an element using either a CSS selector or an XPath expression.
# A selector value starting with "//" is treated as XPath (needed for elements
# only identifiable by visible text, e.g. a logout link with no id/class);
# everything else is treated as a normal CSS selector.
function Find-SeElementAny {
    param($Driver, [string]$Selector, [double]$Timeout = 10)
    if ($Selector -like '//*') {
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

foreach ($group in $AllRows | Group-Object SiteName) {

    $siteName = $group.Name
    $loginRow = $group.Group | Where-Object { $_.RowType -eq 'Login' } | Select-Object -First 1
    $pageRows = $group.Group | Where-Object { $_.RowType -eq 'Page' }

    if (-not $loginRow) {
        $Results.Add([PSCustomObject]@{
            Site=$siteName; Page="CONFIG"; URL=""; Status="Error"
            Error="No Login row found in config"; Timestamp=(Get-Date); Screenshot=""; DurationSec=0
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
        Invoke-SeScreenshot -Driver $driver | Save-SeScreenshot -Path $preLoginShot

        $userField = Find-SeElementAny -Driver $driver -Selector $loginRow.UserFieldSelector -Timeout 15
        Send-SeKeys -Element $userField -Keys $cred.UserName

        $passField = Find-SeElementAny -Driver $driver -Selector $loginRow.PassFieldSelector -Timeout 15
        Send-SeKeys -Element $passField -Keys $cred.GetNetworkCredential().Password

        $loginBtn = Find-SeElementAny -Driver $driver -Selector $loginRow.LoginButtonSelector -Timeout 15
        Invoke-SeClickSafe -Driver $driver -Element $loginBtn
        Start-Sleep -Seconds 3

        $postLoginEl = $null
        if ($loginRow.PostLoginCheckSelector) {
            $postLoginEl = Find-SeElementAny -Driver $driver -Selector $loginRow.PostLoginCheckSelector -Timeout 10 -ErrorAction SilentlyContinue
        }
        $loginStatus = if ($postLoginEl) { "Success" } else { "Failure" }
        $loginShot = Join-Path $ScreenshotFolder "$siteName`_00_Login.png"
        Invoke-SeScreenshot -Driver $driver | Save-SeScreenshot -Path $loginShot
        $loginDuration = [math]::Round(((Get-Date) - $loginStepStart).TotalSeconds, 1)
        Write-Host "  [$siteName] LOGIN - $loginStatus ($loginDuration sec)"

        $Results.Add([PSCustomObject]@{
            Site=$siteName; Page="LOGIN"; URL=$loginRow.Url; Status=$loginStatus
            Error=$(if ($loginStatus -eq 'Failure') { "Post-login element not found: $($loginRow.PostLoginCheckSelector)" } else { "" })
            Timestamp=(Get-Date); Screenshot=$loginShot; DurationSec=$loginDuration
        })

        if ($loginStatus -eq "Failure") { continue }   # skip pages if login didn't work

        # ---- PAGES ----
        foreach ($page in $pageRows) {
            $status = "Success"; $errMsg = ""
            $safeName = ($page.PageName -replace '[^\w\-]', '_')
            $shotPath = Join-Path $ScreenshotFolder "$siteName`_$safeName.png"
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

                Invoke-SeScreenshot -Driver $driver | Save-SeScreenshot -Path $shotPath
            }
            catch {
                $status = "Error"
                $errMsg = $_.Exception.Message
                try { Invoke-SeScreenshot -Driver $driver | Save-SeScreenshot -Path $shotPath } catch { }
            }

            $pageDuration = [math]::Round(((Get-Date) - $pageStepStart).TotalSeconds, 1)
            Write-Host "  [$siteName] $($page.PageName) - $status ($pageDuration sec) - $($page.PageUrl)"

            $Results.Add([PSCustomObject]@{
                Site=$siteName; Page=$page.PageName; URL=$page.PageUrl; Status=$status
                Error=$errMsg; Timestamp=(Get-Date); Screenshot=$shotPath; DurationSec=$pageDuration
            })
        }

        # ---- LOGOUT ----
        if ($loginRow.LogoutSelector) {
            $logoutStepStart = Get-Date
            try {
                if ($loginRow.LogoutMenuSelector) {
                    $menuBtn = Find-SeElementAny -Driver $driver -Selector $loginRow.LogoutMenuSelector -Timeout 10
                    Invoke-SeClickSafe -Driver $driver -Element $menuBtn
                    Start-Sleep -Seconds 1
                }
                $logoutBtn = Find-SeElementAny -Driver $driver -Selector $loginRow.LogoutSelector -Timeout 10
                Invoke-SeClickSafe -Driver $driver -Element $logoutBtn
                Start-Sleep -Seconds 2

                $logoutShot = Join-Path $ScreenshotFolder "$siteName`_99_Logout.png"
                Invoke-SeScreenshot -Driver $driver | Save-SeScreenshot -Path $logoutShot
                $logoutDuration = [math]::Round(((Get-Date) - $logoutStepStart).TotalSeconds, 1)
                Write-Host "  [$siteName] LOGOUT - Success ($logoutDuration sec)"

                $Results.Add([PSCustomObject]@{
                    Site=$siteName; Page="LOGOUT"; URL=$driver.Url; Status="Success"
                    Error=""; Timestamp=(Get-Date); Screenshot=$logoutShot; DurationSec=$logoutDuration
                })
            } catch {
                $logoutDuration = [math]::Round(((Get-Date) - $logoutStepStart).TotalSeconds, 1)
                $Results.Add([PSCustomObject]@{
                    Site=$siteName; Page="LOGOUT"; URL=""; Status="Error"
                    Error=$_.Exception.Message; Timestamp=(Get-Date); Screenshot=""; DurationSec=$logoutDuration
                })
            }
        }

    }
    catch {
        $Results.Add([PSCustomObject]@{
            Site=$siteName; Page="SITE-LEVEL"; URL=$loginRow.Url; Status="Error"
            Error=$_.Exception.Message; Timestamp=(Get-Date); Screenshot=""; DurationSec=0
        })
    }
    finally {
        if ($driver) { Stop-SeDriver -Driver $driver }
    }
}

# ================= EXCEL REPORT =================
$ExcelPath = Join-Path $OutputFolder "TestResults.xlsx"
$Results |
    Select-Object Site, Page, URL, Status, Timestamp, DurationSec, Error |
    Export-Excel -Path $ExcelPath -WorksheetName "Results" -TableName "Results" `
                  -AutoSize -FreezeTopRow -BoldTopRow `
                  -ConditionalText @(
                      New-ConditionalText -Text "Success" -BackgroundColor LightGreen
                      New-ConditionalText -Text "Failure" -BackgroundColor LightYellow
                      New-ConditionalText -Text "Error"   -BackgroundColor LightPink
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

Write-Host "Done."
Write-Host "Started      : $($ScriptStartTime.ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Host "Finished     : $($ScriptEndTime.ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Host "Total time   : $($TotalElapsed.ToString('hh\:mm\:ss'))"
Write-Host "Excel report : $ExcelPath"
Write-Host "Word report  : $docxPath"
Write-Host "Screenshots  : $ScreenshotFolder"
