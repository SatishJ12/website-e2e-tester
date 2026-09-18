<#
.SYNOPSIS
    Keeps drivers\chromedriver.exe in sync with your installed Chrome version.
    Auto-detects Chrome (no hardcoded version), skips the download entirely
    if the driver you already have already matches, and only hits the
    network when a real mismatch is found.

.USAGE
    .\Update-ChromeDriver.ps1
    .\Update-ChromeDriver.ps1 -DriverFolder "C:\Code\website-e2e-tester\drivers"
#>

param(
    [string]$DriverFolder = "C:\Code\website-e2e-tester\drivers"
)

# Force TLS 1.2 - Windows PowerShell 5.1 sometimes defaults to an older
# protocol that this API (and PSGallery, for the same reason) will reject.
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# ---- 1. Find installed Chrome and its version ----
$chromeCandidates = @(
    "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe",
    "C:\Program Files\Google\Chrome\Application\chrome.exe",
    "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
)
$chromePath = $chromeCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $chromePath) {
    Write-Error "Could not find Chrome in any of the usual locations: $($chromeCandidates -join ', ')"
    exit 1
}
$chromeVersion = (Get-Item $chromePath).VersionInfo.ProductVersion
Write-Host "Installed Chrome : $chromeVersion  ($chromePath)"

# ---- 2. Check the driver you already have - skip everything if it already matches ----
$driverPath = Join-Path $DriverFolder "chromedriver.exe"
if (Test-Path $driverPath) {
    $driverVersion = (Get-Item $driverPath).VersionInfo.ProductVersion
    Write-Host "Current driver   : $driverVersion  ($driverPath)"
    if ($driverVersion -eq $chromeVersion) {
        Write-Host "Already matches - nothing to download." -ForegroundColor Green
        exit 0
    }
    Write-Host "Mismatch - downloading a matching chromedriver..." -ForegroundColor Yellow
} else {
    Write-Host "No chromedriver.exe found in $DriverFolder - downloading..." -ForegroundColor Yellow
}

# ---- 3. Look up and download the matching chromedriver ----
try {
    $json = Invoke-RestMethod -Uri "https://googlechromelabs.github.io/chrome-for-testing/known-good-versions-with-downloads.json" -ErrorAction Stop
} catch {
    Write-Error "Could not reach the Chrome-for-Testing version API (googlechromelabs.github.io). Check your internet connection, or whether a firewall/proxy is blocking that domain. Details: $($_.Exception.Message)"
    exit 1
}

$match = $json.versions | Where-Object { $_.version -eq $chromeVersion }
if (-not $match) {
    # No exact build published for this patch - fall back to the closest
    # version sharing the same major.minor.build (e.g. "152.0.7977.*")
    $buildPrefix = ($chromeVersion -split '\.')[0..2] -join '.'
    $match = $json.versions | Where-Object { $_.version -like "$buildPrefix.*" } | Select-Object -Last 1
}
if (-not $match) {
    Write-Error "No matching chromedriver found for Chrome version $chromeVersion in the Chrome-for-Testing catalog."
    exit 1
}

$driverUrl = ($match.downloads.chromedriver | Where-Object { $_.platform -eq 'win64' }).url
if (-not $driverUrl) {
    Write-Error "Chrome-for-Testing catalog has an entry for $($match.version) but no win64 chromedriver download - this shouldn't normally happen."
    exit 1
}
Write-Host "Using chromedriver: $($match.version)"

New-Item -ItemType Directory -Path $DriverFolder -Force | Out-Null
$zipPath     = Join-Path $env:TEMP "chromedriver.zip"
$extractPath = Join-Path $env:TEMP "chromedriver-extract"

try {
    Invoke-WebRequest -Uri $driverUrl -OutFile $zipPath -ErrorAction Stop
} catch {
    Write-Error "Failed to download chromedriver from $driverUrl. Check your internet connection or proxy settings. Details: $($_.Exception.Message)"
    exit 1
}

try {
    Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force -ErrorAction Stop

    # If a previous test run crashed and left chromedriver.exe running, Copy-Item
    # below would fail with "file in use". Stop it first - no admin rights needed,
    # since any user can terminate their own processes, and this only ever
    # targets a chromedriver process running from this exact file path.
    Get-Process -Name chromedriver -ErrorAction SilentlyContinue |
        Where-Object { $_.Path -eq $driverPath } |
        Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 300

    Copy-Item (Join-Path $extractPath "chromedriver-win64\chromedriver.exe") -Destination $driverPath -Force -ErrorAction Stop
} catch {
    Write-Error "Downloaded chromedriver but couldn't extract/install it (file may be corrupted, in use by a running process, or the zip's internal folder structure changed). Details: $($_.Exception.Message)"
    exit 1
}

# Clean up the temp zip/extract so they don't pile up on repeated runs
Remove-Item $zipPath -ErrorAction SilentlyContinue
Remove-Item $extractPath -Recurse -ErrorAction SilentlyContinue

Write-Host "Done - chromedriver $($match.version) is now at $driverPath" -ForegroundColor Green
