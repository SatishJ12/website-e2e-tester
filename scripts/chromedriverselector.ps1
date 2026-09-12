$chromeVersion = "152.0.7977.83"
$json = Invoke-RestMethod -Uri "https://googlechromelabs.github.io/chrome-for-testing/known-good-versions-with-downloads.json"

# Try an exact match first
$match = $json.versions | Where-Object { $_.version -eq $chromeVersion }

# If no exact match, fall back to the closest version sharing the same build number (7977)
if (-not $match) {
    $match = $json.versions | Where-Object { $_.version -like "152.0.7977.*" } | Select-Object -Last 1
}

$driverUrl = ($match.downloads.chromedriver | Where-Object { $_.platform -eq 'win64' }).url
Write-Host "Using chromedriver version: $($match.version)"
Write-Host "Download URL: $driverUrl"

# Download and extract
New-Item -ItemType Directory -Path C:\Code\website-e2e-tester\drivers -Force | Out-Null
Invoke-WebRequest -Uri $driverUrl -OutFile "$env:TEMP\chromedriver.zip"
Expand-Archive -Path "$env:TEMP\chromedriver.zip" -DestinationPath "$env:TEMP\chromedriver-extract" -Force
Copy-Item "$env:TEMP\chromedriver-extract\chromedriver-win64\chromedriver.exe" -Destination "C:\Code\website-e2e-tester\drivers\chromedriver.exe" -Force