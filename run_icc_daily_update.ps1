param(
    [string]$Url = $env:ICC_URL,
    [switch]$Headless,
    [string]$DownloadFile
)

$ErrorActionPreference = "Stop"
Set-Location -Path $PSScriptRoot

$arguments = @(".\icc_daily_update.py")

if ($Url) {
    $arguments += @("--url", $Url)
}

if ($Headless) {
    $arguments += "--headless"
}

if ($DownloadFile) {
    $arguments += @("--download-file", $DownloadFile)
}

py @arguments
