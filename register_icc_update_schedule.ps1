param(
    [string]$TaskName = "EFC LSS ICC Update",
    [string]$Url = $env:ICC_URL,
    [string[]]$Times = @("08:00", "13:00"),
    [bool]$Headless = $true,
    [bool]$Deploy = $true
)

$ErrorActionPreference = "Stop"
Set-Location -Path $PSScriptRoot

$scriptPath = Join-Path $PSScriptRoot "run_icc_daily_update.ps1"
if (-not (Test-Path -LiteralPath $scriptPath)) {
    throw "Update script not found: $scriptPath"
}

$taskArgs = @(
    "-NoProfile",
    "-ExecutionPolicy", "Bypass",
    "-File", ('"{0}"' -f $scriptPath)
)

if ($Url) {
    $taskArgs += @("-Url", ('"{0}"' -f $Url))
}
if ($Headless) {
    $taskArgs += "-Headless"
}
if ($Deploy) {
    $taskArgs += "-Deploy"
}

$action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument ($taskArgs -join " ") `
    -WorkingDirectory $PSScriptRoot

$triggers = foreach ($time in $Times) {
    New-ScheduledTaskTrigger -Daily -At $time
}

$settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Hours 2)

$user = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$principal = New-ScheduledTaskPrincipal `
    -UserId $user `
    -LogonType Interactive `
    -RunLevel Limited

$existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($existing) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
}

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $action `
    -Trigger $triggers `
    -Settings $settings `
    -Principal $principal `
    -Description "Updates the EFC/LSS ICC dashboard at 08:00 and 13:00, then deploys changed data to GitHub Pages." | Out-Null

Write-Host "Registered scheduled task: $TaskName"
Write-Host ("Times: {0}" -f ($Times -join ", "))
Write-Host ("Action: powershell.exe {0}" -f ($taskArgs -join " "))
if (-not $Url) {
    Write-Warning "ICC_URL is not set. The task is registered, but ICC download needs ICC_URL or a reusable browser session."
}
