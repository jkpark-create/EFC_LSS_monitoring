param(
    [string]$Url = $env:ICC_URL,
    [switch]$Headless,
    [switch]$XPlatform,
    [string]$DownloadFile,
    [switch]$Deploy,
    [switch]$NoLog,
    [int]$XPlatformAttempts = 3
)

$ErrorActionPreference = "Stop"
Set-Location -Path $PSScriptRoot

$logPath = $null
if (-not $NoLog) {
    $logDir = Join-Path $PSScriptRoot "logs"
    New-Item -ItemType Directory -Force -Path $logDir | Out-Null
    $logPath = Join-Path $logDir ("icc_daily_update_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
    Start-Transcript -Path $logPath -Append | Out-Null
}

try {
    Write-Host ("ICC daily update started at {0}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"))

    if ($XPlatform -and -not $DownloadFile) {
        $downloadDir = Join-Path $PSScriptRoot "downloads"
        New-Item -ItemType Directory -Force -Path $downloadDir | Out-Null
        $DownloadFile = Join-Path $downloadDir ("xplatform_DynamicList_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

        Write-Host "Downloading ICC data through XPlatform."
        for ($attempt = 1; $attempt -le $XPlatformAttempts; $attempt++) {
            $attemptFile = $DownloadFile
            if ($attempt -gt 1) {
                $baseName = [System.IO.Path]::GetFileNameWithoutExtension($DownloadFile)
                $extension = [System.IO.Path]::GetExtension($DownloadFile)
                $attemptFile = Join-Path $downloadDir ("{0}_attempt{1}{2}" -f $baseName, $attempt, $extension)
            }

            $xplatformArguments = @(
                ".\xplatform_icc_helper.py",
                "download",
                "--launch-timeout", "180",
                "--search-wait", "60",
                "--export-timeout", "180",
                "--export-attempts", "3",
                "--output-file", $attemptFile
            )

            Write-Host ("XPlatform download attempt {0}/{1}." -f $attempt, $XPlatformAttempts)
            py @xplatformArguments
            if ($LASTEXITCODE -eq 0) {
                $DownloadFile = $attemptFile
                break
            }

            $lastExitCode = $LASTEXITCODE
            Get-Process XPlatform -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
            if ($attempt -lt $XPlatformAttempts) {
                Write-Warning ("XPlatform download failed with exit code {0}; retrying after cleanup." -f $lastExitCode)
                Start-Sleep -Seconds 15
            } else {
                throw "xplatform_icc_helper.py failed after $XPlatformAttempts attempts; last exit code $lastExitCode"
            }
        }
    }

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
    if ($LASTEXITCODE -ne 0) {
        throw "icc_daily_update.py failed with exit code $LASTEXITCODE"
    }

    if ($Deploy) {
        Write-Host "Deploy option enabled. Committing and pushing dashboard data when changed."

        git add index.html data.json
        if ($LASTEXITCODE -ne 0) {
            throw "git add failed with exit code $LASTEXITCODE"
        }

        git diff --cached --quiet -- index.html data.json
        $diffExit = $LASTEXITCODE

        if ($diffExit -eq 1) {
            $message = "Update ICC dashboard data " + (Get-Date -Format "yyyy-MM-dd HH:mm")
            git commit -m $message
            if ($LASTEXITCODE -ne 0) {
                throw "git commit failed with exit code $LASTEXITCODE"
            }

            git push origin main
            if ($LASTEXITCODE -ne 0) {
                throw "git push failed with exit code $LASTEXITCODE"
            }
        } elseif ($diffExit -eq 0) {
            Write-Host "No index.html/data.json changes to deploy."
        } else {
            throw "git diff failed with exit code $diffExit"
        }
    }

    Write-Host ("ICC daily update completed at {0}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"))
} catch {
    Write-Error $_
    exit 1
} finally {
    if ($logPath) {
        Stop-Transcript | Out-Null
        Write-Host "Log written to $logPath"
    }
}
