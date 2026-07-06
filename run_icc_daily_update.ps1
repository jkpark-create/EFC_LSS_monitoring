param(
    [string]$Url = $env:ICC_URL,
    [switch]$Headless,
    [switch]$XPlatform,
    [string]$DownloadFile,
    [switch]$Deploy,
    [switch]$NoLog,
    [int]$XPlatformAttempts = 3,
    [int]$Weeks = 0,
    [string]$Date = $env:ICC_RUN_DATE,
    [int]$TargetYear = 0,
    [int]$TargetWeek = 0,
    [int]$StartYear = 0,
    [int]$StartWeek = 0,
    [switch]$SkipPagesDeployCheck,
    [int]$PagesDeployRetries = 1,
    [int]$PagesDeployTimeoutSeconds = 900,
    [int]$PagesDeployPollSeconds = 20
)

$ErrorActionPreference = "Stop"
Set-Location -Path $PSScriptRoot

function Get-PagesWorkflowRun {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CommitSha,
        [long[]]$ExcludeRunIds = @()
    )

    $runsJson = gh run list --limit 30 --json databaseId,headSha,status,conclusion,url,createdAt,workflowName,name 2>$null
    if ($LASTEXITCODE -ne 0) {
        throw "gh run list failed with exit code $LASTEXITCODE"
    }

    $runs = @($runsJson | ConvertFrom-Json)
    $matchingRuns = @($runs |
        Where-Object {
            $_.headSha -eq $CommitSha `
                -and $_.workflowName -eq "pages-build-deployment" `
                -and ($ExcludeRunIds -notcontains [long]($_.databaseId))
        } |
        Sort-Object createdAt -Descending)

    $activeRun = $matchingRuns |
        Where-Object { $_.status -ne "completed" } |
        Select-Object -First 1
    if ($activeRun) {
        return $activeRun
    }

    $successfulRun = $matchingRuns |
        Where-Object { $_.conclusion -eq "success" } |
        Select-Object -First 1
    if ($successfulRun) {
        return $successfulRun
    }

    return $matchingRuns | Select-Object -First 1
}

function Wait-PagesWorkflowRun {
    param(
        [Parameter(Mandatory = $true)]
        [long]$RunId,
        [Parameter(Mandatory = $true)]
        [int]$TimeoutSeconds,
        [Parameter(Mandatory = $true)]
        [int]$PollSeconds
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)

    while ((Get-Date) -lt $deadline) {
        $runJson = gh run view $RunId --json status,conclusion,url 2>$null
        if ($LASTEXITCODE -ne 0) {
            throw "gh run view $RunId failed with exit code $LASTEXITCODE"
        }

        $run = $runJson | ConvertFrom-Json
        $conclusion = if ($run.conclusion) { $run.conclusion } else { "pending" }
        Write-Host ("GitHub Pages run {0}: {1}/{2}" -f $RunId, $run.status, $conclusion)

        if ($run.status -eq "completed") {
            return $run
        }

        Start-Sleep -Seconds $PollSeconds
    }

    throw "GitHub Pages run $RunId did not complete within $TimeoutSeconds seconds"
}

function Get-GitHubRepoName {
    $repoName = gh repo view --json nameWithOwner --jq .nameWithOwner 2>$null
    if ($LASTEXITCODE -ne 0 -or -not $repoName) {
        throw "gh repo view failed with exit code $LASTEXITCODE"
    }

    return $repoName.Trim()
}

function Request-GitHubPagesBuild {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoName
    )

    $buildJson = gh api --method POST "repos/$RepoName/pages/builds" 2>$null
    if ($LASTEXITCODE -ne 0) {
        throw "gh api pages/builds failed with exit code $LASTEXITCODE"
    }

    $build = $buildJson | ConvertFrom-Json
    Write-Host ("Requested GitHub Pages rebuild: {0}" -f $build.status)
}

function Restart-GitHubPagesDeployment {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoName,
        [Parameter(Mandatory = $true)]
        [long]$RunId
    )

    $pagesJson = gh api "repos/$RepoName/pages" 2>$null
    if ($LASTEXITCODE -ne 0) {
        throw "gh api pages failed with exit code $LASTEXITCODE"
    }

    $pages = $pagesJson | ConvertFrom-Json
    if ($pages.build_type -eq "legacy") {
        Request-GitHubPagesBuild -RepoName $RepoName
        return $true
    }

    gh run rerun $RunId --failed | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "gh run rerun $RunId failed with exit code $LASTEXITCODE"
    }

    return $false
}

function Wait-GitHubPagesDeployment {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CommitSha,
        [Parameter(Mandatory = $true)]
        [int]$Retries,
        [Parameter(Mandatory = $true)]
        [int]$TimeoutSeconds,
        [Parameter(Mandatory = $true)]
        [int]$PollSeconds
    )

    if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
        Write-Warning "GitHub CLI (gh) is not available; skipping GitHub Pages deployment check."
        return
    }

    $repoName = Get-GitHubRepoName
    $attempt = 0
    $excludeRunIds = @()
    while ($attempt -le $Retries) {
        $run = $null
        $discoveryDeadline = (Get-Date).AddSeconds([Math]::Min(180, $TimeoutSeconds))

        while (-not $run -and (Get-Date) -lt $discoveryDeadline) {
            $run = Get-PagesWorkflowRun -CommitSha $CommitSha -ExcludeRunIds $excludeRunIds
            if (-not $run) {
                Write-Host "Waiting for GitHub Pages workflow run to appear."
                Start-Sleep -Seconds $PollSeconds
            }
        }

        if (-not $run) {
            throw "GitHub Pages workflow run was not found for commit $CommitSha"
        }

        Write-Host ("Watching GitHub Pages workflow run: {0}" -f $run.url)
        try {
            $result = Wait-PagesWorkflowRun -RunId $run.databaseId -TimeoutSeconds $TimeoutSeconds -PollSeconds $PollSeconds
        } catch {
            if ($attempt -ge $Retries) {
                throw
            }

            Write-Warning ("GitHub Pages run did not complete: {0}. Requesting retry." -f $_.Exception.Message)
            $startedNewRun = Restart-GitHubPagesDeployment -RepoName $repoName -RunId $run.databaseId
            if ($startedNewRun) {
                $excludeRunIds += [long]$run.databaseId
            }
            $attempt += 1
            Start-Sleep -Seconds $PollSeconds
            continue
        }

        if ($result.conclusion -eq "success") {
            Write-Host "GitHub Pages deployment succeeded."
            return
        }

        if ($attempt -ge $Retries) {
            throw ("GitHub Pages deployment failed after {0} attempt(s): {1}" -f ($attempt + 1), $result.url)
        }

        Write-Warning ("GitHub Pages deployment failed with conclusion '{0}'. Requesting retry." -f $result.conclusion)
        $startedNewRun = Restart-GitHubPagesDeployment -RepoName $repoName -RunId $run.databaseId
        if ($startedNewRun) {
            $excludeRunIds += [long]$run.databaseId
        }
        $attempt += 1
        Start-Sleep -Seconds $PollSeconds
    }
}

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
            if ($Weeks -gt 0) {
                $xplatformArguments += @("--weeks", $Weeks)
            }
            if ($Date) {
                $xplatformArguments += @("--date", $Date)
            }
            if ($TargetYear -gt 0) {
                $xplatformArguments += @("--target-year", $TargetYear)
            }
            if ($TargetWeek -gt 0) {
                $xplatformArguments += @("--target-week", $TargetWeek)
            }
            if ($StartYear -gt 0) {
                $xplatformArguments += @("--start-year", $StartYear)
            }
            if ($StartWeek -gt 0) {
                $xplatformArguments += @("--start-week", $StartWeek)
            }

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
    if ($Weeks -gt 0) {
        $arguments += @("--weeks", $Weeks)
    }
    if ($Date) {
        $arguments += @("--date", $Date)
    }
    if ($TargetYear -gt 0) {
        $arguments += @("--target-year", $TargetYear)
    }
    if ($TargetWeek -gt 0) {
        $arguments += @("--target-week", $TargetWeek)
    }
    if ($StartYear -gt 0) {
        $arguments += @("--start-year", $StartYear)
    }
    if ($StartWeek -gt 0) {
        $arguments += @("--start-week", $StartWeek)
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

            if (-not $SkipPagesDeployCheck) {
                $commitSha = (git rev-parse HEAD).Trim()
                if ($LASTEXITCODE -ne 0) {
                    throw "git rev-parse failed with exit code $LASTEXITCODE"
                }

                Wait-GitHubPagesDeployment `
                    -CommitSha $commitSha `
                    -Retries $PagesDeployRetries `
                    -TimeoutSeconds $PagesDeployTimeoutSeconds `
                    -PollSeconds $PagesDeployPollSeconds
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
