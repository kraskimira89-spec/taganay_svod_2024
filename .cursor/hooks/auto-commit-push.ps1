$ErrorActionPreference = "Stop"

$repo = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path
$logPath = Join-Path $PSScriptRoot "auto-commit-push.log"

function Write-HookLog {
    param([string]$Message)
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -LiteralPath $logPath -Encoding UTF8 -Value "[$ts] $Message"
}

try {
    Set-Location $repo

    $inside = git rev-parse --is-inside-work-tree 2>$null
    if ($LASTEXITCODE -ne 0 -or $inside -ne "true") {
        Write-HookLog "skip: not a git repository"
        Write-Output "{}"
        exit 0
    }

    git add -A
    $pending = git status --porcelain
    if (-not $pending) {
        Write-HookLog "skip: no changes"
        Write-Output "{}"
        exit 0
    }

    $subject = "Автокоммит после выполнения задания"
    $body = "Создано автоматически hook-ом Cursor."
    git commit -m $subject -m $body
    if ($LASTEXITCODE -ne 0) {
        Write-HookLog "commit failed"
        Write-Output "{}"
        exit 0
    }

    git push origin HEAD
    if ($LASTEXITCODE -eq 0) {
        Write-HookLog "commit and push succeeded"
    }
    else {
        Write-HookLog "push failed; commit remains local"
    }

    Write-Output "{}"
    exit 0
}
catch {
    Write-HookLog ("error: " + $_.Exception.Message)
    Write-Output "{}"
    exit 0
}
