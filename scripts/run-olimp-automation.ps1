param(
  [switch]$NoPull,
  [switch]$NoCommit,
  [switch]$NoPush,
  [switch]$NoPagesCheck,
  [string]$PagesUrl = "https://a1reklamaru.github.io/titantokos/olimpiyskaya-derevnya-dashboard.html"
)

$ErrorActionPreference = "Stop"

function Test-CanWriteGitDir() {
  $gitDir = Join-Path $PSScriptRoot "..\.git"
  if (-not (Test-Path -LiteralPath $gitDir)) { return $false }
  $probe = Join-Path $gitDir "__codex_write_probe.tmp"
  try {
    New-Item -ItemType File -LiteralPath $probe -Force -ErrorAction Stop | Out-Null
    Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
    return $true
  }
  catch {
    return $false
  }
}

function Try-Run([scriptblock]$Action, [string]$Label) {
  try {
    & $Action
    return $true
  }
  catch {
    Write-Warning "${Label}: $($_.Exception.Message)"
    return $false
  }
}

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
Push-Location $repoRoot
try {
  $canWriteGit = Test-CanWriteGitDir

  if (-not $NoPull) {
    if ($canWriteGit) {
      Try-Run { git pull --ff-only } "git pull --ff-only" | Out-Null
    }
    else {
      Write-Warning "Skipping git pull --ff-only: no write access to .git"
    }
  }

  & powershell -ExecutionPolicy Bypass -File .\scripts\build-olimp-dashboard.ps1

  if (-not $canWriteGit) {
    Write-Warning "No write access to .git: skipping git add/commit/push (dashboard built locally)."
  }
  else {
    $status = git status --porcelain -- .\olimpiyskaya-derevnya-dashboard.html .\scripts\build-olimp-dashboard.ps1
    if (-not $status) {
      Write-Host "No changes in tracked files."
    }
    else {
      git add .\olimpiyskaya-derevnya-dashboard.html .\scripts\build-olimp-dashboard.ps1

      if (-not $NoCommit) {
        $hasStaged = git diff --cached --name-only
        if ($hasStaged) {
          git -c user.name="codex-automation" -c user.email="codex-automation@users.noreply.github.com" commit -m "Update olimp dashboard"
        }
      }

      if (-not $NoPush) {
        Try-Run { git push origin main } "git push origin main" | Out-Null
      }
    }
  }

  if (-not $NoPagesCheck) {
    $cb = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    $url = "$PagesUrl?cb=$cb"
      Try-Run {
        $r = Invoke-WebRequest -Uri $url -UseBasicParsing -Headers @{ "Cache-Control" = "no-cache"; "Pragma" = "no-cache" } -TimeoutSec 60
      Write-Host "GitHub Pages OK: $($r.StatusCode) ($url)"
      } "GitHub Pages check" | Out-Null
    }
  }
finally {
  Pop-Location
}

exit 0
