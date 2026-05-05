# Push this folder to remote repo "Undergraduate_Graduation_Project"
# Prereq: Git for Windows, empty GitHub repo (no README)

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$machinePath = [System.Environment]::GetEnvironmentVariable("Path", "Machine")
$userPath = [System.Environment]::GetEnvironmentVariable("Path", "User")
if ($machinePath -or $userPath) {
    $env:Path = @($machinePath, $userPath, $env:Path) -join ";"
}

function Find-GitExe {
    $pf86 = [Environment]::GetEnvironmentVariable("ProgramFiles(x86)")
    $extra = @()
    if ($pf86) { $extra += (Join-Path $pf86 "Git\cmd\git.exe") }
    $candidates = @(
        (Get-Command git -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source)
        "${env:ProgramFiles}\Git\cmd\git.exe"
        "${env:ProgramFiles}\Git\bin\git.exe"
    ) + $extra + @(
        "${env:LocalAppData}\Programs\Git\cmd\git.exe"
    ) | Where-Object { $_ -and (Test-Path -LiteralPath $_) }
    return $candidates | Select-Object -First 1
}

$script:GitExe = Find-GitExe
if (-not $script:GitExe) {
    Write-Host "git.exe not found. Install Git for Windows, restart terminal, or check:"
    Write-Host "  C:\Program Files\Git\cmd\git.exe"
    exit 1
}

function Invoke-Git {
    param([Parameter(Mandatory = $true)][string[]]$GitArgs)
    & $script:GitExe @GitArgs
    if ($LASTEXITCODE -ne 0) {
        throw "git failed (exit $LASTEXITCODE): git $($GitArgs -join ' ')"
    }
}

Write-Host "Using Git: $script:GitExe"
Write-Host ""

if (-not (Test-Path .git)) {
    Invoke-Git @("init")
    Invoke-Git @("branch", "-M", "main")
}

Invoke-Git @("add", "-A")
$dirty = & $script:GitExe status --porcelain
if ($dirty) {
    Invoke-Git @("commit", "-m", "chore: BIT thesis (bithesis + Ch.2 Franka/LeRobot)")
} else {
    Write-Host "Nothing to commit, skipping."
}

Write-Host ""
Write-Host "Paste remote URL (HTTPS or SSH), e.g."
Write-Host '  https://github.com/YOUR_USER/Undergraduate_Graduation_Project.git'
$remoteUrl = Read-Host "Remote URL"
if ([string]::IsNullOrWhiteSpace($remoteUrl)) {
    Write-Host "No URL. Run later:"
    Write-Host "  git remote add origin <url>"
    Write-Host "  git push -u origin main"
    exit 0
}

# Do not use "git remote get-url origin" here: git writes to stderr when missing, and with
# $ErrorActionPreference=Stop PowerShell treats that as terminating. Listing remotes is benign.
$remoteNames = @(& $script:GitExe @("remote"))
if ($LASTEXITCODE -ne 0) {
    throw "git remote failed (exit $LASTEXITCODE)"
}
$hasOrigin = $remoteNames -contains "origin"
if ($hasOrigin) {
    Invoke-Git @("remote", "set-url", "origin", $remoteUrl)
} else {
    Invoke-Git @("remote", "add", "origin", $remoteUrl)
}

Invoke-Git @("push", "-u", "origin", "main")
Write-Host "Done."
