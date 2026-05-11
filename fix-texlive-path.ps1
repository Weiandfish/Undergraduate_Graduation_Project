# Finds TeX Live's bin\win32 (xelatex) and appends it to the *user* PATH if missing.
# Usage:
#   powershell -ExecutionPolicy Bypass -File .\fix-texlive-path.ps1
#   powershell -ExecutionPolicy Bypass -File .\fix-texlive-path.ps1 -TeXBin "D:\texlive\2025\bin\win32"

param(
    [string] $TeXBin = $null
)

$ErrorActionPreference = 'Stop'

function Add-UserPathEntry {
    param([string] $Dir)
    if (-not (Test-Path (Join-Path $Dir 'xelatex.exe'))) {
        throw "Not a TeX Live bin directory (no xelatex.exe): $Dir"
    }
    $norm = (Resolve-Path $Dir).Path.TrimEnd('\')
    $userPath = [Environment]::GetEnvironmentVariable('Path', 'User')
    $parts = @()
    if ($userPath) { $parts = $userPath -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ } }
    if ($parts -contains $norm) {
        Write-Host "[fix-texlive-path] User PATH already contains: $norm"
        return
    }
    $newPath = if ($userPath) { "$userPath;$norm" } else { $norm }
    [Environment]::SetEnvironmentVariable('Path', $newPath, 'User')
    Write-Host "[fix-texlive-path] Appended to user PATH: $norm"
}

if ($TeXBin) {
    Add-UserPathEntry -Dir $TeXBin
    Write-Host "[fix-texlive-path] Close and reopen this terminal (or restart Cursor), then run: xelatex --version"
    exit 0
}

$roots = @(
    "$env:USERPROFILE\texlive",
    'C:\texlive',
    'D:\texlive',
    'E:\texlive',
    'F:\texlive'
)

$found = New-Object System.Collections.Generic.List[string]
foreach ($root in $roots) {
    if (-not (Test-Path $root)) { continue }
    Get-ChildItem -Path $root -Directory -ErrorAction SilentlyContinue | ForEach-Object {
        $xe = Join-Path $_.FullName 'bin\win32\xelatex.exe'
        if (Test-Path $xe) {
            $found.Add((Split-Path $xe -Parent)) | Out-Null
        }
    }
}

if ($found.Count -eq 0) {
    Write-Host @"
[fix-texlive-path] Could not find xelatex under:
  $env:USERPROFILE\texlive\<year>\bin\win32
  C:\texlive\<year>\bin\win32
  (or D:/E:/F:\texlive\...)

TeX Live may be installed in a custom folder. Search for xelatex.exe in Explorer, then run:
  powershell -ExecutionPolicy Bypass -File .\fix-texlive-path.ps1 -TeXBin "FULL_PATH_TO_bin\win32"

During install, you can also enable "Add TeX Live to PATH" or add that folder manually in:
  Settings -> System -> About -> Advanced system settings -> Environment Variables -> Path (User)
"@
    exit 1
}

# Prefer lexicographically last path (usually newest year when roots are equal).
$chosen = ($found | Sort-Object -Descending)[0]
Write-Host "[fix-texlive-path] Found xelatex in: $chosen"
Add-UserPathEntry -Dir $chosen
Write-Host @"

[fix-texlive-path] Next step: CLOSE this terminal tab completely (or restart Cursor), open a new one, then:
  xelatex --version
  .\compile.bat
(Existing processes do not see updated PATH until restarted.)
"@
