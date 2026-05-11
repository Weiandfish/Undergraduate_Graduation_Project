@echo off
setlocal EnableDelayedExpansion
cd /d "%~dp0"

REM Reload Machine + User PATH from the registry. IDE terminals (e.g. Cursor) often inherit an
REM outdated PATH from when the editor started; TeX works in a "fresh" PowerShell but not here until restart.
for /f "usebackq delims=" %%P in (`powershell -NoProfile -ExecutionPolicy Bypass -Command "[Environment]::GetEnvironmentVariable('Path','Machine') + ';' + [Environment]::GetEnvironmentVariable('Path','User')"`) do set "PATH_REGISTRY=%%P"
if defined PATH_REGISTRY set "PATH=%PATH_REGISTRY%;%PATH%"

REM Add MiKTeX / TeX Live bin dirs to PATH for this session when tools are not on system PATH.
call :prepend_if_texbin "%ProgramFiles%\MiKTeX\miktex\bin\x64"
call :prepend_if_texbin "%ProgramFiles%\MiKTeX\miktex\bin"
call :prepend_if_texbin "%LocalAppData%\Programs\MiKTeX\miktex\bin\x64"
call :prepend_if_texbin "%LocalAppData%\Programs\MiKTeX\miktex\bin"
for /d %%Y in ("C:\texlive\*") do call :prepend_if_texbin "%%~fY\bin\win32"
for /d %%Y in ("%USERPROFILE%\texlive\*") do call :prepend_if_texbin "%%~fY\bin\win32"
for /d %%Y in ("D:\texlive\*") do call :prepend_if_texbin "%%~fY\bin\win32"
for /d %%Y in ("E:\texlive\*") do call :prepend_if_texbin "%%~fY\bin\win32"
for /d %%Y in ("F:\texlive\*") do call :prepend_if_texbin "%%~fY\bin\win32"

where latexmk >nul 2>&1
if %ERRORLEVEL% equ 0 goto :with_latexmk

where xelatex >nul 2>&1
if %ERRORLEVEL% equ 0 goto :manual_xelatex

echo.
echo [compile.bat] ERROR: Neither latexmk nor xelatex was found.
echo.
echo Install a TeX distribution, then either:
echo   1^) Add its "bin" folder to your user or system PATH, or
echo   2^) Install MiKTeX to the default location, or TeX Live under C:\texlive\
echo      ^(this script prepends those locations automatically if they exist^).
echo.
echo   MiKTeX:  https://miktex.org/download
echo   TeX Live: https://tug.org/texlive/windows.html
echo.
echo If TeX works in a new PowerShell window but not here, fully quit and restart Cursor,
echo or run: powershell -ExecutionPolicy Bypass -File .\fix-texlive-path.ps1
echo.
exit /b 127

:prepend_if_texbin
if exist "%~1\xelatex.exe" set "PATH=%~1;%PATH%"
goto :eof

:with_latexmk
echo [compile.bat] latexmk clean ^(main.tex aux etc.^)...
latexmk -C main.tex
if errorlevel 1 (
  echo [compile.bat] clean finished with warnings or errors - still attempting build.
)
echo [compile.bat] latexmk build ^(nonstop mode^)...
latexmk -interaction=nonstopmode -xelatex main.tex
exit /b %ERRORLEVEL%

:manual_xelatex
echo [compile.bat] latexmk not found; building with xelatex + biber ^(BIThesis workflow^)...

set XETEX=xelatex -interaction=nonstopmode -file-line-error

%XETEX% main.tex
if errorlevel 1 exit /b 1

where biber >nul 2>&1
if %ERRORLEVEL% equ 0 (
  biber main
  if errorlevel 1 exit /b 1
) else (
  echo [compile.bat] WARNING: biber not found; skipping bibliography pass.
)

%XETEX% main.tex
if errorlevel 1 exit /b 1
%XETEX% main.tex
exit /b %ERRORLEVEL%
