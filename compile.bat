@echo off
setlocal
cd /d "%~dp0"

echo [compile.bat] latexmk clean (main.tex aux etc.)...
latexmk -C main.tex
if errorlevel 1 (
  echo [compile.bat] clean finished with warnings or errors - still attempting build.
)

echo [compile.bat] latexmk build (nonstop mode^)...
latexmk -interaction=nonstopmode main.tex
exit /b %ERRORLEVEL%
