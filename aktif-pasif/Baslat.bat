@echo off
setlocal
cd /d "%~dp0"
if exist "%~dp0Domain-Aktif-Pasif-Kontrol.exe" (
  start "" "%~dp0Domain-Aktif-Pasif-Kontrol.exe"
  goto :eof
)
where py >nul 2>nul
if not errorlevel 1 (
  py -3 domain_checker.py
  goto :eof
)
where python >nul 2>nul
if not errorlevel 1 (
  python domain_checker.py
  goto :eof
)
echo Python 3 bulunamadi. https://www.python.org/downloads/ adresinden Python 3 kurup tekrar deneyin.
echo Kurulum sirasinda "Add Python to PATH" secenegini isaretleyin.
pause
