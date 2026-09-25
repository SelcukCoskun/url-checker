@echo off
setlocal
cd /d "%~dp0"
where py >nul 2>nul
if not errorlevel 1 goto use_py
where python >nul 2>nul
if not errorlevel 1 goto use_python
echo Python 3 bulunamadi. EXE olusturmak icin once Python 3 kurun.
pause
exit /b 1

:use_py
py -3 -m pip install --user pyinstaller -r requirements.txt
if errorlevel 1 goto build_failed
py -3 -m PyInstaller --noconfirm --clean --onefile --windowed --distpath . --specpath build --name Domain-Aktif-Pasif-Kontrol --collect-all httpx --collect-all httpcore domain_checker.py
goto build_result

:use_python
python -m pip install --user pyinstaller -r requirements.txt
if errorlevel 1 goto build_failed
python -m PyInstaller --noconfirm --clean --onefile --windowed --distpath . --specpath build --name Domain-Aktif-Pasif-Kontrol --collect-all httpx --collect-all httpcore domain_checker.py
goto build_result

:build_failed
echo Paketleme bagimliliklari kurulamadi. Internet baglantisini ve Python pip kurulumunu kontrol edin.
pause
exit /b 1

:build_result
if errorlevel 1 (
  echo EXE olusturulamadi. Python surumu veya PyInstaller uyumlulugunu kontrol edin.
) else (
  echo EXE hazir: Domain-Aktif-Pasif-Kontrol.exe
  echo Diger bilgisayara EXE dosyasini ve taranacak XLSX dosyalarini birlikte kopyalayin.
)
pause
