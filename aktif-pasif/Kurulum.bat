@echo off
setlocal
cd /d "%~dp0"
where py >nul 2>nul
if not errorlevel 1 goto use_py
where python >nul 2>nul
if not errorlevel 1 goto use_python
echo Python 3 bulunamadi. https://www.python.org/downloads/ adresinden Python 3 kurup tekrar deneyin.
echo Kurulum sirasinda "Add Python to PATH" secenegini isaretleyin.
pause
exit /b 1

:use_py
py -3 -m pip install --user -r requirements.txt
goto show_result

:use_python
python -m pip install --user -r requirements.txt

:show_result
if errorlevel 1 (
  echo Bagimliliklar kurulamadi. Internet baglantisini ve Python pip kurulumunu kontrol edin.
  echo .XLS dosyalari xlrd olmadan okunamaz; HTTPX olmadan denetleyici yerlesik istemciyi kullanir.
) else (
  echo HTTPX ve XLS dosyasi okuyucusu kuruldu. run.py veya Baslat.bat dosyasini calistirabilirsiniz.
)
pause
