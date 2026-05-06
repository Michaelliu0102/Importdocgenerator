@echo off
setlocal

cd /d "%~dp0"

if not exist ".venv" (
  echo Creating virtual environment .venv ...
  py -3 -m venv .venv
)

call ".venv\Scripts\activate.bat"

python -m pip uninstall -y tkinterdnd2
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install --upgrade pyinstaller

pyinstaller ^
  --noconfirm ^
  --windowed ^
  --name "ClearanceOS" ^
  --paths "%cd%" ^
  --hidden-import main ^
  --hidden-import eur_invoice_standalone ^
  --add-data "templates;templates" ^
  --add-data "export_templates;export_templates" ^
  --add-data "data;data" ^
  gui_app.py

set "TOC_FILE=build\ClearanceOS\PYZ-00.toc"
if not exist "%TOC_FILE%" (
  echo Error: Missing %TOC_FILE%
  exit /b 1
)

findstr /c:"('main'," "%TOC_FILE%" >nul
if not %errorlevel%==0 (
  echo Error: PyInstaller output does not contain main
  exit /b 1
)

findstr /c:"('eur_invoice_standalone'," "%TOC_FILE%" >nul
if not %errorlevel%==0 (
  echo Error: PyInstaller output does not contain eur_invoice_standalone
  exit /b 1
)

if exist "build\ClearanceOS\warn-ClearanceOS.txt" (
  findstr /c:"missing module named main" "build\ClearanceOS\warn-ClearanceOS.txt" >nul
  if %errorlevel%==2 (
    echo Error: Unable to inspect build\ClearanceOS\warn-ClearanceOS.txt
    exit /b 1
  )
)

echo.
echo Build complete: dist\ClearanceOS\ClearanceOS.exe
if not defined CI pause
exit /b 0
