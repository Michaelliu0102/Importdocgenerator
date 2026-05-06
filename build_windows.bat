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

set "WARN_FILE=build\ClearanceOS\warn-ClearanceOS.txt"
if exist "%WARN_FILE%" (
  findstr /c:"missing module named main" "%WARN_FILE%" >nul
  if not errorlevel 1 (
    echo Error: PyInstaller did not collect main.py
    exit /b 1
  )
)

echo.
echo Build complete: dist\ClearanceOS\ClearanceOS.exe
if not defined CI pause
