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
  --name "报关资料生成器" ^
  --add-data "templates;templates" ^
  --add-data "export_templates;export_templates" ^
  --add-data "data;data" ^
  gui_app.py

echo.
echo Build complete: dist\报关资料生成器\报关资料生成器.exe
pause
