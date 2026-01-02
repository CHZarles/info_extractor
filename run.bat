@echo off
setlocal
pushd %~dp0

if not exist .venv\Scripts\activate.bat (
  echo [ERROR] 未找到虛擬環境 .venv，請先在此目錄運行: python -m venv .venv && .venv\Scripts\activate && pip install -r requirements.txt
  pause
  exit /b 1
)

call .venv\Scripts\activate
python app.py

popd
pause
endlocal
