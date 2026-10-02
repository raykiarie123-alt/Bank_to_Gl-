@echo off
cd /d "%~dp0"
echo Installing / checking required parts (first time takes a few minutes)...
python -m pip install -r requirements.txt
echo.
echo Starting the app. Your browser will open. Keep this black window open while you work.
python -m streamlit run app.py
pause
