@echo off
cd /d "%~dp0"
git pull origin main
python -m pip install -q -r requirements.txt
python linepay_summary\main.py || pause
