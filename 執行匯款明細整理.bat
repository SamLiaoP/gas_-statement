@echo off
cd /d "%~dp0"
git pull origin main
python -m pip install -q -r requirements.txt
python LinePay匯款明細整理\main.py || pause
