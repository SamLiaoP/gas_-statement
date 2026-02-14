@echo off
cd /d "%~dp0"
git pull origin main
python -m pip install -q -r requirements.txt
python 電子支付對帳程式\main.py || pause
