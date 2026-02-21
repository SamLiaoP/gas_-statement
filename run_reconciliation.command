#!/bin/bash
cd "$(dirname "$0")"
git pull origin main
python3 -m pip install -q -r requirements.txt
python3 reconciliation/main.py || read -p "發生錯誤，按 Enter 關閉視窗..."
