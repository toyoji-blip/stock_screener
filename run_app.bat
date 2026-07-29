@echo off
cd /d "C:\Users\toyoj\Documents\stock-screening-webapp"

:: Pythonサーバーを別ウィンドウでバックグラウンド起動
start "StockAppServer" py app.py

:: サーバーが完全に立ち上がるまで2秒待機
timeout /t 2 /nobreak >nul

:: EdgeでWebアプリを開く
start "" msedge "http://127.0.0.1:5000"