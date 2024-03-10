@echo off
REM Flaskアプリのプロセスをチェック
tasklist | findstr "python" > nul
if errorlevel 1 (
    echo Starting Flask app...
    REM VBScriptを使用してPythonアプリをバックグラウンドで実行
    start wscript.exe run_flask_app.vbs
    timeout /t 3
)

REM ブラウザでFlaskアプリを開く
start http://127.0.0.1:5000