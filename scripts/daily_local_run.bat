@echo off
rem 每天由 Windows 任务计划程序调用：本机运行 gd_psych，并把 web/data/gd-psych.json 推送到 GitHub。
rem 日志：logs\gd_psych-日期.log
rem 手动试运行：双击本文件，或在命令行加参数，如  daily_local_run.bat --dry-run
chcp 65001 >nul
cd /d "%~dp0.."
set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1
if not exist ".venv\Scripts\python.exe" (
    echo 找不到 .venv，请先在项目目录运行：python -m venv .venv ^&^& .venv\Scripts\python -m pip install -r requirements.txt
    exit /b 1
)
".venv\Scripts\python.exe" scripts\local_gd_psych.py %*
exit /b %ERRORLEVEL%
