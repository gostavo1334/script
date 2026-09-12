@echo off
chcp 65001 >nul

:: Excel Range to Telegram Bot
:: Runs in background - captures only data area, not full screen

python "%~dp0simple_excel_to_telegram.py" %*
