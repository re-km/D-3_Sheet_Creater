@echo off
chcp 65001 > nul
cd /d "%~dp0"
echo ========================================================
echo D-3シート作成ツール
echo [注意] 先にExcelで対象のファイル(5桁_D-3.xlsm)を開いてください
echo ========================================================
python addin_clicker.py
echo.
echo 処理が終了しました。
pause
