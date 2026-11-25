@echo off
REM Gera o FolhaDeCarga.exe no seu computador (precisa do Python instalado).
REM Normalmente nao precisa: o GitHub gera o .exe sozinho a cada mudanca
REM (aba Actions do repositorio). Use este arquivo so se quiser gerar na mao.

cd /d "%~dp0"
py -m pip install -r requirements.txt pyinstaller
py -m PyInstaller --onefile --windowed --name FolhaDeCarga --add-data "mapa_ordem.csv;." folha_carga_app.py

echo.
echo Pronto: o arquivo esta em dist\FolhaDeCarga.exe
pause
