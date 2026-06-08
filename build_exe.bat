@echo off
title Empacotando Leitor de Notas Ficais
echo ============================================
echo  Empacotando Leitor de Notas Ficais de Hardware
echo ============================================
echo.

pip install pyinstaller
if %ERRORLEVEL% neq 0 (
    echo Erro ao instalar PyInstaller
    pause
    exit /b 1
)

pyinstaller --onefile --windowed ^
    --name "Leitor de Notas Ficais" ^
    --add-data "leitor_de_pdf.py;." ^
    --hidden-import "PyPDF2" ^
    --hidden-import "openpyxl" ^
    --hidden-import "openpyxl.styles" ^
    --hidden-import "openpyxl.cell._writer" ^
    --hidden-import "rich" ^
    --hidden-import "rich.table" ^
    --hidden-import "rich.console" ^
    leitor_de_pdf_gui.py

if %ERRORLEVEL% neq 0 (
    echo.
    echo Erro ao empacotar
    pause
    exit /b 1
)

echo.
echo ============================================
echo  Executavel criado em: dist\Leitor de Notas Ficais.exe
echo ============================================
echo.
echo  Para executar, copie a pasta inteira para
echo  o Windows ou use o atalho na area de trabalho.
echo.
pause
