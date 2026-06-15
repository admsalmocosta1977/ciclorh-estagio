@echo off
title CICLO RH — Sistema de Estagio
color 1F
echo.
echo  ============================================
echo   CICLO RH - Sistema de Gerenciamento de
echo   Estagio
echo  ============================================
echo.
echo  Iniciando o servidor...
echo  Abra o navegador em: http://localhost:5000
echo.
echo  Para encerrar, feche esta janela.
echo.

cd /d "%~dp0"
"C:\Users\usuario\AppData\Local\Programs\Python\Python312\python.exe" app.py

pause
