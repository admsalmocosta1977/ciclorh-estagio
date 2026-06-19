@echo off
cd /d "%~dp0"
git add -A
git commit -m "Atualizacao do sistema"
git push origin main
echo.
echo Pronto! Arquivos enviados ao GitHub.
pause
