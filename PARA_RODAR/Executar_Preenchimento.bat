@echo off
cd /d "%~dp0"
if not exist "PreencherFotosExcel.exe" (
  echo Executavel nao encontrado nesta pasta.
  pause
  exit /b 1
)
"%~dp0PreencherFotosExcel.exe"
echo.
echo Processamento finalizado. Pressione uma tecla para fechar.
pause >nul
