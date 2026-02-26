@echo off
setlocal

REM Gera executavel do script "Inserir foto.py" usando PyInstaller.
REM Requisitos:
REM   python -m pip install pyinstaller pywin32

cd /d "%~dp0"

if not exist "Inserir foto.py" (
  echo Arquivo "Inserir foto.py" nao encontrado nesta pasta.
  exit /b 1
)

echo Limpando builds anteriores...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist "__pycache__" rmdir /s /q __pycache__

set "PY_CMD=python"
where py >nul 2>nul
if not errorlevel 1 set "PY_CMD=py"

echo Gerando executavel...
%PY_CMD% -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --onefile ^
  --console ^
  --name PreencherFotosExcel ^
  --hidden-import win32com ^
  --hidden-import win32com.client ^
  --hidden-import pythoncom ^
  --hidden-import pywintypes ^
  --hidden-import win32timezone ^
  "Inserir foto.py"

if errorlevel 1 (
  echo.
  echo Falha ao gerar o executavel.
  exit /b 1
)

echo.
echo Executavel gerado em:
echo   dist\PreencherFotosExcel.exe
echo.
echo Uso recomendado:
echo   1) Coloque as planilhas na pasta "PLANILHAS" (opcional, ao lado do .exe)
echo   2) Coloque as imagens na pasta "FOTOS" (opcional, ao lado do .exe)
echo   3) Execute o .exe

endlocal
