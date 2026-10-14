@echo off
REM ---------------------------------------------------------------------------
REM  Gera o executavel FormatadorDocx.exe (Windows).
REM  Basta fazer duplo clique neste ficheiro, com o Python 3 instalado.
REM  O resultado fica em dist\FormatadorDocx.exe
REM ---------------------------------------------------------------------------
cd /d "%~dp0\.."
echo.
echo === A instalar dependencias ===
python -m pip install --upgrade pip
python -m pip install -r requirements.txt pyinstaller
if errorlevel 1 goto erro

echo.
echo === A executar o teste de fumo ===
python scripts\teste_fumo.py
if errorlevel 1 goto erro

echo.
echo === A gerar o executavel ===
python -m PyInstaller build.spec --noconfirm --clean
if errorlevel 1 goto erro

echo.
echo Concluido. O executavel esta em: %cd%\dist\FormatadorDocx.exe
pause
exit /b 0

:erro
echo.
echo Ocorreu um erro. Verifique as mensagens acima.
pause
exit /b 1
