@echo off
REM ================================================================
REM Сборка DocCheck.exe — единого исполняемого файла.
REM
REM Требования к машине, на которой собирается .exe:
REM   - Windows 10/11
REM   - Python 3.10+ (https://python.org, при установке поставьте
REM     галочку "Add Python to PATH").
REM
REM На рабочей машине, где будет запускаться DocCheck.exe,
REM ставить Python НЕ нужно — всё запаковано внутри.
REM ================================================================

setlocal
cd /d "%~dp0"

echo === DocCheck build ===
echo.

where python >nul 2>&1
if errorlevel 1 (
    echo [!] Python не найден в PATH.
    echo     Установите Python 3.10+ с https://python.org
    echo     и не забудьте галочку "Add Python to PATH".
    pause
    exit /b 1
)

echo Устанавливаю зависимости...
python -m pip install --upgrade pip
python -m pip install python-docx tkinterdnd2 pyinstaller
if errorlevel 1 (
    echo [!] Не удалось установить зависимости.
    pause
    exit /b 1
)

echo.
echo Собираю DocCheck.exe...

if exist build rmdir /s /q build
if exist dist  rmdir /s /q dist

python -m PyInstaller --onefile --windowed ^
    --name DocCheck ^
    --add-data "..\python\check_word_doc.py;." ^
    --add-data "..\blank_template.txt;." ^
    --collect-all tkinterdnd2 ^
    app.py

if errorlevel 1 (
    echo [!] Сборка не удалась.
    pause
    exit /b 1
)

REM Положим blank_template.txt рядом с готовым .exe — чтобы пользователь
REM мог его редактировать (внутри .exe он только как fallback).
copy /Y "..\blank_template.txt" "dist\blank_template.txt" >nul

echo.
echo ================================================================
echo Готово!
echo.
echo   dist\DocCheck.exe          - исполняемый файл
echo   dist\blank_template.txt    - образец бланка (редактируется)
echo.
echo Скопируйте обе вещи в одну папку на рабочей машине и запускайте.
echo ================================================================
pause
