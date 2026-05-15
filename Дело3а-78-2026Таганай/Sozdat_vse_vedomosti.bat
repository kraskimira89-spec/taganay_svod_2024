@echo off
chcp 65001 >nul
REM Ведомости разделов 01–11. Скрипт: репозиторий scripts\sozdat_vedomost_razdela.py
REM При необходимости укажите базовую ссылку Я.Диска:
set "YANDEX_BASE="

set "REPO=%~dp0.."
set "SCRIPT=%REPO%\scripts\sozdat_vedomost_razdela.py"
cd /d "%~dp0"

python --version >nul 2>&1
if errorlevel 1 (
    echo ОШИБКА: Python не в PATH
    pause
    exit /b 1
)
if not exist "%SCRIPT%" (
    echo ОШИБКА: не найден %SCRIPT%
    pause
    exit /b 1
)

echo ================================================================
echo  Ведомости 01–11 (имена папок как в этом каталоге)
echo ================================================================
echo.

set SECTIONS=01_Иск 02_Отчёты 03_Учредительные 04_Соглашения 05_ЕКЖЯ 06_Финансы 07_Кадры 08_Аренда 09_Нормативка 10_Зарплаты 11_Поставщики

for %%S in (%SECTIONS%) do (
    if exist "%%S\" (
        echo --- %%S
        if defined YANDEX_BASE (
            python "%SCRIPT%" "%%S" --yandex-base "%YANDEX_BASE%"
        ) else (
            python "%SCRIPT%" "%%S"
        )
        echo.
    ) else (
        echo [SKIP] нет папки: %%S
    )
)

echo ================================================================
echo  Готово. Файлы: 00_Ведомость_XX_*.xlsx в каждой папке раздела
echo ================================================================
