@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

set "FILE=Список_використаних_джерел.docx"

set "P1=%APPDATA%\Claude\local-agent-mode-sessions\f5207617-d742-43be-b44b-0e749532c157\efdda5bd-cfbf-4703-81d1-4a54ffa1763d\local_f8cde63e-6bbc-4a3f-9669-de2bed4a39e1\outputs\!FILE!"
set "P2=%LOCALAPPDATA%\Packages\Claude_pzs8sxrjxfjjc\LocalCache\Roaming\Claude\local-agent-mode-sessions\f5207617-d742-43be-b44b-0e749532c157\efdda5bd-cfbf-4703-81d1-4a54ffa1763d\local_f8cde63e-6bbc-4a3f-9669-de2bed4a39e1\outputs\!FILE!"

set "DST=%~dp0!FILE!"

echo Шукаю !FILE!...
echo.

if exist "!P1!" (
    echo Знайдено: !P1!
    copy /Y "!P1!" "!DST!" >nul
    goto done
)

if exist "!P2!" (
    echo Знайдено: !P2!
    copy /Y "!P2!" "!DST!" >nul
    goto done
)

echo Шукаю в інших місцях AppData...
for /f "delims=" %%i in ('dir /s /b "%APPDATA%\!FILE!" 2^>nul') do (
    echo Знайдено: %%i
    copy /Y "%%i" "!DST!" >nul
    goto done
)
for /f "delims=" %%i in ('dir /s /b "%LOCALAPPDATA%\!FILE!" 2^>nul') do (
    echo Знайдено: %%i
    copy /Y "%%i" "!DST!" >nul
    goto done
)

echo.
echo Файл НЕ знайдено автоматично.
echo Перегляньте паки Cowork вручну і знайдіть: !FILE!
pause
exit /b 1

:done
echo.
echo ГОТОВО!
echo Файл скопійовано у:
echo !DST!
echo.
echo Можете відкривати у Word.
pause
