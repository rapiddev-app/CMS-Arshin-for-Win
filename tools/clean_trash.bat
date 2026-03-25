@echo off
chcp 65001 >nul
echo ===================================================
echo   ПОЛНАЯ ОЧИСТКА СИСТЕМЫ ОТ STAROGO SCRIPTMANAGER
echo ===================================================
echo.
echo ВНИМАНИЕ: Запустите этот файл от имени Администратора!
echo (Правая кнопка мыши -> Запуск от имени администратора)
echo.
pause

echo.
echo [1/5] Остановка процессов...
taskkill /F /IM python.exe /FI "MODULES eq python.exe" /FI "WINDOWTITLE eq ScriptManager*" >nul 2>&1
taskkill /F /IM pythonw.exe >nul 2>&1

echo.
echo [2/5] Удаление папки установки...
if exist "C:\Program Files\ScriptManager" (
    rd /s /q "C:\Program Files\ScriptManager"
    echo    Папка удалена.
) else (
    echo    Папка не найдена или уже удалена.
)

echo.
echo [3/5] Удаление ярлыка с рабочего стола...
if exist "%USERPROFILE%\Desktop\ScriptManager.lnk" (
    del /f /q "%USERPROFILE%\Desktop\ScriptManager.lnk"
    echo    Ярлык удален.
) else (
    echo    Ярлык не найден.
)

echo.
echo [4/5] Удаление из автозагрузки...
if exist "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\ScriptManager.lnk" (
    del /f /q "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\ScriptManager.lnk"
    echo    Автозагрузка очищена.
)

echo.
echo [5/5] Очистка реестра...
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ScriptManager" /f >nul 2>&1
echo    Запись в реестре удалена (если была).

echo.
echo ===================================================
echo   Очистка завершена!
echo   Теперь можно устанавливать новую версию.
echo ===================================================
pause
