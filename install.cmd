@echo off

setlocal

:: Путь к вашему пакетному файлу (замените на имя вашего файла)
set target_file=%~dp0sign.cmd

:: Путь к ярлыку на рабочем столе, с пробелами в имени
set shortcut_name=ЭП Подпись
set shortcut_path=%USERPROFILE%\Desktop\%shortcut_name%.lnk

:: Путь к стандартной иконке (например, из shell32.dll)
:: Пример иконки документа
set icon_location=shell32.dll,269

:: Проверяем, существует ли ярлык, и удаляем его, если он есть
if exist "%shortcut_path%" (
    del "%shortcut_path%"
)

:: Создаём ярлык с помощью PowerShell
powershell -command "$s = (New-Object -COM WScript.Shell).CreateShortcut('%shortcut_path%'); $s.TargetPath = '%target_file%'; $s.IconLocation = '%icon_location%'; $s.Save()"

echo Ярлык был создан на рабочем столе.
endlocal
exit /b
