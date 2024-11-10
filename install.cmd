@echo off

setlocal

:: ���� � ��襬� ����⭮�� 䠩�� (������� �� ��� ��襣� 䠩��)
set target_file=%~dp0sign.cmd

:: ���� � ���� �� ࠡ�祬 �⮫�, � �஡����� � �����
set shortcut_name=�� �������
set shortcut_path=%USERPROFILE%\Desktop\%shortcut_name%.lnk

:: ���� � �⠭���⭮� ������ (���ਬ��, �� shell32.dll)
:: �ਬ�� ������ ���㬥��
set icon_location=shell32.dll,269

:: �஢��塞, ������� �� ���, � 㤠�塞 ���, �᫨ �� ����
if exist "%shortcut_path%" (
    del "%shortcut_path%"
)

:: ������ ��� � ������� PowerShell
powershell -command "$s = (New-Object -COM WScript.Shell).CreateShortcut('%shortcut_path%'); $s.TargetPath = '%target_file%'; $s.IconLocation = '%icon_location%'; $s.Save()"

echo ��� �� ᮧ��� �� ࠡ�祬 �⮫�.
endlocal
exit /b
