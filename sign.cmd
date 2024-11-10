@echo off
chcp 1251 >nul
echo Starting signature process...
cscript //nologo "%~dp0sign.vbs" %*
pause