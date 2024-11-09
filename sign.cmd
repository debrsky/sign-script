@echo off
chcp 1251 >nul
echo Starting signature process...
cscript //nologo sign.vbs %*
pause