@echo off
rem --- paths -------------------------------------------------
setlocal
set SCRIPT=%~dp0html2pdf.ps1
set HTML=%~dp0smoke.html
rem -----------------------------------------------------------

powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -InputHtml "%HTML%"