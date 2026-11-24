@echo off
setlocal

rem M-Files Toolkit - Windows Terminal'de iki bolmeli acilis.
rem Akis: guncelleyici (ayni pencerede, sessiz) -> Windows Terminal -> BOOT SEQUENCE -> ana menu.

set "ROOT=%~dp0"
set "ROOT=%ROOT:~0,-1%"

echo.
echo M-Files Toolkit baslatiliyor...

rem --------------------------------------------------
rem GUNCELLEME KONTROLU
rem --------------------------------------------------

if exist "%ROOT%\updater.ps1" (
    powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%ROOT%\updater.ps1"
)

rem Guncelleyici kendi dosyasini calisirken degistiremez; yeni surumunu .new olarak birakir, burada uygulanir.
if exist "%ROOT%\updater.ps1.new" move /y "%ROOT%\updater.ps1.new" "%ROOT%\updater.ps1" >nul

rem baslat.bat kendisi calisirken degistirilemez. Parantezli blok cmd tarafindan bastan okundugu icin
rem blok icinde dosyayi degistirip yeni surumu cagirmak ve cikmak guvenlidir.
if exist "%ROOT%\baslat.bat.new" (
    move /y "%ROOT%\baslat.bat.new" "%~f0" >nul
    call "%~f0"
    exit /b
)

rem --------------------------------------------------
rem WINDOWS TERMINAL KONTROLU
rem --------------------------------------------------

where wt.exe >nul 2>nul
if errorlevel 1 (
    echo Windows Terminal ^(wt.exe^) bulunamadi.
    echo Microsoft Store'dan Windows Terminal kurun.
    pause
    exit /b 1
)

rem --------------------------------------------------
rem PANEL ORANI
rem --------------------------------------------------

rem Sag panelin orani Data\Common\Config\terminal.json icindeki MenuWidth'ten hesaplanir
rem (guncelleyici hesaplayip split.txt'ye yazar). Dosya yoksa ya da gecersizse menu %30, sag panel %70 olur.
set "RIGHT=0.70"
if exist "%ROOT%\Data\Temp\Update\split.txt" set /p RIGHT=<"%ROOT%\Data\Temp\Update\split.txt"

rem --------------------------------------------------
rem PROGRAMI BASLAT
rem --------------------------------------------------

set "SESSION=mft%RANDOM%%RANDOM%"
set "PSARGS=-NoLogo -NoProfile -ExecutionPolicy Bypass"

wt.exe -w new --maximized --title "M-Files Toolkit" -d "%ROOT%" powershell.exe %PSARGS% -File "%ROOT%\menu.ps1" -Session %SESSION% ; split-pane -V -s %RIGHT% -d "%ROOT%" powershell.exe %PSARGS% -STA -File "%ROOT%\panel.ps1" -Session %SESSION% ; move-focus left
