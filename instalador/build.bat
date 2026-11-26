@echo off
rem ==========================================================================
rem  LT Tools - gera o instalador (instalador\LTTools_Instalador.exe)
rem  Usa o csc.exe que j  vem no Windows (.NET Framework 4).
rem  Antes: atualize o plugin no CorelDRAW (botÆo Atualizar) e FECHE o Corel,
rem  para o LTPlugin.gms da pasta GMS estar com a versÆo nova.
rem ==========================================================================
setlocal enabledelayedexpansion
set R=%~dp0..
set T=%TEMP%\ltb
set CSC=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe
rmdir /s /q "%T%" 2>nul
mkdir "%T%"

rem pega o LTPlugin.gms instalado no CorelDRAW (Graphics ou Technical Suite)
set GMS=
for /d %%D in ("%APPDATA%\Corel\*") do if exist "%%D\Draw\GMS\LTPlugin.gms" set GMS=%%D\Draw\GMS\LTPlugin.gms
if "%GMS%"=="" (
  if exist "%~dp0LTPlugin.gms" ( set GMS=%~dp0LTPlugin.gms ) else ( echo LTPlugin.gms nao encontrado & exit /b 1 )
)
echo GMS: %GMS%
copy /y "%GMS%" "%T%\LTPlugin.gms" >nul
copy /y "%GMS%" "%~dp0LTPlugin.gms" >nul

copy /y "%R%\icones\*.*" "%T%\" >nul
copy /y "%R%\LTTools.bas" "%T%\" >nul
copy /y "%R%\LTUpdate.bas" "%T%\" >nul
copy /y "%R%\frmLT_codigo.txt" "%T%\" >nul
copy /y "%~dp0Instalador.cs" "%T%\" >nul

set RSP=%T%\build.rsp
> "%RSP%" echo /nologo
>>"%RSP%" echo /target:winexe
>>"%RSP%" echo /out:%T%\LTTools_Instalador.exe
>>"%RSP%" echo /win32icon:%T%\logo.ico
>>"%RSP%" echo /reference:System.Windows.Forms.dll
for %%F in ("%T%\*.bas" "%T%\*.txt" "%T%\*.gms" "%T%\*.png" "%T%\*.bmp" "%T%\*.ico") do >>"%RSP%" echo /resource:%%~fF,%%~nxF
>>"%RSP%" echo %T%\Instalador.cs

"%CSC%" @"%RSP%"
if errorlevel 1 ( echo ERRO ao compilar & exit /b 1 )
copy /y "%T%\LTTools_Instalador.exe" "%~dp0LTTools_Instalador.exe" >nul
echo OK: %~dp0LTTools_Instalador.exe
