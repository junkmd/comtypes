@echo off
rem AvmcIfc.dll のビルド
rem 使い方: buid_avmc.bat [x86|x64] [Debug|Release]
rem   デフォルト: x86 Debug

setlocal

set "ARCH=%~1"
if "%ARCH%"=="" set "ARCH=x86"

set "CONFIG=%~2"
if "%CONFIG%"=="" set "CONFIG=Debug"

rem VS のインストールパスを決定
set "VSWHERE=%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe"
if not exist "%VSWHERE%" set "VSWHERE=vswhere"

for /f "usebackq delims=" %%i in (`"%VSWHERE%" -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath`) do set "VS_PATH=%%i"

if not defined VS_PATH (
    echo [buid_avmc] Visual Studio のインストールパスを取得できませんでした。
    exit /b 1
)
echo [buid_avmc] VS_PATH=%VS_PATH%

rem vcvarsall.bat の存在確認
set "VCVARS=%VS_PATH%\VC\Auxiliary\Build\vcvarsall.bat"
if not exist "%VCVARS%" (
    echo [buid_avmc] vcvarsall.bat が見つかりません: %VCVARS%
    exit /b 1
)

rem VS 環境を初期化してビルド
call "%VCVARS%" %ARCH%
if errorlevel 1 exit /b 1

msbuild AvmcIfc.vcxproj /p:Configuration=%CONFIG% /p:Platform=%ARCH%
exit /b %errorlevel%