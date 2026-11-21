# AvmcIfc (InProc COM サーバー) の自前ビルド

`src/source/Avmc` にある ATL 製 InProc COM サーバー `AvmcIfc` を、リポジトリ内の
`tools/build_avmc_server.bat` で自前ビルドするための手順をまとめる。

## 目的

`comtypes` のテスト (`src/comtypes/test/test_avmc.py`) は、`AvmcIfc.Avmc.1` という
COM オブジェクトを `CreateObject` で生成し、`FindAllAvmc()` が返す `DeviceInfo`
レコードの SAFEARRAY を検証する。このテストを動かすには、テストダブルである
`AvmcIfc.dll` (x86) / `AvmcIfc_x64.dll` (x64) をソースからビルドして登録しておく
必要がある。

従来は `src/source/Avmc/build-dll.py` が `jaraco.develop` の `vstudio` ヘルパー経由で
`msbuild` を呼んでいたが、この方法は Python パッケージ (`jaraco.develop`) に依存し、
CI ランナーで追加インストールが必要になる。そこで、VS の検出・ATL の追加・
ツールセットの再ターゲットまでを 1 つのバッチファイルにまとめ、追加の Python
依存なしでビルドできるようにした。

## 使い方

リポジトリ直下から、`tools` ディレクトリに移動して実行する。

```cmd
cd tools
build_avmc_server.bat
```

VS のインストールパスを自動検出できない場合は、引数で明示的に指定する。

```cmd
cd tools
build_avmc_server.bat "C:\Program Files\Microsoft Visual Studio\18\Enterprise"
```

### 出力

ビルドに成功すると、以下が生成される。

| ファイル | アーキテクチャ |
|---|---|
| `src\source\Avmc\Debug\AvmcIfc.dll` | x86 (Win32) |
| `src\source\Avmc\Debug\AvmcIfc_x64.dll` | x64 |

### スクリプトが行うこと

1. `vswhere` で VS のインストールパスを検出する (引数指定があればそちらを優先)。
2. ATL (`atlmfc\include\atlbase.h`) の有無を確認し、無ければ VS インストーラ
   (`setup.exe modify --add Microsoft.VisualStudio.Component.VC.ATL`) で追加する。
3. `vcvarsall.bat` の存在を確認する。
4. インストール済みの最新 Platform Toolset (例: `v145`) を検出する。
5. `vcvarsall x86` で環境を初期化し、`Debug|Win32` をビルドする。
6. `vcvarsall x64` で環境を初期化し、`Debug|x64` をビルドする。

### 注意点

- `AvmcIfc.vcxproj` は `ToolsVersion 4.0` のため、既定では Platform Toolset
  `v100` (VS2010) を要求する。スクリプトはインストール済みの最新ツールセットを
  検出して `/p:PlatformToolset=` で上書きするため、VS2010 のビルドツールは不要。
- ビルドの `CustomBuildStep` が `regsvr32` で DLL を登録する。x86 と x64 は同じ
  `Debug` ディレクトリと `regsvr32.trg` を共有するため、2 回目のビルドでは登録が
  スキップされる (登録は `tools/run_tests.bat` が別途行うため実害はない)。
- 登録 (`regsvr32`) には管理者権限が必要。権限が無い環境ではビルドが失敗する
  可能性がある。

## スクリプト内容 (`tools/build_avmc_server.bat`)

```bat
@echo off
rem AvmcIfc (InProc COM サーバー) のビルド
rem
rem 使い方:
rem   build_avmc_server.bat                  VS のインストールパスを vswhere で自動検出
rem   build_avmc_server.bat "<VSインストールパス>"  検出結果を引数で指定 (例: C:\Program Files\Microsoft Visual Studio\18\Enterprise)
rem
rem 出力:
rem   src\source\Avmc\Debug\AvmcIfc.dll      (x86 / Win32)
rem   src\source\Avmc\Debug\AvmcIfc_x64.dll  (x64)

setlocal

rem --- 1. VS のインストールパスを決定 ---
set "VS_PATH=%~1"
if defined VS_PATH goto :have_vs_path

set "VSWHERE=%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe"
if not exist "%VSWHERE%" set "VSWHERE=vswhere"

for /f "usebackq delims=" %%i in (`"%VSWHERE%" -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath`) do set "VS_PATH=%%i"

:have_vs_path
if not defined VS_PATH (
    echo [build_avmc_server] Visual Studio のインストールパスを取得できませんでした。
    echo                     引数で指定してください: build_avmc_server.bat "^<VSインストールパス^>"
    exit /b 1
)
echo [build_avmc_server] VS_PATH=%VS_PATH%

rem --- 2. ATL の有無を確認し、無ければインストール ---
rem AvmcIfc は ATL (atlbase.h / atlcom.h) に依存する。MSVC 本体だけでは
rem ATL ヘッダが入っていないことがあるため、無ければ VS インストーラで追加する。
set "ATL_FOUND="
for /d %%d in ("%VS_PATH%\VC\Tools\MSVC\*") do (
    if exist "%%~d\atlmfc\include\atlbase.h" set "ATL_FOUND=1"
)
if not defined ATL_FOUND (
    echo [build_avmc_server] ATL が見つかりません。VS インストーラで追加します...
    set "VS_SETUP=%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\setup.exe"
    if not exist "%VS_SETUP%" (
        echo [build_avmc_server] setup.exe が見つかりません: %VS_SETUP%
        exit /b 1
    )
    "%VS_SETUP%" modify --installPath "%VS_PATH%" --add Microsoft.VisualStudio.Component.VC.ATL --quiet --norestart --wait
    if errorlevel 1 (
        echo [build_avmc_server] ATL のインストールに失敗しました。
        exit /b 1
    )
    echo [build_avmc_server] ATL のインストールが完了しました。
)

rem --- 3. vcvarsall.bat の存在確認 ---
set "VCVARS=%VS_PATH%\VC\Auxiliary\Build\vcvarsall.bat"
if not exist "%VCVARS%" (
    echo [build_avmc_server] vcvarsall.bat が見つかりません: %VCVARS%
    exit /b 1
)

rem --- 3b. Platform Toolset を検出 ---
rem AvmcIfc.vcxproj は ToolsVersion 4.0 のため既定で v100 (VS2010) を要求する。
rem インストール済みの最新ツールセット (例: v145) を検出して上書きする。
set "TOOLSET="
for /d %%d in ("%VS_PATH%\MSBuild\Microsoft\VC\v*") do (
    for /d %%t in ("%%~d\Platforms\x64\PlatformToolsets\v*") do set "TOOLSET=%%~nxt"
)
if not defined TOOLSET (
    echo [build_avmc_server] Platform Toolset を検出できませんでした。
    exit /b 1
)
echo [build_avmc_server] PlatformToolset=%TOOLSET%

rem --- 4. Avmc ディレクトリへ移動 (スクリプトの位置基準) ---
cd /d "%~dp0..\src\source\Avmc"
if errorlevel 1 (
    echo [build_avmc_server] Avmc ディレクトリへ移動できませんでした。
    exit /b 1
)

rem --- 5. x86 (Win32) をビルド ---
call "%VCVARS%" x86
if errorlevel 1 exit /b 1

msbuild AvmcIfc.sln /p:Configuration=Debug /p:Platform=Win32 /p:PlatformToolset=%TOOLSET%
if errorlevel 1 exit /b 1

rem --- 6. x64 をビルド ---
call "%VCVARS%" x64
if errorlevel 1 exit /b 1

msbuild AvmcIfc.sln /p:Configuration=Debug /p:Platform=x64 /p:PlatformToolset=%TOOLSET%
if errorlevel 1 exit /b 1

echo [build_avmc_server] ビルド完了: Debug\AvmcIfc.dll (x86) / Debug\AvmcIfc_x64.dll (x64)
exit /b 0
```

## 検証結果

Windows Server 2025 + Visual Studio 2026 (18.12.0, Platform Toolset `v145`) で
`build_avmc_server.bat` を実行し、`Debug|Win32` と `Debug|x64` の両方が
`Build succeeded` となることを確認した。生成物:

- `src\source\Avmc\Debug\AvmcIfc.dll` (x86)
- `src\source\Avmc\Debug\AvmcIfc_x64.dll` (x64)

ビルド時の警告 (`strcpy` の C4996、`FT_HANDLE` → `ULONG` の C4311/C4302、
`AvmcIfc_x64.dll` と `.def` の `LIBRARY` 名の不一致による LNK4070/MSB8012) は
いずれも既存ソース由来で、ビルドは成功する。
