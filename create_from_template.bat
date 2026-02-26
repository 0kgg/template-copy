@echo off
chcp 932 >nul
setlocal EnableDelayedExpansion

REM ==========================================
REM 基本パス
REM ==========================================
set "BASE_DIR=%~dp0"
set "CONFIG_PATH=%BASE_DIR%settings.ini"

REM ==========================================
REM 初回設定
REM settings.ini がなければ作成
REM ==========================================
if not exist "%CONFIG_PATH%" goto INIT_CONFIG
goto LOAD_CONFIG

:INIT_CONFIG
echo 初回設定を開始します．
echo.

:INPUT_TEMPLATE
set "TEMPLATE_NAME="
set /p "TEMPLATE_NAME=テンプレートファイル名を入力してください（例: tmp.pptx）: "

if "%TEMPLATE_NAME%"=="" (
    echo テンプレートファイル名が空です．
    goto INPUT_TEMPLATE
)

call :VALIDATE_NAME "%TEMPLATE_NAME%"
if errorlevel 1 goto INPUT_TEMPLATE

if not exist "%BASE_DIR%%TEMPLATE_NAME%" (
    echo 指定されたテンプレートファイルが見つかりません．
    echo "%BASE_DIR%%TEMPLATE_NAME%"
    echo 同じフォルダに置いてから入力してください．
    echo.
    goto INPUT_TEMPLATE
)

:INPUT_AUTHOR
set "AUTHOR_NAME="
set /p "AUTHOR_NAME=作成者名を入力してください: "

if "%AUTHOR_NAME%"=="" (
    echo 作成者名が空です．
    goto INPUT_AUTHOR
)

call :VALIDATE_NAME "%AUTHOR_NAME%"
if errorlevel 1 goto INPUT_AUTHOR

(
    echo TEMPLATE_NAME=%TEMPLATE_NAME%
    echo AUTHOR_NAME=%AUTHOR_NAME%
) > "%CONFIG_PATH%"

echo.
echo 設定ファイルを作成しました．
echo "%CONFIG_PATH%"
echo.
goto LOAD_CONFIG

REM ==========================================
REM 設定読み込み
REM ==========================================
:LOAD_CONFIG
for /f "usebackq tokens=1,* delims==" %%A in ("%CONFIG_PATH%") do (
    if /i "%%A"=="TEMPLATE_NAME" set "TEMPLATE_NAME=%%B"
    if /i "%%A"=="AUTHOR_NAME" set "AUTHOR_NAME=%%B"
)

if "%TEMPLATE_NAME%"=="" (
    echo settings.ini の TEMPLATE_NAME が読めません．
    pause
    exit /b 1
)

if "%AUTHOR_NAME%"=="" (
    echo settings.ini の AUTHOR_NAME が読めません．
    pause
    exit /b 1
)

set "TEMPLATE_PATH=%BASE_DIR%%TEMPLATE_NAME%"

if not exist "%TEMPLATE_PATH%" (
    echo テンプレートファイルが見つかりません．
    echo "%TEMPLATE_PATH%"
    pause
    exit /b 1
)

for %%F in ("%TEMPLATE_PATH%") do set "EXT=%%~xF"

for /f %%I in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd"') do set "TODAY=%%I"

REM ==========================================
REM メイン処理
REM ==========================================
:INPUT_PROJECT
set "PROJECT_NAME="
set /p "PROJECT_NAME=プロジェクト名を入力してください: "

if "%PROJECT_NAME%"=="" (
    echo プロジェクト名が空です．
    goto INPUT_PROJECT
)

call :VALIDATE_NAME "%PROJECT_NAME%"
if errorlevel 1 goto INPUT_PROJECT

call :BUILD_TARGET "%PROJECT_NAME%" TARGET_PATH

if not exist "%TARGET_PATH%" goto CREATE_FILE

:HANDLE_DUPLICATE
call :BUILD_BASE_NAME "%PROJECT_NAME%" BASE_NAME
call :FIND_NEXT_SUFFIX "%BASE_NAME%" NEXT_PATH

echo.
echo 同名ファイルが存在します．
echo 既存: "%TARGET_PATH%"
echo 候補: "%NEXT_PATH%"
echo.
echo Enter: 連番で作成
echo N    : 再入力名を入力
echo C    : 中止
set "DUP_ACTION="
set /p "DUP_ACTION=選択してください [Enter/N/C]: "

if /i "%DUP_ACTION%"=="C" (
    echo 処理を中止しました．
    pause
    exit /b 0
)

if /i "%DUP_ACTION%"=="N" goto INPUT_NEW_PROJECT

set "TARGET_PATH=%NEXT_PATH%"
goto CREATE_FILE

:INPUT_NEW_PROJECT
set "REINPUT_NAME="
set /p "REINPUT_NAME=再入力名を入力してください: "

if "%REINPUT_NAME%"=="" (
    echo 再入力名が空です．
    goto INPUT_NEW_PROJECT
)

call :VALIDATE_NAME "%REINPUT_NAME%"
if errorlevel 1 goto INPUT_NEW_PROJECT

call :BUILD_TARGET "%REINPUT_NAME%" TARGET_PATH

if exist "%TARGET_PATH%" (
    echo そのファイル名も既に存在します．
    echo "%TARGET_PATH%"
    echo.
    goto INPUT_NEW_PROJECT
)

goto CREATE_FILE

:CREATE_FILE
echo.
echo 作成先:
echo "%TARGET_PATH%"

copy "%TEMPLATE_PATH%" "%TARGET_PATH%" >nul
if errorlevel 1 (
    echo ファイル作成に失敗しました．
    pause
    exit /b 1
)

echo.
echo 作成しました．
echo "%TARGET_PATH%"

start "" "%TARGET_PATH%"

echo.
echo 3秒後にこの画面を閉じます．
timeout /t 3 /nobreak >nul
exit /b 0

REM ==========================================
REM YYYYMMDD-名前-作成者名 を組み立て
REM ==========================================
:BUILD_BASE_NAME
setlocal
set "NAME_PART=%~1"
set "OUT=%TODAY%-%NAME_PART%-%AUTHOR_NAME%"
endlocal & set "%~2=%OUT%"
goto :eof

REM ==========================================
REM 完全パスを組み立て
REM ==========================================
:BUILD_TARGET
setlocal
set "NAME_PART=%~1"
set "OUT=%BASE_DIR%%TODAY%-%NAME_PART%-%AUTHOR_NAME%%EXT%"
endlocal & set "%~2=%OUT%"
goto :eof

REM ==========================================
REM 禁止文字チェック
REM 禁止文字 \ / : * ? " < > |
REM ==========================================
:VALIDATE_NAME
setlocal
set "CHK=%~1"

echo(%CHK%| findstr /c:"\" >nul && goto NG
echo(%CHK%| findstr /c:"/" >nul && goto NG
echo(%CHK%| findstr /c:":" >nul && goto NG
echo(%CHK%| findstr /c:"*" >nul && goto NG
echo(%CHK%| findstr /c:"?" >nul && goto NG
echo(%CHK%| findstr /c:"""" >nul && goto NG
echo(%CHK%| findstr /c:"<" >nul && goto NG
echo(%CHK%| findstr /c:">" >nul && goto NG
echo(%CHK%| findstr /c:"|" >nul && goto NG

endlocal
exit /b 0

:NG
echo.
echo 次の文字はファイル名に使えません．
echo \ / : * ? " ^< ^> ^|
echo 入力し直してください．
echo.
endlocal
exit /b 1

REM ==========================================
REM 空いている連番を探す
REM ==========================================
:FIND_NEXT_SUFFIX
setlocal EnableDelayedExpansion
set "FN_BASE=%~1"
set /a NUM=2

:FIND_LOOP
set "CANDIDATE=%BASE_DIR%%FN_BASE%_!NUM!!EXT!"
if exist "!CANDIDATE!" (
    set /a NUM+=1
    goto FIND_LOOP
)

endlocal & set "%~2=%CANDIDATE%"
goto :eof
