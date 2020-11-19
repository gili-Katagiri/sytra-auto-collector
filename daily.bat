@echo off
cd /d %~dp0

setlocal
set stockpath=
set wadpath=
set mspath=
set rsspath=

rem import summary_base.csv & lf to crlf
scp %stockpath%/summary_base.csv .\sbase.csv
if errorlevel 1 (
    @echo on 
    echo ImportError
    exit /B
)
find <sbase.csv /V "irohanihoheto chirinuruwo" >summary_base.csv

rem today is next update day?
FOR /F %%i in ('findstr /r "[0-9]*-[0-9]*-[0-9]*" summary_base.csv') do set u_day=%%i
call python dayat.py %u_day%

if "%1"=="F" ( 
    @echo on
    echo F OPTION True
    @echo off
) else (
    if errorlevel 3 (
        @echo on
        echo FALSE
        exit /B
    )
)

rem start Applications
start "WAD" %wadpath%
%mspath%
%rsspath%
start transer.xlsm

call python mksummary.py

rem send summary
scp summary.csv %stockpath%/summary.csv
if not errorlevel 1 ( del sbase.csv summary_base.csv summary.csv )
endlocal
