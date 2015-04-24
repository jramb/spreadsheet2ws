@echo off
rem Spreadsheet loader by J Ramb (2014-2015)
rem https://github.com/jramb/spreadsheet2ws

set DIRNAME=%~dp0
if "%DIRNAME%" == "" set DIRNAME=.
set APP_BASE_NAME=%~n0
set APP_HOME=%DIRNAME%..
set CMD_LINE_ARGS=%*

REM echo %DIRNAME%\bin\spreadsheet2ws %CMD_LINE_ARGS%
%DIRNAME%bin\spreadsheet2ws.bat %CMD_LINE_ARGS%

