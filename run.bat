@echo off
rem Spreadsheet loader by J Ramb (2014)

java -cp setup;lib\* xxcust.spread2ws.Excel2WS 2>>error.log %*
