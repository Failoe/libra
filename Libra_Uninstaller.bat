@echo off

@RD /S /Q "C:\Tableau Scripts\Libra"
DEL C:\"Tableau Scripts\favicon.ico"

REG delete HKEY_CURRENT_USER\Software\Classes\Tableau.Workbook.2\shell\\"Libra" /f
REG delete HKEY_CURRENT_USER\Software\Classes\Tableau.PackagedWorkbook.2\shell\\"Libra" /f
DEL "%~f0"