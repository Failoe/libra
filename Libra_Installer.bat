@echo off

@RD /S /Q C:\\"Tableau Scripts\Libra"
md C:\\"Tableau Scripts\Libra"

xcopy %cd%\Libra C:\"Tableau Scripts\Libra"\ /S /Y /Q
xcopy %cd%\favicon.ico C:\"Tableau Scripts"\ /Y /Q
xcopy %cd%\Libra_Uninstaller.bat C:\"Tableau Scripts"\ /Y /Q

REG ADD HKEY_CURRENT_USER\Software\Classes\Tableau.Workbook.2\shell\\"Libra"\command /d "C:\Tableau Scripts\Libra\Libra.exe \"%%1\"" /f
REG ADD HKEY_CURRENT_USER\Software\Classes\Tableau.PackagedWorkbook.2\shell\\"Libra"\command /d "C:\Tableau Scripts\Libra\Libra.exe \"%%1\"" /f

REG delete HKEY_CURRENT_USER\Software\Classes\Tableau.Workbook.2\shell\\"Workbook Version" /f
REG delete HKEY_CURRENT_USER\Software\Classes\Tableau.PackagedWorkbook.2\shell\\"Workbook Version" /f
REG delete HKEY_CURRENT_USER\Software\Classes\Tableau.Workbook.2\shell\\"Check Version" /f
REG delete HKEY_CURRENT_USER\Software\Classes\Tableau.PackagedWorkbook.2\shell\\"Check Version" /f
