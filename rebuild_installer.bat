del "C:\Users\rlynch\Google Drive\Tableau\Workbook Parsing\installer.7z"
@RD /S /Q "C:\Users\rlynch\Google Drive\Tableau\Workbook Parsing\__pycache__"
@RD /S /Q "C:\Users\rlynch\Google Drive\Tableau\Workbook Parsing\Libra"

cd %~dp0
cd C:\Python35\Scripts
pyinstaller "C:/Users/rlynch/Google Drive/Tableau/Workbook Parsing/Libra.py" --distpath "C:/Users/rlynch/Google Drive/Tableau/Workbook Parsing/" --clean -D -n "Libra" -w -y

cd C:\Program Files\7-Zip\
7z.exe a -t7z "%~dp0\installer" "%~dp0\Libra_Installer.bat" "%~dp0\Libra" "%~dp0\favicon.ico" "%~dp0\Libra_Uninstaller.bat" -mx5

cd %~dp0
copy /b 7zS.sfx + config.txt + installer.7z LibraInstaller.exe /Y
