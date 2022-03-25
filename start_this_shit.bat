echo OFF

NET SESSION >nul 2>&1

IF %ERRORLEVEL% EQU 0 (

   echo.

) ELSE (

   echo.-------------------------------------------------------------

   echo ERROR: YOU ARE NOT RUNNING THIS WITH ADMINISTRATOR PRIVILEGES.

   echo. -------------------------------------------------------------

   echo. If you're seeing this, it means you don't have admin privileges!

   pause

   echo.

   echo. You will need to restart this program with Administrator privileges by right-clicking and select "Run As Administrator"

   pause 

    echo.

   echo Press any key to leave this program. Make sure to Run As Administrator next time!

   pause

   EXIT /B 1

)

powershell -NoProfile -InputFormat None -ExecutionPolicy Bypass -Command "[System.Net.ServicePointManager]::SecurityProtocol = 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))" && SET "PATH=%PATH%;%ALLUSERSPROFILE%\chocolatey\bin"

choco feature enable -n=allowGlobalConfirmation

echo Chocolatey is ready to begin installing packages!

pause

choco install python
choco install git

echo python %CD%\code\main.py > start.bat

mkdir code
cd code

echo Download Programm
git clone https://github.com/tobi187/efly_combine_reports.git

echo Download Packages
python -m pip install efly_combine_reports\requirements.txt

echo "finshed"

