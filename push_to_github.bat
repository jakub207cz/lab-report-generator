@echo off
set /p repo_url="Vlozte URL vaseho noveho GitHub repozitare (napr. https://github.com/vase-jmeno/lab-report-generator.git): "

if "%repo_url%"=="" goto error

echo.
echo Pridavam vzdaleny repozitar...
git remote add origin %repo_url%

echo.
echo Prejmenovavam vetev na main...
git branch -M main

echo.
echo Odesilam kod na GitHub...
git push -u origin main

echo.
echo HOTOVO! Nyni bezte na Streamlit Cloud a dejte Deploy.
pause
goto :eof

:error
echo Musite zadat URL repozitare!
pause
