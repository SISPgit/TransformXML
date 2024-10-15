Copy@echo off
echo Diegiamos reikalingos bibliotekos...
REM Įsitikiname, kad pip yra atnaujintas
python -m pip install --upgrade pip
REM Diegiame reikalingas bibliotekas
pip install wheel
pip install paramiko
pip install pandas
pip install lxml
pip install deepdiff
pip install openpyxl
pip install pyinstaller
echo.
echo Bibliotekų diegimas baigtas.
pause