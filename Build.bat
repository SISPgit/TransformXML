@echo off
chcp 65001 > nul
echo Pradedamas EXE failo kūrimas...

REM Tikriname, ar egzistuoja TransformXML2.py failas
if not exist TransformXML2.py (
    echo Klaida: TransformXML2.py failas nerastas.
    echo Įsitikinkite, kad šis BAT failas yra toje pačioje direktorijoje kaip ir TransformXML2.py.
    pause
    exit /b
)

REM Sukuriame EXE failą
pyinstaller --onefile --noconsole --hidden-import=smtplib --hidden-import=email.mime.multipart --hidden-import=email.mime.text TransformXML2.py

REM Tikriname, ar EXE failas buvo sėkmingai sukurtas
if exist dist\TransformXML2.exe (
    echo.
    echo EXE failas sėkmingai sukurtas. Jį rasite dist direktorijoje.
    echo.
    echo Norėdami paleisti programą, eikite į dist direktoriją ir paleiskite TransformXML2.exe.
) else (
    echo.
    echo Klaida kuriant EXE failą. Patikrinkite klaidas aukščiau.
)

echo.
pause