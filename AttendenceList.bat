@echo off
set source="C:\Users\rg36218.ZENDER\Downloads"
set target="C:\Users\rg36218.ZENDER\OneDrive - Zensar Technologies Ltd\Derivco\Flows"

FOR /F "delims=" %%I IN ('DIR %source%\*.* /A:-D /O:-D /B') DO COPY %source%\"%%I" %target% & echo %%I & GOTO :END
:END
echo "----------------------- File is saved to One drive successfully !!-------------------------------------------"
TIMEOUT 15