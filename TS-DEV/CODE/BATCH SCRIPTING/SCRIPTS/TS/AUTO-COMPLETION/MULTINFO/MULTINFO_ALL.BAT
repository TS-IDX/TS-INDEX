@ECHO OFF
ECHO MULTINFO V0.3
ECHO.
ECHO FILLING ALL SUB-FOLDERS WITH NFO-DATA...
ECHO.
@for /d %%X in (*) do "xcopy.exe" "TS.NFO" "%%X\" >NUL
ECHO + NFO
@for /d %%X in (*) do "xcopy.exe" "TS.DIZ" "%%X\" >NUL
ECHO + DIZ
@for /d %%X in (*) do "xcopy.exe" "TS.URL" "%%X\" >NUL
ECHO + URL
@for /d %%X in (*) do "xcopy.exe" "TS.PNG" "%%X\" >NUL
ECHO + PNG
ECHO.
ECHO OK
ECHO.