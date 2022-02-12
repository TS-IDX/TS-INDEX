@ECHO OFF
ECHO.
ECHO Compressing all Sub-Folders to 7Z...
ECHO.
@for /d %%X in (*) do "7z.exe" a "%%X.7Z" "%%X\"
ECHO.
ECHO OK
ECHO.