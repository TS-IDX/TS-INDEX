@ECHO OFF
ECHO.
ECHO Compressing all Sub-Folders to ZIP...
ECHO.
@for /d %%X in (*) do "7z.exe" a -tzip "%%X.ZIP" "%%X\"
ECHO.
ECHO OK
ECHO.