since.ascii.cyber.mp3 - background music, must be in the same directory as program
since.ico             - since standard icon file
since.o               - library file containing the icon, must be in same directory as qb64.exe
sinceA.rc             - standard resource file

to convert your own icon to library file use windres e.g.:

C:\qb64\internal\c\c_compiler\bin\windres -i sinceA.rc -o since.o

note: icon is only used for the binary not in title window