'
'
'          +--[ Intro ]-------------------------------------+
'          |                                                |
'          |  Intro v1.0                                    |
'          |  Copyright (c) 2019 TSC. All Rights Reserved.  |
'          |                                                |
'          |  thomas.schilb@live.de                         |
'          |                                                |
'          +------------[ www.thomasschilb.tk ]-------------+
'
'
'
_SCREENMOVE _MIDDLE
_TITLE "tsc-intro"
DECLARE LIBRARY "since"
END DECLARE
CLS
SCREEN 13
_FULLSCREEN
'LOCATE 2, 1: COLOR 23: PRINT "ANYKEY = Quit     ALT+ENTER = Fullscreen"
LOCATE 25, 1: COLOR 23: PRINT "(c)2019            thomas.schilb@live.de"
DIM WilliamTell& ' handle to hold sound file
WilliamTell& = _SNDOPEN("ts-intro.mp3") '        load MP3 file into RAM
_SNDLOOP WilliamTell&
'_SNDPLAY WilliamTell& '                                   play MP3 file from RAM
'_SNDPLAYFILE "since.mid"
'
' TSC ASCII
'
    FOR i% = 16 TO 31
        COLOR i%
        LOCATE 11, 5: PRINT "   ________ ________.________   "
        LOCATE 12, 5: PRINT "       |    |______ |           "
        LOCATE 13, 5: PRINT "       |    ______| |______     "
        _DELAY 0.1
    NEXT i%
    FOR iii% = 16 TO 31
        COLOR iii%
        LOCATE 15, 5: PRINT "   -  code. design. music.  -"
        _DELAY 0.1
    NEXT iii%

    FOR iiii% = 31 TO 16 STEP -1
        COLOR iiii%
        LOCATE 15, 5: PRINT "   -  code. design. music.  -"
        _DELAY 0.1
    NEXT iiii%

    FOR ii% = 31 TO 16 STEP -1

        COLOR ii%
        LOCATE 11, 5: PRINT "   ________ ________.________   "
        LOCATE 12, 5: PRINT "       |    |______ |           "
        LOCATE 13, 5: PRINT "       |    ______| |______     "
        _DELAY 0.1
    NEXT ii%
LOCATE 25, 1: COLOR 0: PRINT "(c)2012-2017       thomas.schilb@live.de"
'
' 3d wire cube
'
cls
DIM a%(500, 2), b&(8000)
LOCATE 2, 1: COLOR 23: PRINT "ANYKEY = Quit            thomasschilb.de"
LOCATE 25, 1: COLOR 23: PRINT "(c)2019           thomas.schilb@live.de"
LOCATE 20, 18: color 15:PRINT ".t.s.c."
DEF SEG = VARSEG(b&(0))
FOR i% = 0 TO 1499
    a%(i% \ 3 + 1, i% MOD 3) = (INT(RND * 2) * 2 - 1) * INT(50 * COS(RND * 1.570794))
NEXT i%

DO
_DELAY 0.01
    GET (112, 52)-(209, 149), b&()
    FOR i& = VARPTR(b&(1)) TO VARPTR(b&(1)) + 9603
        POKE i&, (PEEK(i&) - 1) * -(PEEK(i&) > 17)
    NEXT i&
    PUT (112, 52), b&(), PSET
    z% = (z% + 1) MOD 576
    FOR i% = 1 TO 500
        nx = (a%(i%, 0) * COS(z% / 91.67325) - a%(i%, 1) * SIN(z% / 91.67325)) * COS(z% / 30.55775) - a%(i%, 2) * SIN(z% / 30.55775)
        ny = (a%(i%, 0) * SIN(z% / 91.67325) + a%(i%, 1) * COS(z% / 91.67325)) * COS(z% / 45.83662) - ((a%(i%, 0) * COS(z% / 91.67325) - a%(i%, 1) * SIN(z% / 91.67325)) * SIN(z% / 30.55775) + a%(i%, 2) * COS(z% / 30.55775)) * SIN(z% / 45.83662)
        nz = (a%(i%, 0) * SIN(z% / 91.67325) + a%(i%, 1) * COS(z% / 91.67325)) * SIN(z% / 45.83662) + ((a%(i%, 0) * COS(z% / 91.67325) - a%(i%, 1) * SIN(z% / 91.67325)) * SIN(z% / 30.55775) + a%(i%, 2) * COS(z% / 30.55775)) * COS(z% / 45.83662)
        IF POINT(160 + nx * 200 / (ny + 400), 100 - nz * 200 / (ny + 400)) < 31 - (ny + 80) \ 15 THEN PSET (160 + nx * 200 / (ny + 400), 100 - nz * 200 / (ny + 400)), 31 - (ny + 80) \ 15
    NEXT i%
LOOP UNTIL INKEY$ <> ""

_SNDCLOSE WilliamTell& '                                  remove MP3 from RAM
SYSTEM
