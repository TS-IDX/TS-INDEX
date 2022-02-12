'
'
'          +--[ WireCube ]----------------------------------+
'          |                                                |
'          |  WireCube v2.1.0                               |
'          |  Copyright (c) 2019 TSC. All Rights Reserved.  |
'          |                                                |
'          |  thomas.schilb@live.de                         |
'          |                                                |
'          +------------[ www.thomasschilb.tk ]-------------+
'
'
'
DECLARE LIBRARY "since"
END DECLARE

DIM a%(500, 2), b&(8000)
SCREEN 13
_FULLSCREEN
COLOR 15: LOCATE 20, 18: PRINT "+-------[ www.thomasschilb.tk ]--------+"
COLOR 15: LOCATE 4, 18: PRINT "+-------------[ WireCube ]-------------+"
COLOR 15: LOCATE 6: PRINT "|                                      |"
COLOR 15: LOCATE 6: PRINT "|                                      |"
COLOR 15: LOCATE 6: PRINT "|                                      |"
COLOR 15: LOCATE 7: PRINT "|                                      |"
COLOR 15: LOCATE 8: PRINT "|                                      |"
COLOR 15: LOCATE 9: PRINT "|                                      |"
COLOR 15: LOCATE 10: PRINT "|                                      |"
COLOR 15: LOCATE 11: PRINT "|                                      |"
COLOR 15: LOCATE 12: PRINT "|                                      |"
COLOR 15: LOCATE 13: PRINT "|                                      |"
COLOR 15: LOCATE 14: PRINT "|                                      |"
COLOR 15: LOCATE 15: PRINT "|                                      |"
COLOR 15: LOCATE 16: PRINT "|                                      |"
COLOR 15: LOCATE 17: PRINT "|                                      |"
COLOR 15: LOCATE 18: PRINT "|                                      |"
COLOR 15: LOCATE 19: PRINT "|                                      |"
COLOR 15: LOCATE 20: PRINT "|                                      |"


DEF SEG = VARSEG(b&(0))
FOR i% = 0 TO 1499
    _DELAY .001
    a%(i% \ 3 + 1, i% MOD 3) = (INT(RND * 2) * 2 - 1) * INT(50 * COS(RND * 1.570794))
NEXT i%
DO
    _DELAY 0.001
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
SYSTEM
