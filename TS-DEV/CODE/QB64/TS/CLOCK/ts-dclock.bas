'
'
'          +--[ dClock ]------------------------------------+
'          |                                                |
'          |  digitalClock v1.0.1                           |
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

DECLARE SUB FELT (X%, Y%, XX%, YY%, FARBEN%)
DECLARE SUB KLPRINT (B$, B%, K%, G%, I%)
DECLARE SUB DATUM ()
DECLARE SUB FELD (A%, B%, AA%, BB%, C%)
_FULLSCREEN
SCREEN 12
WIDTH 80, 60
FELT 0, 0, 639, 479, 0
FELT 36, 78, 597, 154, 0

'CONST PI = 3.141592
DIM SHARED PI
PI = 3.141592

XCXX = VAL(MID$(TIME$, 7, 2))
XCXY = VAL(MID$(TIME$, 4, 2))
XCXZ = VAL(MID$(TIME$, 1, 2))
1.1
XCXB = PI * 2 / 12 - .0001
XCXD = PI * 2 / 60 - .0001
XCXE = PI * 2 / 60 - .0001
456
IF INKEY$ <> "" THEN END
IF XCXX <> VAL(MID$(TIME$, 7, 2)) THEN XCXX = VAL(MID$(TIME$, 7, 2)): GFF = 1 ELSE GFF = 0: GOTO 456
IF XCXY <> VAL(MID$(TIME$, 4, 2)) THEN XCXY = VAL(MID$(TIME$, 4, 2)): GFG = 1 ELSE GFG = 0
IF XCXZ <> VAL(MID$(TIME$, 1, 2)) THEN XCXZ = VAL(MID$(TIME$, 1, 2)): GFH = 1 ELSE GFH = 0
XCXC = VAL(MID$(TIME$, 1, 2))
XCXF = VAL(MID$(TIME$, 4, 2))
XCXG = VAL(MID$(TIME$, 7, 2))
IF XCXC > 12 THEN XCXC = XCXC - 12
IF XCXC < 4 THEN XCXC = XCXC + 12
IF XCXC > 3 THEN XCXC = XCXC - 3
IF XCXF < 16 THEN XCXF = XCXF + 60
IF XCXF > 15 THEN XCXF = XCXF - 15
IF XCXG < 16 THEN XCXG = XCXG + 60
IF XCXG > 15 THEN XCXG = XCXG - 15
XCXA = XCXB * XCXC
XCXH = XCXD * XCXF
XCXI = XCXE * XCXG
XCXAA = XCXA
XCXHH = XCXH
XCXII = XCXI
IF TEMPTIME$ <> TIME$ THEN TEMPTIME$ = TIME$: DATUM
KLPRINT MID$(DATE$, 4, 2) + ".", 18, 576, 0, 4
KLPRINT MID$(DATE$, 1, 2) + ".", 18, 594, 0, 4
KLPRINT MID$(DATE$, 7, 4), 18, 611, 0, 4
KLPRINT "", 18, 0, 0, 4
COLOR 15
A = -1
B = -1
C = -1
D = -1
E = -1
F = -1
LOCATE 27, 27: PRINT "лл"
LOCATE 28, 27: PRINT "лл"
LOCATE 29, 51: PRINT "лл"
LOCATE 30, 51: PRINT "лл"
20
IF A <> VAL(MID$(TIME$, 1, 1)) THEN A = VAL(MID$(TIME$, 1, 1)): GOTO 2 ELSE GOTO 1
2
SELECT CASE A
    CASE 0:
        LOCATE 25, 7: PRINT "лллллллл"
        LOCATE 26, 7: PRINT "лл    лл"
        LOCATE 27, 7: PRINT "лл    лл"
        LOCATE 28, 7: PRINT "лл    лл"
        LOCATE 29, 7: PRINT "лл    лл"
        LOCATE 30, 7: PRINT "лл    лл"
        LOCATE 31, 7: PRINT "лллллллл"
    CASE 1
        LOCATE 25, 7: PRINT "      лл"
        LOCATE 26, 7: PRINT "      лл"
        LOCATE 27, 7: PRINT "      лл"
        LOCATE 28, 7: PRINT "      лл"
        LOCATE 29, 7: PRINT "      лл"
        LOCATE 30, 7: PRINT "      лл"
        LOCATE 31, 7: PRINT "      лл"
    CASE 2:
        LOCATE 25, 7: PRINT "лллллллл"
        LOCATE 26, 7: PRINT "      лл"
        LOCATE 27, 7: PRINT "      лл"
        LOCATE 28, 7: PRINT "лллллллл"
        LOCATE 29, 7: PRINT "лл      "
        LOCATE 30, 7: PRINT "лл      "
        LOCATE 31, 7: PRINT "лллллллл"
END SELECT
1
IF B <> VAL(MID$(TIME$, 2, 1)) THEN B = VAL(MID$(TIME$, 2, 1)): GOTO 4 ELSE GOTO 3
4
SELECT CASE B
    CASE 0:
        LOCATE 25, 17: PRINT "лллллллл"
        LOCATE 26, 17: PRINT "лл    лл"
        LOCATE 27, 17: PRINT "лл    лл"
        LOCATE 28, 17: PRINT "лл    лл"
        LOCATE 29, 17: PRINT "лл    лл"
        LOCATE 30, 17: PRINT "лл    лл"
        LOCATE 31, 17: PRINT "лллллллл"
    CASE 1
        LOCATE 25, 17: PRINT "      лл"
        LOCATE 26, 17: PRINT "      лл"
        LOCATE 27, 17: PRINT "      лл"
        LOCATE 28, 17: PRINT "      лл"
        LOCATE 29, 17: PRINT "      лл"
        LOCATE 30, 17: PRINT "      лл"
        LOCATE 31, 17: PRINT "      лл"
    CASE 2:
        LOCATE 25, 17: PRINT "лллллллл"
        LOCATE 26, 17: PRINT "      лл"
        LOCATE 27, 17: PRINT "      лл"
        LOCATE 28, 17: PRINT "лллллллл"
        LOCATE 29, 17: PRINT "лл      "
        LOCATE 30, 17: PRINT "лл      "
        LOCATE 31, 17: PRINT "лллллллл"
    CASE 3:
        LOCATE 25, 17: PRINT "лллллллл"
        LOCATE 26, 17: PRINT "      лл"
        LOCATE 27, 17: PRINT "      лл"
        LOCATE 28, 17: PRINT "лллллллл"
        LOCATE 29, 17: PRINT "      лл"
        LOCATE 30, 17: PRINT "      лл"
        LOCATE 31, 17: PRINT "лллллллл"
    CASE 4:
        LOCATE 25, 17: PRINT "лл    лл"
        LOCATE 26, 17: PRINT "лл    лл"
        LOCATE 27, 17: PRINT "лл    лл"
        LOCATE 28, 17: PRINT "лллллллл"
        LOCATE 29, 17: PRINT "      лл"
        LOCATE 30, 17: PRINT "      лл"
        LOCATE 31, 17: PRINT "      лл"
    CASE 5:
        LOCATE 25, 17: PRINT "лллллллл"
        LOCATE 26, 17: PRINT "лл      "
        LOCATE 27, 17: PRINT "лл      "
        LOCATE 28, 17: PRINT "лллллллл"
        LOCATE 29, 17: PRINT "      лл"
        LOCATE 30, 17: PRINT "      лл"
        LOCATE 31, 17: PRINT "лллллллл"
    CASE 6:
        LOCATE 25, 17: PRINT "лллллллл"
        LOCATE 26, 17: PRINT "лл      "
        LOCATE 27, 17: PRINT "лл      "
        LOCATE 28, 17: PRINT "лллллллл"
        LOCATE 29, 17: PRINT "лл    лл"
        LOCATE 30, 17: PRINT "лл    лл"
        LOCATE 31, 17: PRINT "лллллллл"
    CASE 7:
        LOCATE 25, 17: PRINT "лллллллл"
        LOCATE 26, 17: PRINT "      лл"
        LOCATE 27, 17: PRINT "      лл"
        LOCATE 28, 17: PRINT "      лл"
        LOCATE 29, 17: PRINT "      лл"
        LOCATE 30, 17: PRINT "      лл"
        LOCATE 31, 17: PRINT "      лл"
    CASE 8:
        LOCATE 25, 17: PRINT "лллллллл"
        LOCATE 26, 17: PRINT "лл    лл"
        LOCATE 27, 17: PRINT "лл    лл"
        LOCATE 28, 17: PRINT "лллллллл"
        LOCATE 29, 17: PRINT "лл    лл"
        LOCATE 30, 17: PRINT "лл    лл"
        LOCATE 31, 17: PRINT "лллллллл"
    CASE 9:
        LOCATE 25, 17: PRINT "лллллллл"
        LOCATE 26, 17: PRINT "лл    лл"
        LOCATE 27, 17: PRINT "лл    лл"
        LOCATE 28, 17: PRINT "лллллллл"
        LOCATE 29, 17: PRINT "      лл"
        LOCATE 30, 17: PRINT "      лл"
        LOCATE 31, 17: PRINT "лллллллл"
END SELECT
3
IF C <> VAL(MID$(TIME$, 4, 1)) THEN C = VAL(MID$(TIME$, 4, 1)): GOTO 6 ELSE GOTO 5
6
SELECT CASE C
    CASE 0:
        LOCATE 25, 31: PRINT "лллллллл"
        LOCATE 26, 31: PRINT "лл    лл"
        LOCATE 27, 31: PRINT "лл    лл"
        LOCATE 28, 31: PRINT "лл    лл"
        LOCATE 29, 31: PRINT "лл    лл"
        LOCATE 30, 31: PRINT "лл    лл"
        LOCATE 31, 31: PRINT "лллллллл"
    CASE 1
        LOCATE 25, 31: PRINT "      лл"
        LOCATE 26, 31: PRINT "      лл"
        LOCATE 27, 31: PRINT "      лл"
        LOCATE 28, 31: PRINT "      лл"
        LOCATE 29, 31: PRINT "      лл"
        LOCATE 30, 31: PRINT "      лл"
        LOCATE 31, 31: PRINT "      лл"
    CASE 2:
        LOCATE 25, 31: PRINT "лллллллл"
        LOCATE 26, 31: PRINT "      лл"
        LOCATE 27, 31: PRINT "      лл"
        LOCATE 28, 31: PRINT "лллллллл"
        LOCATE 29, 31: PRINT "лл      "
        LOCATE 30, 31: PRINT "лл      "
        LOCATE 31, 31: PRINT "лллллллл"
    CASE 3:
        LOCATE 25, 31: PRINT "лллллллл"
        LOCATE 26, 31: PRINT "      лл"
        LOCATE 27, 31: PRINT "      лл"
        LOCATE 28, 31: PRINT "лллллллл"
        LOCATE 29, 31: PRINT "      лл"
        LOCATE 30, 31: PRINT "      лл"
        LOCATE 31, 31: PRINT "лллллллл"
    CASE 4:
        LOCATE 25, 31: PRINT "лл    лл"
        LOCATE 26, 31: PRINT "лл    лл"
        LOCATE 27, 31: PRINT "лл    лл"
        LOCATE 28, 31: PRINT "лллллллл"
        LOCATE 29, 31: PRINT "      лл"
        LOCATE 30, 31: PRINT "      лл"
        LOCATE 31, 31: PRINT "      лл"
    CASE 5:
        LOCATE 25, 31: PRINT "лллллллл"
        LOCATE 26, 31: PRINT "лл      "
        LOCATE 27, 31: PRINT "лл      "
        LOCATE 28, 31: PRINT "лллллллл"
        LOCATE 29, 31: PRINT "      лл"
        LOCATE 30, 31: PRINT "      лл"
        LOCATE 31, 31: PRINT "лллллллл"
    CASE 6:
        LOCATE 25, 31: PRINT "лллллллл"
        LOCATE 26, 31: PRINT "лл      "
        LOCATE 27, 31: PRINT "лл      "
        LOCATE 28, 31: PRINT "лллллллл"
        LOCATE 29, 31: PRINT "лл    лл"
        LOCATE 30, 31: PRINT "лл    лл"
        LOCATE 31, 31: PRINT "лллллллл"
    CASE 7:
        LOCATE 25, 31: PRINT "лллллллл"
        LOCATE 26, 31: PRINT "      лл"
        LOCATE 27, 31: PRINT "      лл"
        LOCATE 28, 31: PRINT "      лл"
        LOCATE 29, 31: PRINT "      лл"
        LOCATE 30, 31: PRINT "      лл"
        LOCATE 31, 31: PRINT "      лл"
    CASE 8:
        LOCATE 25, 31: PRINT "лллллллл"
        LOCATE 26, 31: PRINT "лл    лл"
        LOCATE 27, 31: PRINT "лл    лл"
        LOCATE 28, 31: PRINT "лллллллл"
        LOCATE 29, 31: PRINT "лл    лл"
        LOCATE 30, 31: PRINT "лл    лл"
        LOCATE 31, 31: PRINT "лллллллл"
    CASE 9:
        LOCATE 25, 31: PRINT "лллллллл"
        LOCATE 26, 31: PRINT "лл    лл"
        LOCATE 27, 31: PRINT "лл    лл"
        LOCATE 28, 31: PRINT "лллллллл"
        LOCATE 29, 31: PRINT "      лл"
        LOCATE 30, 31: PRINT "      лл"
        LOCATE 31, 31: PRINT "лллллллл"
END SELECT
5
IF D <> VAL(MID$(TIME$, 5, 1)) THEN D = VAL(MID$(TIME$, 5, 1)): GOTO 8 ELSE GOTO 7
8
SELECT CASE D
    CASE 0:
        LOCATE 25, 41: PRINT "лллллллл"
        LOCATE 26, 41: PRINT "лл    лл"
        LOCATE 27, 41: PRINT "лл    лл"
        LOCATE 28, 41: PRINT "лл    лл"
        LOCATE 29, 41: PRINT "лл    лл"
        LOCATE 30, 41: PRINT "лл    лл"
        LOCATE 31, 41: PRINT "лллллллл"
    CASE 1
        LOCATE 25, 41: PRINT "      лл"
        LOCATE 26, 41: PRINT "      лл"
        LOCATE 27, 41: PRINT "      лл"
        LOCATE 28, 41: PRINT "      лл"
        LOCATE 29, 41: PRINT "      лл"
        LOCATE 30, 41: PRINT "      лл"
        LOCATE 31, 41: PRINT "      лл"
    CASE 2:
        LOCATE 25, 41: PRINT "лллллллл"
        LOCATE 26, 41: PRINT "      лл"
        LOCATE 27, 41: PRINT "      лл"
        LOCATE 28, 41: PRINT "лллллллл"
        LOCATE 29, 41: PRINT "лл      "
        LOCATE 30, 41: PRINT "лл      "
        LOCATE 31, 41: PRINT "лллллллл"
    CASE 3:
        LOCATE 25, 41: PRINT "лллллллл"
        LOCATE 26, 41: PRINT "      лл"
        LOCATE 27, 41: PRINT "      лл"
        LOCATE 28, 41: PRINT "лллллллл"
        LOCATE 29, 41: PRINT "      лл"
        LOCATE 30, 41: PRINT "      лл"
        LOCATE 31, 41: PRINT "лллллллл"
    CASE 4:
        LOCATE 25, 41: PRINT "лл    лл"
        LOCATE 26, 41: PRINT "лл    лл"
        LOCATE 27, 41: PRINT "лл    лл"
        LOCATE 28, 41: PRINT "лллллллл"
        LOCATE 29, 41: PRINT "      лл"
        LOCATE 30, 41: PRINT "      лл"
        LOCATE 31, 41: PRINT "      лл"
    CASE 5:
        LOCATE 25, 41: PRINT "лллллллл"
        LOCATE 26, 41: PRINT "лл      "
        LOCATE 27, 41: PRINT "лл      "
        LOCATE 28, 41: PRINT "лллллллл"
        LOCATE 29, 41: PRINT "      лл"
        LOCATE 30, 41: PRINT "      лл"
        LOCATE 31, 41: PRINT "лллллллл"
    CASE 6:
        LOCATE 25, 41: PRINT "лллллллл"
        LOCATE 26, 41: PRINT "лл      "
        LOCATE 27, 41: PRINT "лл      "
        LOCATE 28, 41: PRINT "лллллллл"
        LOCATE 29, 41: PRINT "лл    лл"
        LOCATE 30, 41: PRINT "лл    лл"
        LOCATE 31, 41: PRINT "лллллллл"
    CASE 7:
        LOCATE 25, 41: PRINT "лллллллл"
        LOCATE 26, 41: PRINT "      лл"
        LOCATE 27, 41: PRINT "      лл"
        LOCATE 28, 41: PRINT "      лл"
        LOCATE 29, 41: PRINT "      лл"
        LOCATE 30, 41: PRINT "      лл"
        LOCATE 31, 41: PRINT "      лл"
    CASE 8:
        LOCATE 25, 41: PRINT "лллллллл"
        LOCATE 26, 41: PRINT "лл    лл"
        LOCATE 27, 41: PRINT "лл    лл"
        LOCATE 28, 41: PRINT "лллллллл"
        LOCATE 29, 41: PRINT "лл    лл"
        LOCATE 30, 41: PRINT "лл    лл"
        LOCATE 31, 41: PRINT "лллллллл"
    CASE 9:
        LOCATE 25, 41: PRINT "лллллллл"
        LOCATE 26, 41: PRINT "лл    лл"
        LOCATE 27, 41: PRINT "лл    лл"
        LOCATE 28, 41: PRINT "лллллллл"
        LOCATE 29, 41: PRINT "      лл"
        LOCATE 30, 41: PRINT "      лл"
        LOCATE 31, 41: PRINT "лллллллл"
END SELECT
7
IF E <> VAL(MID$(TIME$, 7, 1)) THEN E = VAL(MID$(TIME$, 7, 1)): GOTO 10 ELSE GOTO 9
10
SELECT CASE E
    CASE 0:
        LOCATE 25, 55: PRINT "лллллллл"
        LOCATE 26, 55: PRINT "лл    лл"
        LOCATE 27, 55: PRINT "лл    лл"
        LOCATE 28, 55: PRINT "лл    лл"
        LOCATE 29, 55: PRINT "лл    лл"
        LOCATE 30, 55: PRINT "лл    лл"
        LOCATE 31, 55: PRINT "лллллллл"
    CASE 1
        LOCATE 25, 55: PRINT "      лл"
        LOCATE 26, 55: PRINT "      лл"
        LOCATE 27, 55: PRINT "      лл"
        LOCATE 28, 55: PRINT "      лл"
        LOCATE 29, 55: PRINT "      лл"
        LOCATE 30, 55: PRINT "      лл"
        LOCATE 31, 55: PRINT "      лл"
    CASE 2:
        LOCATE 25, 55: PRINT "лллллллл"
        LOCATE 26, 55: PRINT "      лл"
        LOCATE 27, 55: PRINT "      лл"
        LOCATE 28, 55: PRINT "лллллллл"
        LOCATE 29, 55: PRINT "лл      "
        LOCATE 30, 55: PRINT "лл      "
        LOCATE 31, 55: PRINT "лллллллл"
    CASE 3:
        LOCATE 25, 55: PRINT "лллллллл"
        LOCATE 26, 55: PRINT "      лл"
        LOCATE 27, 55: PRINT "      лл"
        LOCATE 28, 55: PRINT "лллллллл"
        LOCATE 29, 55: PRINT "      лл"
        LOCATE 30, 55: PRINT "      лл"
        LOCATE 31, 55: PRINT "лллллллл"
    CASE 4:
        LOCATE 25, 55: PRINT "лл    лл"
        LOCATE 26, 55: PRINT "лл    лл"
        LOCATE 27, 55: PRINT "лл    лл"
        LOCATE 28, 55: PRINT "лллллллл"
        LOCATE 29, 55: PRINT "      лл"
        LOCATE 30, 55: PRINT "      лл"
        LOCATE 31, 55: PRINT "      лл"
    CASE 5:
        LOCATE 25, 55: PRINT "лллллллл"
        LOCATE 26, 55: PRINT "лл      "
        LOCATE 27, 55: PRINT "лл      "
        LOCATE 28, 55: PRINT "лллллллл"
        LOCATE 29, 55: PRINT "      лл"
        LOCATE 30, 55: PRINT "      лл"
        LOCATE 31, 55: PRINT "лллллллл"
    CASE 6:
        LOCATE 25, 55: PRINT "лллллллл"
        LOCATE 26, 55: PRINT "лл      "
        LOCATE 27, 55: PRINT "лл      "
        LOCATE 28, 55: PRINT "лллллллл"
        LOCATE 29, 55: PRINT "лл    лл"
        LOCATE 30, 55: PRINT "лл    лл"
        LOCATE 31, 55: PRINT "лллллллл"
    CASE 7:
        LOCATE 25, 55: PRINT "лллллллл"
        LOCATE 26, 55: PRINT "      лл"
        LOCATE 27, 55: PRINT "      лл"
        LOCATE 28, 55: PRINT "      лл"
        LOCATE 29, 55: PRINT "      лл"
        LOCATE 30, 55: PRINT "      лл"
        LOCATE 31, 55: PRINT "      лл"
    CASE 8:
        LOCATE 25, 55: PRINT "лллллллл"
        LOCATE 26, 55: PRINT "лл    лл"
        LOCATE 27, 55: PRINT "лл    лл"
        LOCATE 28, 55: PRINT "лллллллл"
        LOCATE 29, 55: PRINT "лл    лл"
        LOCATE 30, 55: PRINT "лл    лл"
        LOCATE 31, 55: PRINT "лллллллл"
    CASE 9:
        LOCATE 25, 55: PRINT "лллллллл"
        LOCATE 26, 55: PRINT "лл    лл"
        LOCATE 27, 55: PRINT "лл    лл"
        LOCATE 28, 55: PRINT "лллллллл"
        LOCATE 29, 55: PRINT "      лл"
        LOCATE 30, 55: PRINT "      лл"
        LOCATE 31, 55: PRINT "лллллллл"
END SELECT
9
IF F <> VAL(MID$(TIME$, 8, 1)) THEN F = VAL(MID$(TIME$, 8, 1)): GOTO 12 ELSE GOTO 1.1
12
SELECT CASE F
    CASE 0:
        LOCATE 25, 65: PRINT "лллллллл"
        LOCATE 26, 65: PRINT "лл    лл"
        LOCATE 27, 65: PRINT "лл    лл"
        LOCATE 28, 65: PRINT "лл    лл"
        LOCATE 29, 65: PRINT "лл    лл"
        LOCATE 30, 65: PRINT "лл    лл"
        LOCATE 31, 65: PRINT "лллллллл"
    CASE 1:
        LOCATE 25, 65: PRINT "      лл"
        LOCATE 26, 65: PRINT "      лл"
        LOCATE 27, 65: PRINT "      лл"
        LOCATE 28, 65: PRINT "      лл"
        LOCATE 29, 65: PRINT "      лл"
        LOCATE 30, 65: PRINT "      лл"
        LOCATE 31, 65: PRINT "      лл"
    CASE 2:
        LOCATE 25, 65: PRINT "лллллллл"
        LOCATE 26, 65: PRINT "      лл"
        LOCATE 27, 65: PRINT "      лл"
        LOCATE 28, 65: PRINT "лллллллл"
        LOCATE 29, 65: PRINT "лл      "
        LOCATE 30, 65: PRINT "лл      "
        LOCATE 31, 65: PRINT "лллллллл"
    CASE 3:
        LOCATE 25, 65: PRINT "лллллллл"
        LOCATE 26, 65: PRINT "      лл"
        LOCATE 27, 65: PRINT "      лл"
        LOCATE 28, 65: PRINT "лллллллл"
        LOCATE 29, 65: PRINT "      лл"
        LOCATE 30, 65: PRINT "      лл"
        LOCATE 31, 65: PRINT "лллллллл"
    CASE 4:
        LOCATE 25, 65: PRINT "лл    лл"
        LOCATE 26, 65: PRINT "лл    лл"
        LOCATE 27, 65: PRINT "лл    лл"
        LOCATE 28, 65: PRINT "лллллллл"
        LOCATE 29, 65: PRINT "      лл"
        LOCATE 30, 65: PRINT "      лл"
        LOCATE 31, 65: PRINT "      лл"
    CASE 5:
        LOCATE 25, 65: PRINT "лллллллл"
        LOCATE 26, 65: PRINT "лл      "
        LOCATE 27, 65: PRINT "лл      "
        LOCATE 28, 65: PRINT "лллллллл"
        LOCATE 29, 65: PRINT "      лл"
        LOCATE 30, 65: PRINT "      лл"
        LOCATE 31, 65: PRINT "лллллллл"
    CASE 6:
        LOCATE 25, 65: PRINT "лллллллл"
        LOCATE 26, 65: PRINT "лл      "
        LOCATE 27, 65: PRINT "лл      "
        LOCATE 28, 65: PRINT "лллллллл"
        LOCATE 29, 65: PRINT "лл    лл"
        LOCATE 30, 65: PRINT "лл    лл"
        LOCATE 31, 65: PRINT "лллллллл"
    CASE 7:
        LOCATE 25, 65: PRINT "лллллллл"
        LOCATE 26, 65: PRINT "      лл"
        LOCATE 27, 65: PRINT "      лл"
        LOCATE 28, 65: PRINT "      лл"
        LOCATE 29, 65: PRINT "      лл"
        LOCATE 30, 65: PRINT "      лл"
        LOCATE 31, 65: PRINT "      лл"
    CASE 8:
        LOCATE 25, 65: PRINT "лллллллл"
        LOCATE 26, 65: PRINT "лл    лл"
        LOCATE 27, 65: PRINT "лл    лл"
        LOCATE 28, 65: PRINT "лллллллл"
        LOCATE 29, 65: PRINT "лл    лл"
        LOCATE 30, 65: PRINT "лл    лл"
        LOCATE 31, 65: PRINT "лллллллл"
    CASE 9:
        LOCATE 25, 65: PRINT "лллллллл"
        LOCATE 26, 65: PRINT "лл    лл"
        LOCATE 27, 65: PRINT "лл    лл"
        LOCATE 28, 65: PRINT "лллллллл"
        LOCATE 29, 65: PRINT "      лл"
        LOCATE 30, 65: PRINT "      лл"
        LOCATE 31, 65: PRINT "лллллллл"
END SELECT
GOTO 1.1

DEFINT A-Z
SUB DATUM
    REM KLPRINT TIME$, 18, 5, 0, 4
    IF MID$(TIME$, 1, 2) <> "00" THEN GOTO 3.1
    IF MID$(TIME$, 4, 2) <> "00" THEN GOTO 3.1
    IF MID$(TIME$, 7, 2) <> "00" THEN GOTO 3.1
    KLPRINT MID$(DATE$, 4, 2) + ".", 18, 476, 0, 4
    KLPRINT MID$(DATE$, 1, 2) + ".", 18, 494, 0, 4
    KLPRINT MID$(DATE$, 7, 4), 18, 511, 0, 4
3.1 END SUB

SUB FELT (X, Y, XX, YY, FARBEN)
    REM LINE (X, Y)-(X, YY), 15
    REM LINE (X, Y)-(XX, Y), 15
    REM LINE (X, YY)-(XX, YY), 8
    REM LINE (XX, Y)-(XX, YY), 8
    VIEW (X + 2, Y + 2)-(XX - 2, YY - 2), FARBEN, FARBEN
    VIEW (0, 0)-(639, 479)
END SUB

SUB KLPRINT (B$, B, K, G, I)
    C = LEN(B$)
    B$ = UCASE$(B$)
    D = K - 6
    E = B - 14
    IF G <> 10 THEN E = E - G
    7.1: F = F + 1
    H = F * 6
    IF F = C + 1 THEN GOTO 198
    IF G = 10 THEN
        VIEW (D + H, E)-(D + H + 4, E + 6), 7, 7
        VIEW (0, 0)-(639, 479)
    END IF
    IF MID$(B$, F, 1) = " " THEN
        GOTO 123
    ELSEIF MID$(B$, F, 1) = "A" THEN
        PSET (D + H + 1, E + 0), I: PSET (D + H + 2, E + 0), I: PSET (D + H + 3, E + 0), I
        PSET (D + H, E + 1), I: PSET (D + H + 4, E + 1), I
        PSET (D + H, E + 2), I: PSET (D + H + 4, E + 2), I
        PSET (D + H, E + 3), I: PSET (D + H + 1, E + 3), I: PSET (D + H + 2, E + 3), I: PSET (D + H + 3, E + 3), I: PSET (D + H + 4, E + 3), I
        PSET (D + H, E + 4), I: PSET (D + H + 4, E + 4), I
        PSET (D + H, E + 5), I: PSET (D + H + 4, E + 5), I
        PSET (D + H, E + 6), I: PSET (D + H + 4, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "B" THEN
        PSET (D + H, E + 0), I: PSET (D + H + 1, E + 0), I: PSET (D + H + 2, E + 0), I: PSET (D + H + 3, E + 0), I
        PSET (D + H, E + 1), I: PSET (D + H + 4, E + 1), I
        PSET (D + H, E + 2), I: PSET (D + H + 4, E + 2), I
        PSET (D + H, E + 3), I: PSET (D + H + 1, E + 3), I: PSET (D + H + 2, E + 3), I: PSET (D + H + 3, E + 3), I
        PSET (D + H, E + 4), I: PSET (D + H + 4, E + 4), I
        PSET (D + H, E + 5), I: PSET (D + H + 4, E + 5), I
        PSET (D + H, E + 6), I: PSET (D + H + 1, E + 6), I: PSET (D + H + 2, E + 6), I: PSET (D + H + 3, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "C" THEN
        PSET (D + H + 1, E + 0), I: PSET (D + H + 2, E + 0), I: PSET (D + H + 3, E + 0), I
        PSET (D + H, E + 1), I: PSET (D + H + 4, E + 1), I
        PSET (D + H, E + 2), I
        PSET (D + H, E + 3), I
        PSET (D + H, E + 4), I
        PSET (D + H, E + 5), I: PSET (D + H + 4, E + 5), I
        PSET (D + H + 1, E + 6), I: PSET (D + H + 2, E + 6), I: PSET (D + H + 3, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "D" THEN
        PSET (D + H, E + 0), I: PSET (D + H + 1, E + 0), I: PSET (D + H + 2, E + 0), I: PSET (D + H + 3, E + 0), I
        PSET (D + H, E + 1), I: PSET (D + H + 4, E + 1), I
        PSET (D + H, E + 2), I: PSET (D + H + 4, E + 2), I
        PSET (D + H, E + 3), I: PSET (D + H + 4, E + 3), I
        PSET (D + H, E + 4), I: PSET (D + H + 4, E + 4), I
        PSET (D + H, E + 5), I: PSET (D + H + 4, E + 5), I
        PSET (D + H, E + 6), I: PSET (D + H + 1, E + 6), I: PSET (D + H + 2, E + 6), I: PSET (D + H + 3, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "E" THEN
        PSET (D + H, E + 0), I: PSET (D + H + 1, E + 0), I: PSET (D + H + 2, E + 0), I: PSET (D + H + 3, E + 0), I: PSET (D + H + 4, E + 0), I
        PSET (D + H, E + 1), I
        PSET (D + H, E + 2), I
        PSET (D + H, E + 3), I: PSET (D + H + 1, E + 3), I: PSET (D + H + 2, E + 3), I: PSET (D + H + 3, E + 3), I
        PSET (D + H, E + 4), I
        PSET (D + H, E + 5), I
        PSET (D + H, E + 6), I: PSET (D + H + 1, E + 6), I: PSET (D + H + 2, E + 6), I: PSET (D + H + 3, E + 6), I: PSET (D + H + 4, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "F" THEN
        PSET (D + H, E + 0), I: PSET (D + H + 1, E + 0), I: PSET (D + H + 2, E + 0), I: PSET (D + H + 3, E + 0), I: PSET (D + H + 4, E + 0), I
        PSET (D + H, E + 1), I
        PSET (D + H, E + 2), I
        PSET (D + H, E + 3), I: PSET (D + H + 1, E + 3), I: PSET (D + H + 2, E + 3), I
        PSET (D + H, E + 4), I
        PSET (D + H, E + 5), I
        PSET (D + H, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "G" THEN
        PSET (D + H + 1, E + 0), I: PSET (D + H + 2, E + 0), I: PSET (D + H + 3, E + 0), I
        PSET (D + H, E + 1), I: PSET (D + H + 4, E + 1), I
        PSET (D + H, E + 2), I
        PSET (D + H, E + 3), I: PSET (D + H + 3, E + 3), I: PSET (D + H + 4, E + 3), I
        PSET (D + H, E + 4), I: PSET (D + H + 4, E + 4), I
        PSET (D + H, E + 5), I: PSET (D + H + 4, E + 5), I
        PSET (D + H + 1, E + 6), I: PSET (D + H + 2, E + 6), I: PSET (D + H + 3, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "H" THEN
        PSET (D + H, E + 0), I: PSET (D + H + 4, E + 0), I
        PSET (D + H, E + 1), I: PSET (D + H + 4, E + 1), I
        PSET (D + H, E + 2), I: PSET (D + H + 4, E + 2), I
        PSET (D + H, E + 3), I: PSET (D + H + 1, E + 3), I: PSET (D + H + 2, E + 3), I: PSET (D + H + 3, E + 3), I: PSET (D + H + 4, E + 3), I
        PSET (D + H, E + 4), I: PSET (D + H + 4, E + 4), I
        PSET (D + H, E + 5), I: PSET (D + H + 4, E + 5), I
        PSET (D + H, E + 6), I: PSET (D + H + 4, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "I" THEN
        PSET (D + H + 1, E + 0), I: PSET (D + H + 2, E + 0), I: PSET (D + H + 3, E + 0), I
        PSET (D + H + 2, E + 1), I
        PSET (D + H + 2, E + 2), I
        PSET (D + H + 2, E + 3), I
        PSET (D + H + 2, E + 4), I
        PSET (D + H + 2, E + 5), I
        PSET (D + H + 1, E + 6), I: PSET (D + H + 2, E + 6), I: PSET (D + H + 3, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "J" THEN
        PSET (D + H + 2, E + 0), I: PSET (D + H + 3, E + 0), I: PSET (D + H + 4, E + 0), I
        PSET (D + H + 3, E + 1), I
        PSET (D + H + 3, E + 2), I
        PSET (D + H + 3, E + 3), I
        PSET (D + H + 3, E + 4), I
        PSET (D + H, E + 5), I: PSET (D + H + 3, E + 5), I
        PSET (D + H + 1, E + 6), I: PSET (D + H + 2, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "K" THEN
        PSET (D + H, E + 0), I: PSET (D + H + 4, E + 0), I
        PSET (D + H, E + 1), I: PSET (D + H + 3, E + 1), I
        PSET (D + H, E + 2), I: PSET (D + H + 2, E + 2), I
        PSET (D + H, E + 3), I: PSET (D + H + 1, E + 3), I
        PSET (D + H, E + 4), I: PSET (D + H + 2, E + 4), I
        PSET (D + H, E + 5), I: PSET (D + H + 3, E + 5), I
        PSET (D + H, E + 6), I: PSET (D + H + 4, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "L" THEN
        PSET (D + H, E + 0), I
        PSET (D + H, E + 1), I
        PSET (D + H, E + 2), I
        PSET (D + H, E + 3), I
        PSET (D + H, E + 4), I
        PSET (D + H, E + 5), I
        PSET (D + H, E + 6), I: PSET (D + H + 1, E + 6), I: PSET (D + H + 2, E + 6), I: PSET (D + H + 3, E + 6), I: PSET (D + H + 4, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "M" THEN
        PSET (D + H, E + 0), I: PSET (D + H + 4, E + 0), I
        PSET (D + H, E + 1), I: PSET (D + H + 1, E + 1), I: PSET (D + H + 3, E + 1), I: PSET (D + H + 4, E + 1), I
        PSET (D + H, E + 2), I: PSET (D + H + 2, E + 2), I: PSET (D + H + 4, E + 2), I
        PSET (D + H, E + 3), I: PSET (D + H + 2, E + 3), I: PSET (D + H + 4, E + 3), I
        PSET (D + H, E + 4), I: PSET (D + H + 4, E + 4), I
        PSET (D + H, E + 5), I: PSET (D + H + 4, E + 5), I
        PSET (D + H, E + 6), I: PSET (D + H + 4, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "N" THEN
        PSET (D + H, E + 0), I: PSET (D + H + 4, E + 0), I
        PSET (D + H, E + 1), I: PSET (D + H + 1, E + 1), I: PSET (D + H + 4, E + 1), I
        PSET (D + H, E + 2), I: PSET (D + H + 2, E + 2), I: PSET (D + H + 4, E + 2), I
        PSET (D + H, E + 3), I: PSET (D + H + 2, E + 3), I: PSET (D + H + 4, E + 3), I
        PSET (D + H, E + 4), I: PSET (D + H + 2, E + 4), I: PSET (D + H + 4, E + 4), I
        PSET (D + H, E + 5), I: PSET (D + H + 3, E + 5), I: PSET (D + H + 4, E + 5), I
        PSET (D + H, E + 6), I: PSET (D + H + 4, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "O" THEN
        PSET (D + H + 1, E + 0), I: PSET (D + H + 2, E + 0), I: PSET (D + H + 3, E + 0), I
        PSET (D + H, E + 1), I: PSET (D + H + 4, E + 1), I
        PSET (D + H, E + 2), I: PSET (D + H + 4, E + 2), I
        PSET (D + H, E + 3), I: PSET (D + H + 4, E + 3), I
        PSET (D + H, E + 4), I: PSET (D + H + 4, E + 4), I
        PSET (D + H, E + 5), I: PSET (D + H + 4, E + 5), I
        PSET (D + H + 1, E + 6), I: PSET (D + H + 2, E + 6), I: PSET (D + H + 3, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "P" THEN
        PSET (D + H, E + 0), I: PSET (D + H + 1, E + 0), I: PSET (D + H + 2, E + 0), I: PSET (D + H + 3, E + 0), I
        PSET (D + H, E + 1), I: PSET (D + H + 4, E + 1), I
        PSET (D + H, E + 2), I: PSET (D + H + 4, E + 2), I
        PSET (D + H, E + 3), I: PSET (D + H + 1, E + 3), I: PSET (D + H + 2, E + 3), I: PSET (D + H + 3, E + 3), I
        PSET (D + H, E + 4), I
        PSET (D + H, E + 5), I
        PSET (D + H, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "Q" THEN
        PSET (D + H + 1, E + 0), I: PSET (D + H + 2, E + 0), I: PSET (D + H + 3, E + 0), I
        PSET (D + H, E + 1), I: PSET (D + H + 4, E + 1), I
        PSET (D + H, E + 2), I: PSET (D + H + 4, E + 2), I
        PSET (D + H, E + 3), I: PSET (D + H + 4, E + 3), I
        PSET (D + H, E + 4), I: PSET (D + H + 2, E + 4), I: PSET (D + H + 4, E + 4), I
        PSET (D + H, E + 5), I: PSET (D + H + 3, E + 5), I: PSET (D + H + 4, E + 5), I
        PSET (D + H + 1, E + 6), I: PSET (D + H + 2, E + 6), I: PSET (D + H + 3, E + 6), I: PSET (D + H + 4, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "R" THEN
        PSET (D + H, E + 0), I: PSET (D + H + 1, E + 0), I: PSET (D + H + 2, E + 0), I: PSET (D + H + 3, E + 0), I
        PSET (D + H, E + 1), I: PSET (D + H + 4, E + 1), I
        PSET (D + H, E + 2), I: PSET (D + H + 4, E + 2), I
        PSET (D + H, E + 3), I: PSET (D + H + 1, E + 3), I: PSET (D + H + 2, E + 3), I: PSET (D + H + 3, E + 3), I
        PSET (D + H, E + 4), I: PSET (D + H + 2, E + 4), I
        PSET (D + H, E + 5), I: PSET (D + H + 3, E + 5), I
        PSET (D + H, E + 6), I: PSET (D + H + 4, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "S" THEN
        PSET (D + H + 1, E + 0), I: PSET (D + H + 2, E + 0), I: PSET (D + H + 3, E + 0), I: PSET (D + H + 4, E + 0), I
        PSET (D + H, E + 1), I
        PSET (D + H, E + 2), I
        PSET (D + H + 1, E + 3), I: PSET (D + H + 2, E + 3), I: PSET (D + H + 3, E + 3), I
        PSET (D + H + 4, E + 4), I
        PSET (D + H + 4, E + 5), I
        PSET (D + H, E + 6), I: PSET (D + H + 1, E + 6), I: PSET (D + H + 2, E + 6), I: PSET (D + H + 3, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "T" THEN
        PSET (D + H, E + 0), I: PSET (D + H + 1, E + 0), I: PSET (D + H + 2, E + 0), I: PSET (D + H + 3, E + 0), I: PSET (D + H + 4, E + 0), I
        PSET (D + H + 2, E + 1), I
        PSET (D + H + 2, E + 2), I
        PSET (D + H + 2, E + 3), I
        PSET (D + H + 2, E + 4), I
        PSET (D + H + 2, E + 5), I
        PSET (D + H + 2, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "U" THEN
        PSET (D + H, E + 0), I: PSET (D + H + 4, E + 0), I
        PSET (D + H, E + 1), I: PSET (D + H + 4, E + 1), I
        PSET (D + H, E + 2), I: PSET (D + H + 4, E + 2), I
        PSET (D + H, E + 3), I: PSET (D + H + 4, E + 3), I
        PSET (D + H, E + 4), I: PSET (D + H + 4, E + 4), I
        PSET (D + H, E + 5), I: PSET (D + H + 4, E + 5), I
        PSET (D + H + 1, E + 6), I: PSET (D + H + 2, E + 6), I: PSET (D + H + 3, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "V" THEN
        PSET (D + H, E + 0), I: PSET (D + H + 4, E + 0), I
        PSET (D + H, E + 1), I: PSET (D + H + 4, E + 1), I
        PSET (D + H, E + 2), I: PSET (D + H + 4, E + 2), I
        PSET (D + H + 1, E + 3), I: PSET (D + H + 3, E + 3), I
        PSET (D + H + 1, E + 4), I: PSET (D + H + 3, E + 4), I
        PSET (D + H + 2, E + 5), I
        PSET (D + H + 2, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "W" THEN
        PSET (D + H, E + 0), I: PSET (D + H + 4, E + 0), I
        PSET (D + H, E + 1), I: PSET (D + H + 4, E + 1), I
        PSET (D + H, E + 2), I: PSET (D + H + 4, E + 2), I
        PSET (D + H, E + 3), I: PSET (D + H + 2, E + 3), I: PSET (D + H + 4, E + 3), I
        PSET (D + H, E + 4), I: PSET (D + H + 2, E + 4), I: PSET (D + H + 4, E + 4), I
        PSET (D + H, E + 5), I: PSET (D + H + 1, E + 5), I: PSET (D + H + 3, E + 5), I: PSET (D + H + 4, E + 5), I
        PSET (D + H, E + 6), I: PSET (D + H + 4, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "X" THEN
        PSET (D + H, E + 0), I: PSET (D + H + 4, E + 0), I
        PSET (D + H, E + 1), I: PSET (D + H + 4, E + 1), I
        PSET (D + H + 1, E + 2), I: PSET (D + H + 3, E + 2), I
        PSET (D + H + 2, E + 3), I
        PSET (D + H + 1, E + 4), I: PSET (D + H + 3, E + 4), I
        PSET (D + H, E + 5), I: PSET (D + H + 4, E + 5), I
        PSET (D + H, E + 6), I: PSET (D + H + 4, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "Y" THEN
        PSET (D + H, E + 0), I: PSET (D + H + 4, E + 0), I
        PSET (D + H, E + 1), I: PSET (D + H + 4, E + 1), I
        PSET (D + H, E + 2), I: PSET (D + H + 4, E + 2), I
        PSET (D + H + 1, E + 3), I: PSET (D + H + 3, E + 3), I
        PSET (D + H + 2, E + 4), I
        PSET (D + H + 2, E + 5), I
        PSET (D + H, E + 6), I: PSET (D + H + 1, E + 6), I: PSET (D + H + 2, E + 6), I: PSET (D + H + 3, E + 6), I: PSET (D + H + 4, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "Z" THEN
        PSET (D + H, E + 0), I: PSET (D + H + 1, E + 0), I: PSET (D + H + 2, E + 0), I: PSET (D + H + 3, E + 0), I: PSET (D + H + 4, E + 0), I
        PSET (D + H, E + 1), I: PSET (D + H + 4, E + 1), I
        PSET (D + H + 3, E + 2), I
        PSET (D + H + 2, E + 3), I
        PSET (D + H + 1, E + 4), I
        PSET (D + H, E + 5), I: PSET (D + H + 4, E + 5), I
        PSET (D + H, E + 6), I: PSET (D + H + 1, E + 6), I: PSET (D + H + 2, E + 6), I: PSET (D + H + 3, E + 6), I: PSET (D + H + 4, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "." THEN
        PSET (D + H + 1, E + 5), I: PSET (D + H + 2, E + 5), I
        PSET (D + H + 1, E + 6), I: PSET (D + H + 2, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "," THEN
        PSET (D + H + 1, E + 4), I: PSET (D + H + 2, E + 4), I
        PSET (D + H + 1, E + 5), I: PSET (D + H + 2, E + 5), I
        PSET (D + H, E + 6), I: PSET (D + H + 1, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "&" THEN
        PSET (D + H + 1, E + 0), I: PSET (D + H + 2, E + 0), I
        PSET (D + H, E + 1), I: PSET (D + H + 3, E + 1), I
        PSET (D + H, E + 2), I: PSET (D + H + 2, E + 2), I
        PSET (D + H + 1, E + 3), I
        PSET (D + H, E + 4), I: PSET (D + H + 2, E + 4), I: PSET (D + H + 4, E + 4), I
        PSET (D + H, E + 5), I: PSET (D + H + 3, E + 5), I
        PSET (D + H + 1, E + 6), I: PSET (D + H + 2, E + 6), I: PSET (D + H + 4, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "(" THEN
        PSET (D + H + 2, E + 0), I: PSET (D + H + 3, E + 0), I
        PSET (D + H + 1, E + 1), I
        PSET (D + H, E + 2), I
        PSET (D + H, E + 3), I
        PSET (D + H, E + 4), I
        PSET (D + H + 1, E + 5), I
        PSET (D + H + 2, E + 6), I: PSET (D + H + 3, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = ")" THEN
        PSET (D + H + 1, E + 0), I: PSET (D + H + 2, E + 0), I
        PSET (D + H + 3, E + 1), I
        PSET (D + H + 4, E + 2), I
        PSET (D + H + 4, E + 3), I
        PSET (D + H + 4, E + 4), I
        PSET (D + H + 3, E + 5), I
        PSET (D + H + 1, E + 6), I: PSET (D + H + 2, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "1" THEN
        PSET (D + H + 3, E + 0), I
        PSET (D + H + 2, E + 1), I: PSET (D + H + 3, E + 1), I
        PSET (D + H + 1, E + 2), I: PSET (D + H + 3, E + 2), I
        PSET (D + H + 3, E + 3), I
        PSET (D + H + 3, E + 4), I
        PSET (D + H + 3, E + 5), I
        PSET (D + H + 1, E + 6), I: PSET (D + H + 2, E + 6), I: PSET (D + H + 3, E + 6), I: PSET (D + H + 4, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "2" THEN
        PSET (D + H + 1, E + 0), I: PSET (D + H + 2, E + 0), I: PSET (D + H + 3, E + 0), I
        PSET (D + H, E + 1), I: PSET (D + H + 4, E + 1), I
        PSET (D + H, E + 2), I: PSET (D + H + 4, E + 2), I
        PSET (D + H + 3, E + 3), I
        PSET (D + H + 2, E + 4), I
        PSET (D + H + 1, E + 5), I
        PSET (D + H, E + 6), I: PSET (D + H + 1, E + 6), I: PSET (D + H + 2, E + 6), I: PSET (D + H + 3, E + 6), I: PSET (D + H + 4, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "3" THEN
        PSET (D + H + 1, E + 0), I: PSET (D + H + 2, E + 0), I: PSET (D + H + 3, E + 0), I
        PSET (D + H, E + 1), I: PSET (D + H + 4, E + 1), I
        PSET (D + H + 4, E + 2), I
        PSET (D + H + 2, E + 3), I: PSET (D + H + 3, E + 3), I
        PSET (D + H + 4, E + 4), I
        PSET (D + H, E + 5), I: PSET (D + H + 4, E + 5), I
        PSET (D + H + 1, E + 6), I: PSET (D + H + 2, E + 6), I: PSET (D + H + 3, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "4" THEN
        PSET (D + H + 3, E + 0), I
        PSET (D + H + 2, E + 1), I: PSET (D + H + 3, E + 1), I
        PSET (D + H + 1, E + 2), I: PSET (D + H + 3, E + 2), I
        PSET (D + H, E + 3), I: PSET (D + H + 1, E + 3), I: PSET (D + H + 2, E + 3), I: PSET (D + H + 3, E + 3), I: PSET (D + H + 4, E + 3), I
        PSET (D + H + 3, E + 4), I
        PSET (D + H + 3, E + 5), I
        PSET (D + H + 3, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "5" THEN
        PSET (D + H, E + 0), I: PSET (D + H + 1, E + 0), I: PSET (D + H + 2, E + 0), I: PSET (D + H + 3, E + 0), I: PSET (D + H + 4, E + 0), I
        PSET (D + H, E + 1), I
        PSET (D + H, E + 2), I
        PSET (D + H, E + 3), I: PSET (D + H + 1, E + 3), I: PSET (D + H + 2, E + 3), I: PSET (D + H + 3, E + 3), I
        PSET (D + H + 4, E + 4), I
        PSET (D + H + 4, E + 5), I
        PSET (D + H, E + 6), I: PSET (D + H + 1, E + 6), I: PSET (D + H + 2, E + 6), I: PSET (D + H + 3, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "6" THEN
        PSET (D + H + 1, E + 0), I: PSET (D + H + 2, E + 0), I: PSET (D + H + 3, E + 0), I
        PSET (D + H, E + 1), I: PSET (D + H + 4, E + 1), I
        PSET (D + H, E + 2), I
        PSET (D + H, E + 3), I: PSET (D + H + 1, E + 3), I: PSET (D + H + 2, E + 3), I: PSET (D + H + 3, E + 3), I
        PSET (D + H, E + 4), I: PSET (D + H + 4, E + 4), I
        PSET (D + H, E + 5), I: PSET (D + H + 4, E + 5), I
        PSET (D + H + 1, E + 6), I: PSET (D + H + 2, E + 6), I: PSET (D + H + 3, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "7" THEN
        PSET (D + H, E + 0), I: PSET (D + H + 1, E + 0), I: PSET (D + H + 2, E + 0), I: PSET (D + H + 3, E + 0), I: PSET (D + H + 4, E + 0), I
        PSET (D + H, E + 1), I: PSET (D + H + 4, E + 1), I
        PSET (D + H + 3, E + 2), I
        PSET (D + H + 3, E + 3), I
        PSET (D + H + 2, E + 4), I
        PSET (D + H + 2, E + 5), I
        PSET (D + H + 2, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "8" THEN
        PSET (D + H + 1, E + 0), I: PSET (D + H + 2, E + 0), I: PSET (D + H + 3, E + 0), I
        PSET (D + H, E + 1), I: PSET (D + H + 4, E + 1), I
        PSET (D + H, E + 2), I: PSET (D + H + 4, E + 2), I
        PSET (D + H + 1, E + 3), I: PSET (D + H + 2, E + 3), I: PSET (D + H + 3, E + 3), I
        PSET (D + H, E + 4), I: PSET (D + H + 4, E + 4), I
        PSET (D + H, E + 5), I: PSET (D + H + 4, E + 5), I
        PSET (D + H + 1, E + 6), I: PSET (D + H + 2, E + 6), I: PSET (D + H + 3, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "9" THEN
        PSET (D + H + 1, E + 0), I: PSET (D + H + 2, E + 0), I: PSET (D + H + 3, E + 0), I
        PSET (D + H, E + 1), I: PSET (D + H + 4, E + 1), I
        PSET (D + H, E + 2), I: PSET (D + H + 4, E + 2), I
        PSET (D + H + 1, E + 3), I: PSET (D + H + 2, E + 3), I: PSET (D + H + 3, E + 3), I: PSET (D + H + 4, E + 3), I
        PSET (D + H + 4, E + 4), I
        PSET (D + H, E + 5), I: PSET (D + H + 4, E + 5), I
        PSET (D + H + 1, E + 6), I: PSET (D + H + 2, E + 6), I: PSET (D + H + 3, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "0" THEN
        PSET (D + H + 1, E + 0), I: PSET (D + H + 2, E + 0), I: PSET (D + H + 3, E + 0), I
        PSET (D + H, E + 1), I: PSET (D + H + 4, E + 1), I
        PSET (D + H, E + 2), I: PSET (D + H + 3, E + 2), I: PSET (D + H + 4, E + 2), I
        PSET (D + H, E + 3), I: PSET (D + H + 2, E + 3), I: PSET (D + H + 4, E + 3), I
        PSET (D + H, E + 4), I: PSET (D + H + 1, E + 4), I: PSET (D + H + 4, E + 4), I
        PSET (D + H, E + 5), I: PSET (D + H + 4, E + 5), I
        PSET (D + H + 1, E + 6), I: PSET (D + H + 2, E + 6), I: PSET (D + H + 3, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = ":" THEN
        PSET (D + H + 1, E + 1), I: PSET (D + H + 2, E + 1), I
        PSET (D + H + 1, E + 2), I: PSET (D + H + 2, E + 2), I
        PSET (D + H + 1, E + 4), I: PSET (D + H + 2, E + 4), I
        PSET (D + H + 1, E + 5), I: PSET (D + H + 2, E + 5), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "-" THEN
        PSET (D + H + 1, E + 3), I: PSET (D + H + 2, E + 3), I
        PSET (D + H + 3, E + 3), I: PSET (D + H + 4, E + 3), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "$" THEN
        PSET (D + H + 2, E + 0), I
        PSET (D + H + 1, E + 1), I: PSET (D + H + 2, E + 1), I: PSET (D + H + 3, E + 1), I: PSET (D + H + 4, E + 1), I
        PSET (D + H, E + 2), I: PSET (D + H + 2, E + 2), I
        PSET (D + H + 1, E + 3), I: PSET (D + H + 2, E + 3), I: PSET (D + H + 3, E + 3), I
        PSET (D + H + 2, E + 4), I: PSET (D + H + 4, E + 4), I
        PSET (D + H, E + 5), I: PSET (D + H + 1, E + 5), I: PSET (D + H + 2, E + 5), I: PSET (D + H + 3, E + 5), I
        PSET (D + H + 2, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "*" THEN
        PSET (D + H, E + 1), I: PSET (D + H + 2, E + 1), I: PSET (D + H + 4, E + 1), I
        PSET (D + H + 1, E + 2), I: PSET (D + H + 2, E + 2), I: PSET (D + H + 3, E + 2), I
        PSET (D + H, E + 3), I: PSET (D + H + 1, E + 3), I: PSET (D + H + 2, E + 3), I: PSET (D + H + 3, E + 3), I: PSET (D + H + 4, E + 3), I
        PSET (D + H + 1, E + 4), I: PSET (D + H + 2, E + 4), I: PSET (D + H + 3, E + 4), I
        PSET (D + H, E + 5), I: PSET (D + H + 2, E + 5), I: PSET (D + H + 4, E + 5), I
    ELSEIF MID$(B$, F, 1) = "/" THEN
        PSET (D + H + 4, E + 0), I
        PSET (D + H + 3, E + 1), I
        PSET (D + H + 3, E + 2), I
        PSET (D + H + 2, E + 3), I
        PSET (D + H + 1, E + 4), I
        PSET (D + H + 1, E + 5), I
        PSET (D + H, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "\" THEN
        PSET (D + H, E + 0), I
        PSET (D + H + 1, E + 1), I
        PSET (D + H + 1, E + 2), I
        PSET (D + H + 2, E + 3), I
        PSET (D + H + 3, E + 4), I
        PSET (D + H + 3, E + 5), I
        PSET (D + H + 4, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "[" THEN
        PSET (D + H, E + 0), I: PSET (D + H + 1, E + 0), I: PSET (D + H + 2, E + 0), I
        PSET (D + H, E + 1), I
        PSET (D + H, E + 2), I
        PSET (D + H, E + 3), I
        PSET (D + H, E + 4), I
        PSET (D + H, E + 5), I
        PSET (D + H, E + 6), I: PSET (D + H + 1, E + 6), I: PSET (D + H + 2, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "]" THEN
        PSET (D + H + 2, E + 0), I: PSET (D + H + 3, E + 0), I: PSET (D + H + 4, E + 0), I
        PSET (D + H + 4, E + 1), I
        PSET (D + H + 4, E + 2), I
        PSET (D + H + 4, E + 3), I
        PSET (D + H + 4, E + 4), I
        PSET (D + H + 4, E + 5), I
        PSET (D + H + 2, E + 6), I: PSET (D + H + 3, E + 6), I: PSET (D + H + 4, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "_" THEN
        PSET (D + H, E + 6), I: PSET (D + H + 1, E + 6), I: PSET (D + H + 2, E + 6), I: PSET (D + H + 3, E + 6), I: PSET (D + H + 4, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "" OR MID$(B$, F, 1) = "" THEN
        PSET (D + H, E + 0), I: PSET (D + H + 4, E + 0), I
        PSET (D + H, E + 2), I: PSET (D + H + 4, E + 2), I
        PSET (D + H, E + 3), I: PSET (D + H + 4, E + 3), I
        PSET (D + H, E + 4), I: PSET (D + H + 4, E + 4), I
        PSET (D + H, E + 5), I: PSET (D + H + 4, E + 5), I
        PSET (D + H + 1, E + 6), I: PSET (D + H + 2, E + 6), I: PSET (D + H + 3, E + 6), I
    ELSEIF MID$(B$, F, 1) = "!" THEN
        PSET (D + H + 1, E + 0), I: PSET (D + H + 2, E + 0), I: PSET (D + H + 3, E + 0), I
        PSET (D + H + 1, E + 1), I: PSET (D + H + 2, E + 1), I: PSET (D + H + 3, E + 1), I
        PSET (D + H + 2, E + 2), I
        PSET (D + H + 2, E + 5), I
        PSET (D + H + 2, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = ">" THEN
        PSET (D + H + 1, E + 0), I
        PSET (D + H + 2, E + 1), I
        PSET (D + H + 3, E + 2), I
        PSET (D + H + 4, E + 3), I
        PSET (D + H + 3, E + 4), I
        PSET (D + H + 2, E + 5), I
        PSET (D + H + 1, E + 6), I: GOTO 123
    ELSEIF MID$(B$, F, 1) = "@" THEN
        PSET (D + H, E + 0), 7: PSET (D + H + 1, E + 0), I: PSET (D + H + 2, E + 0), I: PSET (D + H + 3, E + 0), I: PSET (D + H + 4, E + 0), 7
        PSET (D + H, E + 1), I: PSET (D + H + 1, E + 1), 7: PSET (D + H + 2, E + 1), 7: PSET (D + H + 3, E + 1), 7: PSET (D + H + 4, E + 1), I
        PSET (D + H, E + 2), I: PSET (D + H + 1, E + 2), 7: PSET (D + H + 2, E + 2), I: PSET (D + H + 3, E + 2), I: PSET (D + H + 4, E + 2), I
        PSET (D + H, E + 3), I: PSET (D + H + 1, E + 3), 7: PSET (D + H + 2, E + 3), I: PSET (D + H + 3, E + 3), I: PSET (D + H + 4, E + 3), I
        PSET (D + H, E + 4), I: PSET (D + H + 1, E + 4), 7: PSET (D + H + 2, E + 4), I: PSET (D + H + 3, E + 4), I: PSET (D + H + 4, E + 4), I
        PSET (D + H, E + 5), I: PSET (D + H + 1, E + 5), 7: PSET (D + H + 2, E + 5), 7: PSET (D + H + 3, E + 5), 7: PSET (D + H + 4, E + 5), 7
        PSET (D + H, E + 6), 7: PSET (D + H + 1, E + 6), I: PSET (D + H + 2, E + 6), I: PSET (D + H + 3, E + 6), I: PSET (D + H + 4, E + 6), 7
    ELSEIF MID$(B$, F, 1) = "<" THEN
        PSET (D + H + 3, E + 0), I
        PSET (D + H + 2, E + 1), I
        PSET (D + H + 1, E + 2), I
        PSET (D + H, E + 3), I
        PSET (D + H + 1, E + 4), I
        PSET (D + H + 2, E + 5), I
        PSET (D + H + 3, E + 6), I: GOTO 123
    123 END IF
    IF F - 1 < C THEN GOTO 7.1
198 END SUB

