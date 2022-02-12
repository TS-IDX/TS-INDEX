'
'
'          +--[ Mystify ]-----------------------------------+
'          |                                                |
'          |  Mystify v1.0.0                                |
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

_FULLSCREEN
CONST Sh = 2, Vx = 4, Ln = 10, MaxSpeed = 6, SV = Sh * Vx, SV2 = SV * 2, Ln1 = Ln + 1
RANDOMIZE TIMER
SCREEN 12
DIM Posn%(Sh - 1, Vx - 1, 1, Ln), Vec%(Sh - 1, Vx - 1, 1)
FOR A% = 0 TO 2 * Sh * Vx - 1
    Vec%(A% \ Vx MOD Sh, A% MOD Vx, A% \ SV) = -INT(RND * MaxSpeed + 1) * (INT(RND * 2) * 2 - 1)
    Posn%(A% \ Vx MOD Sh, A% MOD Vx, A% \ SV, 0) = INT(RND * 480)
NEXT
DO
    FOR A% = 0 TO SV2 * Ln1 - 1
        IF A% MOD SV2 = 0 THEN WAIT &H3DA, 8
        LINE (Posn%(A% MOD Sh, A% \ Sh MOD Vx, 0, (A% \ SV2 + (A% \ SV MOD 2) * 2) MOD Ln1) * 1.33, Posn%(A% MOD Sh, A% \ Sh MOD Vx, 1, (A% \ SV2 + (A% \ SV MOD 2) * 2) MOD Ln1))-(Posn%(A% MOD Sh, (A% \ Sh MOD Vx + 1) MOD Vx, 0, (A% \ SV2 + (A% \ SV MOD 2) * 2) MOD Ln1) * 1.33, Posn%(A% MOD Sh, (A% \ Sh MOD Vx + 1) MOD Vx, 1, (A% \ SV2 + (A% \ SV MOD 2) * 2) MOD Ln1)), (INT(TIMER \ (3 + (A% MOD Sh) * 1.618)) MOD 15 + 1) * (1 - A% \ SV MOD 2)
        Posn%(A% MOD Sh, A% \ Sh MOD Vx, A% \ SV MOD 2, (A% \ SV2 + 1) MOD Ln1) = Posn%(A% MOD Sh, A% \ Sh MOD Vx, A% \ SV MOD 2, A% \ SV2) + Vec%(A% MOD Sh, A% \ Sh MOD Vx, A% \ SV MOD 2)
        IF Posn%(A% MOD Sh, A% \ Sh MOD Vx, A% \ SV MOD 2, (A% \ SV2 + 1) MOD Ln1) < 0 OR Posn%(A% MOD Sh, A% \ Sh MOD Vx, A% \ SV MOD 2, (A% \ SV2 + 1) MOD Ln1) > 479 THEN
            Vec%(A% MOD Sh, A% \ Sh MOD Vx, A% \ SV MOD 2) = -SGN(Vec%(A% MOD Sh, A% \ Sh MOD Vx, A% \ SV MOD 2)) * INT(RND * MaxSpeed + 1)
            Posn%(A% MOD Sh, A% \ Sh MOD Vx, A% \ SV MOD 2, (A% \ SV2 + 1) MOD Ln1) = ((Posn%(A% MOD Sh, A% \ Sh MOD Vx, A% \ SV MOD 2, (A% \ SV2 + 1) MOD Ln1) + 240) \ 479) * 479
        END IF
    NEXT
    _DELAY 0.01
'    IF LEN(INKEY$) = 0 THEN GOTO 1
LOOP WHILE INKEY$ = ""
SYSTEM
