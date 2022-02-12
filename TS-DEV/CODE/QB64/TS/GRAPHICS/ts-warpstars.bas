'
'
'          +--[ WarpStars ]---------------------------------+
'          |                                                |
'          |  WarpStars v1.0.1                              |
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

CONST MAXSTARS = 300 ' 50 to 1000 stars
CONST WARPSPEED = 10 ' [0=fast to 10=slow]
_FULLSCREEN
SCREEN 12

TYPE starData
    x AS SINGLE ' x increment value
    y AS SINGLE ' y increment value
    c AS INTEGER ' star colour
    xx AS SINGLE ' x position
    yy AS SINGLE ' y position
END TYPE

DIM i%, s%
DIM star(MAXSTARS) AS starData
 
FOR i% = 1 TO MAXSTARS 'initialize star settings
    star(i%).x = ((RND * 2) - 1) * 3
    star(i%).y = ((RND * 2) - 1) * 3
    SELECT CASE (ABS(star(i%).x) + ABS(star(i%).y))
        CASE IS < 2: star(i%).c = 8
        CASE IS < 4: star(i%).c = 7
        CASE IS <= 6: star(i%).c = 15
    END SELECT
    star(i%).xx = (((star(i%).x * 215) / 2) + 320)
    star(i%).yy = (((star(i%).y * 135) / 2) + 200)
NEXT i%
 
DO ' display starfield, main program loop
    FOR i% = 1 TO MAXSTARS
        PSET (star(i%).xx - star(i%).x, star(i%).yy - star(i%).y), 0
        star(i%).xx = star(i%).xx + star(i%).x
        star(i%).yy = star(i%).yy + star(i%).y
        PSET (star(i%).xx, star(i%).yy), star(i%).c
        IF star(i%).yy > 483 OR star(i%).yy < -4 THEN star(i%).xx = 320: star(i%).yy = 240
        IF star(i%).xx > 643 OR star(i%).xx < -4 THEN star(i%).xx = 320: star(i%).yy = 240
        FOR s% = 0 TO WARPSPEED * 1000: NEXT s%
    NEXT i%
LOOP UNTIL INKEY$ <> ""
SYSTEM

