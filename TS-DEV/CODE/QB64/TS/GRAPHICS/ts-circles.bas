'
'
'          +--[ Mystify ]-----------------------------------+
'          |                                                |
'          |  Circles v1.0.0                                |
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
CLS

SCREEN 12

RANDOMIZE TIMER



'sets the palette



'sets randomness of palette

pal = INT(RND * 5) + 1

IF pal = 2 THEN pal = 256 'green

IF pal = 3 THEN pal = 257 'yellow

IF pal = 4 THEN pal = 65536 'blue

IF pal = 5 THEN pal = 65537 'purple

       

'FOR i = 60 TO 1 STEP -1 'bright in back

FOR i = 1 TO 60 STEP 4 'bright in front

    x = x + 1

    PALETTE x, i * pal

NEXT

'PALETTE 1, 0



'initiation stuff



size = 23

delay = 5000

winX = 160

winY = 240

accel = .8

maxsp = size / 2

minsp = -maxsp



DIM x(size), y(size)

DIM xx(size), yy(size), nx(size), ny(size)

DIM col(size)

DIM r(size)

DIM t(size), v(size), u(size)





'stuffs the arrays



maxX = 640 - winX

maxY = 480 - winY

x = maxX

y = maxY

FOR i = 1 TO size

    xx(i) = minsp

    yy(i) = minsp

    x = x - 400 / size

    y = y - 400 / size

    x(i) = x

    y(i) = y

    r(i) = i

    col(i) = i

NEXT



DO



    'delay



    FOR i = 1 TO delay STEP 1: NEXT



    FOR i = 1 TO size



        'circles that erase



        CIRCLE (t(i), v(i)), u(i), 0

        'CIRCLE (640 - t(i), 480 - v(i)), u(i), 0

        'CIRCLE (t(i), 480 - v(i)), u(i), 0

        'CIRCLE (640 - t(i), v(i)), u(i), 0



        'circles that draw



        'this is the real one, the rest are mirrors

        CIRCLE (x(i), y(i)), r(i), col(i)

       

        'CIRCLE (640 - x(i), 480 - y(i)), r(i), col(i)

        'CIRCLE (x(i), 480 - y(i)), r(i), col(i)

        'CIRCLE (640 - x(i), y(i)), r(i), col(i)



        IF i < size THEN



            'first line is the real one, the rest are mirrors.  No eraser lines are

            'needed because the tail color is black.



            'LINE (x(i), y(i))-(x(i + 1), y(i + 1)), col(i)

            'LINE (640 - x(i), 480 - y(i))-(640 - x(i + 1), 480 - y(i + 1)), col(i)

            'LINE (640 - x(i), y(i))-(640 - x(i + 1), y(i + 1)), col(i)

            'LINE (x(i), 480 - y(i))-(x(i + 1), 480 - y(i + 1)), col(i)



            'connects the lines for some spiffy effects



            'LINE (640 - x(i), y(i))-(640 - x(i + 1), 480 - y(i + 1)), col(i)

            'LINE (x(i), y(i))-(640 - x(i + 1), 480 - y(i + 1)), col(i)

            'LINE (x(i), y(i))-(x(i + 1), 480 - y(i + 1)), col(i)

            'LINE (x(i), 480 - y(i))-(640 - x(i + 1), y(i + 1)), col(i)



        END IF



        'sets eraser variables



        t(i) = x(i): v(i) = y(i): u(i) = r(i)



    NEXT



    'the big bad movement phase



    FOR i = 1 TO size

        'the way this works is the leader object is moved according to these

        'conditions, and then the followers just move along in a line



        IF i = size THEN

            'accelerates, checking for maxsp and minsp



            xx(i) = xx(i) + nx(i)

            yy(i) = yy(i) + ny(i)

            IF xx(i) > maxsp THEN xx(i) = maxsp

            IF xx(i) < minsp THEN xx(i) = minsp

            IF yy(i) > maxsp THEN yy(i) = maxsp

            IF yy(i) < minsp THEN yy(i) = minsp



            'moves the leader object, checking to see if it is in the

            'window



            x(i) = x(i) + xx(i)

            y(i) = y(i) + yy(i)

            IF x(i) > maxX THEN nx(i) = -RND * accel - accel

            IF x(i) < winX THEN nx(i) = RND * accel + accel

            IF y(i) > maxY THEN ny(i) = -RND * accel - accel

            IF y(i) < winY THEN ny(i) = RND * accel + accel

        END IF



        'lets the followers follow



        IF i < size THEN

            x(i) = x(i + 1)

            y(i) = y(i + 1)

        END IF



    NEXT


    _DELAY 0.1
LOOP UNTIL INKEY$ <> ""
SYSTEM


