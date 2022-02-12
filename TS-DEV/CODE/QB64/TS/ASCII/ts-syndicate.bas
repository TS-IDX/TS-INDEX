'
'
'          +--[ Syndicate ]---------------------------------+
'          |                                                |
'          |  Syndicate.cf ASCII v1.0.1                     |
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

SCREEN 12
_FULLSCREEN
CLS
COLOR 8
LOCATE 13, 14: PRINT "                       __ __            __"
LOCATE 14, 14: PRINT " .-----.--.--.-----.--|  |__.----.---.-|  |_.-----."
LOCATE 15, 14: PRINT " |__ --|  |  |     |  _  |  |  __|  _  |   _|  -__|"
LOCATE 16, 14: PRINT " |_____|___  |__|__|_____|__|____|___._|____|_____|"
LOCATE 17, 14: PRINT " -ts---|_____|------------- s y n d i c a t e . c f"
COLOR 4: LOCATE 13, 40: PRINT "__"
DO
    COLOR 7
    LOCATE 13, 14: PRINT "                       __ __            __"
    LOCATE 14, 14: PRINT " .-----.--.--.-----.--|  |__.----.---.-|  |_.-----."
    LOCATE 15, 14: PRINT " |__ --|  |  |     |  _  |  |  __|  _  |   _|  -__|"
    LOCATE 16, 14: PRINT " |_____|___  |__|__|_____|__|____|___._|____|_____|"
    LOCATE 17, 14: PRINT " -ts---|_____|------------- s y n d i c a t e . c f"
    COLOR 4: LOCATE 13, 40: PRINT "__"
    _DELAY 0.1
    COLOR 15
    LOCATE 13, 14: PRINT "                       __ __            __"
    LOCATE 14, 14: PRINT " .-----.--.--.-----.--|  |__.----.---.-|  |_.-----."
    LOCATE 15, 14: PRINT " |__ --|  |  |     |  _  |  |  __|  _  |   _|  -__|"
    LOCATE 16, 14: PRINT " |_____|___  |__|__|_____|__|____|___._|____|_____|"
    LOCATE 17, 14: PRINT " -ts---|_____|------------- s y n d i c a t e . c f"
    COLOR 12: LOCATE 13, 40: PRINT "__"
    _DELAY 0.1
    COLOR 7
    LOCATE 13, 14: PRINT "                       __ __            __"
    LOCATE 14, 14: PRINT " .-----.--.--.-----.--|  |__.----.---.-|  |_.-----."
    LOCATE 15, 14: PRINT " |__ --|  |  |     |  _  |  |  __|  _  |   _|  -__|"
    LOCATE 16, 14: PRINT " |_____|___  |__|__|_____|__|____|___._|____|_____|"
    LOCATE 17, 14: PRINT " -ts---|_____|------------- s y n d i c a t e . c f"
    COLOR 4: LOCATE 13, 40: PRINT "__"
    _DELAY 0.1
    COLOR 8

    LOCATE 13, 14: PRINT "                       __ __            __"
    LOCATE 14, 14: PRINT " .-----.--.--.-----.--|  |__.----.---.-|  |_.-----."
    LOCATE 15, 14: PRINT " |__ --|  |  |     |  _  |  |  __|  _  |   _|  -__|"
    LOCATE 16, 14: PRINT " |_____|___  |__|__|_____|__|____|___._|____|_____|"
    LOCATE 17, 14: PRINT " -ts---|_____|------------- s y n d i c a t e . c f"
    COLOR 12: LOCATE 13, 40: PRINT "__"
    _DELAY 0.1

LOOP WHILE INKEY$ = ""
COLOR 15
CLS
SYSTEM
