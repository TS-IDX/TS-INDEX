<?php

//
// TS-FIGLET 1.0
//
// This is a Clone of Text_Figlet written in PHP4.
//
// Any other ideas? mailto:thomas_schilb@outlook.com
//
// © 2021 TS. MIT. ALL RIGHTS RESERVED.
//

include 'Figlet.php';

$fonts = glob("fonts/*.flf");
@unlink('allfonts.htm');
$style = "<style>pre { font-size: 60% }</style>";
file_put_contents('allfonts.htm', $style, FILE_APPEND) ;
foreach ($fonts as $font) {
    echo $font . PHP_EOL;
    $figlet = new Text_Figlet();
    $figlet->LoadFont($font) ;
    preg_match('/\/(.*)\.flf/i', $font, $matches);
    $font = $matches[1];
    $text = PHP_EOL . "<b>$font</b>" . PHP_EOL;
    $text .= "<pre>".$figlet->LineEcho("TS")."</pre>" . PHP_EOL;
    file_put_contents('allfonts.htm', $text, FILE_APPEND);
}
