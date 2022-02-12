<?php
# ----------------------------------
# -         newsgrab v0.5.0 -
# ----------------------------------
#            2019/04/13

# get url to .html-file

// Lese 14 Zeichen, beginnend mit dem 21. Zeichen
$section = file_get_contents('http://www.heise.de/newsticker/sidebar1.html', NULL, NULL, 20, 14);
var_dump($section);
?>


<?php

  function cleanEmptyspace($string) { 
    return trim( preg_replace('/\s+/', ' ', $string) ); 
} 

//truncate file to zero (for fresh news)
$fp = fopen("sidebar1.html", "r+");
ftruncate($fp, 0);
fclose($fp);

$url='http://www.heise.de/newsticker/sidebar1.html';
//fopen opens webpage in Binary
$handle=fopen($url,"rb");
// initialize
$lines_string="";
// read content line by line
do{
	$data=fread($handle,1024);
	if(strlen($data)==0) {
		break;
	}
	$lines_string.=$data;
}while(true);
//close handle to release resources
fclose($handle);
//output, you can also save it locally on the server

$datei = fopen("sidebar1.html","a");

# filter content
$lines_string_filtered1 = strip_tags($lines_string, '<p><a>'); // alle html tags ausser p und a entfernen
$lines_string_filtered2 = substr($lines_string_filtered1, 190); // entfernen erste 190 zeichen 
$lines_string_filtered3 = cleanEmptyspace($lines_string_filtered2); // entfernen alle leerzeichen 

$lines_string_filtered4 = str_replace("neu laden","",$lines_string_filtered3); // entfernen neu laden
$lines_string_filtered41 = str_replace("(","",$lines_string_filtered4); // klammer offen weg
$lines_string_filtered42 = str_replace(")","<br>",$lines_string_filtered41); // klammer zu weg

$lines_string_filtered43 = str_replace("*","<b><font color='#FF0000'>&bull;</font>&nbsp;&nbsp;</b>",$lines_string_filtered42); // sternzeichen zu punkt - aufzählung

$lines_string_filtered44 = str_replace("</a>","</a><br><br>",$lines_string_filtered43); // entfernen

$lines_string_filtered45 = str_replace("<br><br><br>","<br><br>",$lines_string_filtered44); // entfernen

$lines_string_filtered46 = str_replace('<a href="http://www.heise.de/" target="_content"> </a><br><br> ','',$lines_string_filtered45); // entfernen

$lines_string_filtered5 = str_replace("News ","",$lines_string_filtered46); // news weg

fwrite($datei, $lines_string_filtered5); // filter content speichern
include ("sidebar1.html"); // ausgabe des gespeicherten
?>
