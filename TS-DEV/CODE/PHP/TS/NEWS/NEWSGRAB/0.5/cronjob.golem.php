<?php
# ----------------------------------
# -         newsgrab v0.5.0 -
# ----------------------------------
#            2019/04/13

# get url to .html-file

// Lese 14 Zeichen, beginnend mit dem 21. Zeichen
$section = file_get_contents('http://www.golem.de/htmlheadlines.htm', NULL, NULL, 20, 14);
var_dump($section);
?>


<?php

  function cleanEmptyspaces($string) { 
    return trim( preg_replace('/\s+/', ' ', $string) ); 
} 

#$homepage = file_get_contents('https://www.golem.de/htmlheadlines.htm');
#echo $homepage;
#heise.newsticker.html = sidebar1.html

//truncate file to zero (for fresh news)
$fp = fopen("golem.html", "r+");
ftruncate($fp, 0);
fclose($fp);

$url='http://www.golem.de/htmlheadlines.htm';
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
	$lines_string .= $data;
}while(true);
//close handle to release resources
fclose($handle);
//output, you can also save it locally on the server

$datei = fopen("golem.html","a");

# filter content
$lines_string_filtered1 = strip_tags($lines_string, '<tr><td><p><a><br>'); // alle html tags ausser p und a entfernen
$lines_string_filtered2 = substr($lines_string_filtered1, 0); // erste 0 zeichen entfernen
$lines_string_filtered3 = cleanEmptyspaces($lines_string_filtered2); // alle leerzeichen entfernen
$lines_string_filtered4 = str_replace('<a class="golemdeheadline"' , '<b><font color="#FF0000">&bull;</font></b>&nbsp;&nbsp;<a' , $lines_string_filtered3); // punkt hinzufügen
$lines_string_filtered5 = str_replace('<br />' , '<br><br>' , $lines_string_filtered4); // punkt hinzufügen



fwrite($datei, $lines_string_filtered5); // filter content speichern

include ("golem.html"); // ausgabe des gespeicherten
?>

