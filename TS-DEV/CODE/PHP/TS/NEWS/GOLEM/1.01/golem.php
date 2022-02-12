<?php
/* 
   Newsgrabber - Golem.de
   ----------------------------
   Version: 1.0.1
   Author: Thomas Schilb
   Email: thomas.schilb@live.de
   Last Update: 2019/11/08
   Filename: golem.php
   ----------------------------
*/
header('Content-Type: text/html; charset=utf-8');
/* simplexml load file from url */
$xml_file=simplexml_load_file("https://rss.golem.de/rss.php?feed=ATOM1.0") or die("Error: Cannot load from file");
/* prepare the content, make arrays, filter content */
foreach($xml_file->children() as $entries) {
$title = $entries->title;
$link = $entries->link[href];
$xml_data_title = array($title);
$xml_data_link  = array($link);
if ($xml_data_title[0] == "") {} // Ignore Nullified Content
else {
$xml_raw_data = '<font color="#FFF">&bull;</font>&nbsp;<a href="'.$xml_data_link[0].'">'.$xml_data_title[0].'</a><br><br>';
$xml_filter_00 = str_replace('<a href=""></a><br><br>','',$xml_raw_data);;
$xml_filter_01 = str_replace('<a href="','<a target="_blank" href="',$xml_filter_00);;
$xml_data = $xml_filter_01;
/* show content, title & link */
echo $xml_data;}
}
?>