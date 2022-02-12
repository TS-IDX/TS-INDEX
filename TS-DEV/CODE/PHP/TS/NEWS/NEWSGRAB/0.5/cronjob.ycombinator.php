<?php
# ----------------------------------
# -         newsgrab v0.5.0 -
# ----------------------------------
#            2019/04/13

# get url to .html-file

$rawdata = getHtml('https://news.ycombinator.com/news');

  function cleanWhitespace($string) { 
    return trim( preg_replace('/\s+/', ' ', $string) ); 
} 

//truncate file to zero (for fresh news)
$fp = fopen("ycombinator.html", "r+");
ftruncate($fp, 0);
fclose($fp);


$datei = fopen("ycombinator.html","a");

# filter content to simple .html-file

$lines_string_filtered1 = strip_tags($rawdata, '<a><p><table><tr><td><div>'); // alle html tags ausser: a, p, table, tr, td, div entfernen
$lines_string_filtered11 = substr($lines_string_filtered1, 0); // entfernen erste 190 zeichen 
$lines_string_filtered2 = cleanWhitespace($lines_string_filtered11); // entfernen alle leerzeichen 
$lines_string_filtered3 = str_replace('Hacker News','',$lines_string_filtered2);
$lines_string_filtered4 = str_replace('bgcolor="#f6f6ef"','',$lines_string_filtered3);
$lines_string_filtered5 = str_replace('<td bgcolor="#ff6600">','<td>',$lines_string_filtered4);
$lines_string_filtered50 = str_replace('<table border="0" cellpadding="0" cellspacing="0" width="100%" style="padding:2px"><tr><td style="width:18px;padding-right:4px"><a href="https://news.ycombinator.com"></a></td> <td style="line-height:12pt; height:10px;"><a href="news"></a> <a href="newest">new</a> | <a href="newcomments">comments</a> | <a href="show">show</a> | <a href="ask">ask</a> | <a href="jobs">jobs</a> | <a href="submit">submit</a> </td><td style="text-align:right;padding-right:4px;"> <a href="login?goto=news">login</a> </td> </tr></table>','',$lines_string_filtered5);
$lines_string_filtered50b = str_replace('<a href="news"></a> <a href="newest">new</a> | <a href="front">past</a> | <a href="newcomments">comments</a> | <a href="ask">ask</a> | <a href="show">show</a> | <a href="jobs">jobs</a> | <a href="submit">submit</a> </td><td style="text-align:right;padding-right:4px;"> <a href="login?goto=news">login</a>','',$lines_string_filtered50);
$lines_string_filtered50c = str_replace('<a href="news?p=2" class="morelink" rel="next">More</a>','',$lines_string_filtered50b);
$lines_string_filtered50d = str_replace('<a href="newsguidelines.html">Guidelines</a> | <a href="newsfaq.html">FAQ</a> | <a href="mailto:hn@ycombinator.com">Support</a> | <a href="https://github.com/HackerNews/API">API</a> | <a href="security.html">Security</a> | <a href="lists">Lists</a> | <a href="bookmarklet.html" rel="nofollow">Bookmarklet</a> | <a href="http://www.ycombinator.com/legal/">Legal</a> | <a href="http://www.ycombinator.com/apply/">Apply to YC</a> | <a href="mailto:hn@ycombinator.com">Contact</a>Search: ','',$lines_string_filtered50c);
$lines_string_filtered51 = str_replace('<a href="news?p=2" class="morelink" rel="nofollow">More</a>','',$lines_string_filtered50d);
$lines_string_filtered52 = str_replace('<a href="newsguidelines.html">Guidelines</a> | <a href="newsfaq.html">FAQ</a> | <a href="mailto:hn@ycombinator.com">Support</a> | <a href="https://github.com/HackerNews/API">API</a> | <a href="security.html">Security</a> | <a href="lists">Lists</a> | <a href="bookmarklet.html" rel="nofollow">Bookmarklet</a> | <a href="dmca.html">DMCA</a> | <a href="http://www.ycombinator.com/apply/">Apply to YC</a> | <a href="mailto:hn@ycombinator.com">Contact</a>Search:','',$lines_string_filtered51);
$lines_string_filtered6 = str_replace(".</td>",".&nbsp;</td>",$lines_string_filtered52); // javascript entfernen 4
$lines_string_filtered7 = str_replace('<tr class="morespace" style="height:10px"></tr><tr>','',$lines_string_filtered6);
$lines_string_filtered8 = str_replace('&nbsp;comments</a>','&nbsp;comments</a><br><br>',$lines_string_filtered7);
$lines_string_filtered9 = str_replace('&nbsp;comment</a>','&nbsp;comment</a><br><br>',$lines_string_filtered8);
$lines_string_filtered10 = str_replace('discuss</a>','discuss</a><br><br>',$lines_string_filtered9);
$lines_string_filtered101 = str_replace('hide</a> </td>','hide</a><br><br></td>',$lines_string_filtered10);
$lines_string_filtered102 = str_replace('<a href="https://www.ycombinator.com/apply/"> Applications are open for YC Summer 2018 </a><a href="newsguidelines.html">Guidelines</a> | <a href="newsfaq.html">FAQ</a> | <a href="mailto:hn@ycombinator.com">Support</a> | <a href="https://github.com/HackerNews/API">API</a> | <a href="security.html">Security</a> | <a href="lists">Lists</a> | <a href="bookmarklet.html" rel="nofollow">Bookmarklet</a> | <a href="http://www.ycombinator.com/legal">Legal</a> | <a href="http://www.ycombinator.com/apply">Apply to YC</a> | <a href="mailto:hn@ycombinator.com">Contact</a>Search: ','',$lines_string_filtered101);
#$lines_string_filtered200 = str_replace('<td class="title"><a href="news?p=2" class="morelink" rel="next">More</a></td>','',$lines_string_filtered102);
#$lines_string_filtered201 = str_replace('<a href="newsguidelines.html">Guidelines</a> | <a href="newsfaq.html">FAQ</a> | <a href="mailto:hn@ycombinator.com">Support</a> | <a href="https://github.com/HackerNews/API">API</a> | <a href="security.html">Security</a> | <a href="lists">Lists</a> | <a href="bookmarklet.html" rel="nofollow">Bookmarklet</a> | <a href="http://www.ycombinator.com/legal/">Legal</a> | <a href="http://www.ycombinator.com/apply/">Apply to YC</a> | <a href="mailto:hn@ycombinator.com">Contact</a>Search:','',$lines_string_filtered200);
#$lines_string_filtered202 = str_replace('<a href="news"></a> <a href="newest">new</a> | <a href="front">past</a> | <a href="newcomments">comments</a> | <a href="ask">ask</a> | <a href="show">show</a> | <a href="jobs">jobs</a> | <a href="submit">submit</a> </td><td style="text-align:right;padding-right:4px;"> <a href="login?goto=news">login</a>','',$lines_string_filtered201);
# create points instead of numbers

$lines_string_filtered11 = str_replace('1.&nbsp;','<font color="#FF0000"><b>&bull;</b></font>&nbsp;&nbsp;',$lines_string_filtered102);
$lines_string_filtered12 = str_replace('2.&nbsp;','<font color="#FF0000"><b>&bull;</b></font>&nbsp;&nbsp;',$lines_string_filtered11);
$lines_string_filtered13 = str_replace('3.&nbsp;','<font color="#FF0000"><b>&bull;</b></font>&nbsp;&nbsp;',$lines_string_filtered12);
$lines_string_filtered14 = str_replace('4.&nbsp;','<font color="#FF0000"><b>&bull;</b></font>&nbsp;&nbsp;',$lines_string_filtered13);
$lines_string_filtered15 = str_replace('5.&nbsp;','<font color="#FF0000"><b>&bull;</b></font>&nbsp;&nbsp;',$lines_string_filtered14);
$lines_string_filtered16 = str_replace('6.&nbsp;','<font color="#FF0000"><b>&bull;</b></font>&nbsp;&nbsp;',$lines_string_filtered15);
$lines_string_filtered17 = str_replace('7.&nbsp;','<font color="#FF0000"><b>&bull;</b></font>&nbsp;&nbsp;',$lines_string_filtered16);
$lines_string_filtered18 = str_replace('8.&nbsp;','<font color="#FF0000"><b>&bull;</b></font>&nbsp;&nbsp;',$lines_string_filtered17);
$lines_string_filtered19 = str_replace('9.&nbsp;','<font color="#FF0000"><b>&bull;</b></font>&nbsp;&nbsp;',$lines_string_filtered18);
$lines_string_filtered20 = str_replace('10.&nbsp;','<font color="#FF0000"><b>&bull;</b></font>&nbsp;&nbsp;',$lines_string_filtered19);
$lines_string_filtered21 = str_replace('11.&nbsp;','<font color="#FF0000"><b>&bull;</b></font>&nbsp;&nbsp;',$lines_string_filtered20);
$lines_string_filtered22 = str_replace('12.&nbsp;','<font color="#FF0000"><b>&bull;</b></font>&nbsp;&nbsp;',$lines_string_filtered21);
$lines_string_filtered23 = str_replace('13.&nbsp;','<font color="#FF0000"><b>&bull;</b></font>&nbsp;&nbsp;',$lines_string_filtered22);
$lines_string_filtered24 = str_replace('14.&nbsp;','<font color="#FF0000"><b>&bull;</b></font>&nbsp;&nbsp;',$lines_string_filtered23);
$lines_string_filtered25 = str_replace('15.&nbsp;','<font color="#FF0000"><b>&bull;</b></font>&nbsp;&nbsp;',$lines_string_filtered24);
$lines_string_filtered26 = str_replace('16.&nbsp;','<font color="#FF0000"><b>&bull;</b></font>&nbsp;&nbsp;',$lines_string_filtered25);
$lines_string_filtered27 = str_replace('17.&nbsp;','<font color="#FF0000"><b>&bull;</b></font>&nbsp;&nbsp;',$lines_string_filtered26);
$lines_string_filtered28 = str_replace('18.&nbsp;','<font color="#FF0000"><b>&bull;</b></font>&nbsp;&nbsp;',$lines_string_filtered27);
$lines_string_filtered29 = str_replace('19.&nbsp;','<font color="#FF0000"><b>&bull;</b></font>&nbsp;&nbsp;',$lines_string_filtered28);
$lines_string_filtered30 = str_replace('20.&nbsp;','<font color="#FF0000"><b>&bull;</b></font>&nbsp;&nbsp;',$lines_string_filtered29);
$lines_string_filtered31 = str_replace('21.&nbsp;','<font color="#FF0000"><b>&bull;</b></font>&nbsp;&nbsp;',$lines_string_filtered30);
$lines_string_filtered32 = str_replace('22.&nbsp;','<font color="#FF0000"><b>&bull;</b></font>&nbsp;&nbsp;',$lines_string_filtered31);
$lines_string_filtered33 = str_replace('23.&nbsp;','<font color="#FF0000"><b>&bull;</b></font>&nbsp;&nbsp;',$lines_string_filtered32);
$lines_string_filtered34 = str_replace('24.&nbsp;','<font color="#FF0000"><b>&bull;</b></font>&nbsp;&nbsp;',$lines_string_filtered33);
$lines_string_filtered35 = str_replace('25.&nbsp;','<font color="#FF0000"><b>&bull;</b></font>&nbsp;&nbsp;',$lines_string_filtered34);
$lines_string_filtered36 = str_replace('26.&nbsp;','<font color="#FF0000"><b>&bull;</b></font>&nbsp;&nbsp;',$lines_string_filtered35);
$lines_string_filtered37 = str_replace('27.&nbsp;','<font color="#FF0000"><b>&bull;</b></font>&nbsp;&nbsp;',$lines_string_filtered36);
$lines_string_filtered38 = str_replace('28.&nbsp;','<font color="#FF0000"><b>&bull;</b></font>&nbsp;&nbsp;',$lines_string_filtered37);
$lines_string_filtered39 = str_replace('29.&nbsp;','<font color="#FF0000"><b>&bull;</b></font>&nbsp;&nbsp;',$lines_string_filtered38);
$lines_string_filtered40 = str_replace('30.&nbsp;','<font color="#FF0000"><b>&bull;</b></font>&nbsp;&nbsp;',$lines_string_filtered39);
$lines_string_filtered41 = str_replace('1<font color="#FF0000"><b>&bull;</b></font>&nbsp;&nbsp;','<font color="#FF0000"><b>&bull;</b></font>&nbsp;&nbsp;',$lines_string_filtered40);
$lines_string_filtered42 = str_replace('2<font color="#FF0000"><b>&bull;</b></font>&nbsp;&nbsp;','<font color="#FF0000"><b>&bull;</b></font>&nbsp;&nbsp;',$lines_string_filtered41);

# save filtered contect to  .html-file

fwrite($datei, $lines_string_filtered42);

# view of saved .html-file

include ("ycombinator.html");

# get url to .html-file (function)

function getHtml($url, $post = null) {
    $ch = curl_init();
    curl_setopt($ch, CURLOPT_URL, $url);
    curl_setopt($ch, CURLOPT_RETURNTRANSFER, 1);
    curl_setopt($ch, CURLOPT_FOLLOWLOCATION, 1);
    curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, false);
    curl_setopt($ch, CURLOPT_SSL_VERIFYHOST, 0);
    if(!empty($post)) {
        curl_setopt($ch, CURLOPT_POST, true);
        curl_setopt($ch, CURLOPT_POSTFIELDS, $post);
    } 
    $result = curl_exec($ch);
    curl_close($ch);
    return $result;
}

?>
