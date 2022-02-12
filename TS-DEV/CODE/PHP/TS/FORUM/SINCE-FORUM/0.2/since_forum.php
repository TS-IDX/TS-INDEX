<!-- since forum v0.2 (thomas.schilb@live.de) -->
<html>
<head>
<link rel="icon" 
      type="image/png" 
      href="favicon.png">
<meta http-equiv="Content-Language" content="de">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>since - digital assorted underground news, links and software!</title>
<base target="_blank">
<meta name="author" content="Thomas Schilb">
<meta name="publisher" content="Thomas Schilb">
<meta name="copyright" content="Thomas Schilb">
<meta name="description" content="since - digital assorted underground news, links and software!">
<meta name="keywords" content="audio, code, design, develope, ebook, emulate, erotic, ftp, games, intern, link, music, os, portable, scene, security, share, torrent, tv, unlock, xdcc, since, zince, thomas, schilb, facebook, soundcloud, twitter, youtube">
<meta name="page-topic" content="Bildung">
<meta name="page-type" content="Private Homepage">
<meta name="audience" content="Fortgeschrittene">
<meta name="robots" content="all, index, follow"> 
<meta name="googlebot" content="all, index, follow"> 
<meta name="pagerank" content="10"> 
<meta name="msnbot" content="all,index,follow"> 
<meta name="revisit" content="2 Days"> 
<meta name="revisit-after" content="2 Days"> 
<meta name="alexa" content="100">
<style type="text/css">
a {
	text-decoration: none;
}
a:hover {
	color: red;
}
img:hover {
	color: white;
}
body {
	background-color: black;
	font-family: "Courier New", Courier, monospace;
    font-size: 10px;
}

input, select, button, textarea{
    font-family: inherit;
    font-size: inherit;
}
p {
    font-family: "Courier New", Courier, monospace;
}
</style>
</head>
<body>
<font size="2" color="red"><b>Forum<font size="2" color="white">:&nbsp;since</font></b></font>
<br><br>
<b>
<font  size="2" color="#808080"><center><i><span lang="de">&quot;Welcome to my text-based forum. You can exchange anything<br>like comments, ideas, links but only with care of sense.<br>It is free and you have not to register.&quot;</span></i></center></font></b><br>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#808080" width="100%">
<tr>

<td><font face="Courier New" size="2"><b>Category</b></font></td>
<td><font face="Courier New" size="2"><b>Topics</b></font></td>
<td><font face="Courier New" size="2"><b>Replies</b></font></td>
<td><font face="Courier New" size="2"><b>Last Post</b></font></td>
</tr>
<tr>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
</tr>
<?php
include 'since_forum_connect.php';
$query1 = mysql_query('SELECT * FROM since_forum_main ORDER BY id ASC'); 
while ($output1 = mysql_fetch_assoc($query1))
{
	#$idtopic = $output1['id'];
	echo '<tr>';
	echo '<td><font face="Courier New" size="2"><a target="_self" href="site.php?site=since_forum_topics.php&id='.$output1['id'].'&idtopic='.$output1['id'].'">'.$output1['name'].'</a></font></td>';
	echo '<td><font face="Courier New" size="2">'.$output1['topics'].'</font></td>';
	echo '<td><font face="Courier New" size="2">'.$output1['replies'].'</font></td>';
	if (empty($output1['lastposter']))
		echo '<td><font face="Courier New" size="2">No Posts</font></td>';
	else
		echo '<td><font face="Courier New" size="2">'.$output1['lastposter'].' @ '.date('d-m-y G:i', $output1['lastpostdate']).'</font></td>';
	echo'</tr>';
}
?>
</table>
</font>
</body>
</html>