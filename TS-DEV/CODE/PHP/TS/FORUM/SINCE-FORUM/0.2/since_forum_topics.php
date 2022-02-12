<!-- since forum v0.2 (thomas.schilb@live.de) -->
<?php
ob_start();
$id = (int) $_GET['id'];
$idtopic = (int) $_GET['idtopic'];
if ($id < 1)
{
	header('Location: site.php?site=since_forum.php');
	exit();
}
ob_end_clean();
?>
<html>
<head>
<font size="2" color="red"><b>Forum<font size="2" color="white">:&nbsp;since</font></b><a target="_self" href="site.php?site=since_forum.php">&nbsp;<img border="0" src="images/up.gif" width="11" height="9"></a><br><br>
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
<?php
include 'since_forum_connect.php';
$output1 = mysql_fetch_assoc(mysql_query("SELECT name FROM since_forum_main WHERE id = $id"));
#echo "<b>" . $output1['name'] . "</b><br><br>";
?>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#808080" width="100%">
<tr>
<td><font size="2"><b>Topic</b></font></td>
<td><font size="2"><b>Name</b></font></td>
<td><font size="2"><b>Date</b></font></td>
<td><font size="2"><b>Last post</b></font></td>
<td><font size="2"><b>Replies</b></font></td>
</tr>
<tr>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
</tr>
<?php
$query2 = mysql_query("SELECT * FROM since_forum_topics WHERE forumid = '$id' ORDER BY id ASC");
$query3 = mysql_num_rows($query2);
if ($query3 == 0)
	echo '<td colspan="5"><font face="Courier New" size="2">No Topics</font></td>';
else
{
	while ($output2 = mysql_fetch_assoc($query2))
	{
		echo '<tr>';
		echo '<td><font face="Courier New" size="2"><a target="_self" href="site.php?site=since_forum_replies.php&id='.$output2['id'].'&idtopic='.$id.'">'.wordwrap( $output2['subject'], 43, "\n", true ).'</a></td></font>';
		echo '<td><font face="Courier New" size="2" color="#808080">'.wordwrap( $output2['poster'], 15, "\n", true ).'</td></font>';
		echo '<td><font face="Courier New" size="2">'.date('D-m-y G:i', $output2['date']).'</td></font>';
		if(empty($output2['lastposter']))
			echo '<td colspan="2"><font face="Courier New" size="2" color="#808080">No replies</font></td>';
		else
		{
			echo '<td><font face="Courier New" size="2" color="#808080">'.$output2['lastposter'].' @ '.date('d-m-y G:i', $output2['lastpostdate']).'</td></font>';
			echo '<td><font face="Courier New" size="2">'.$output2['replies'].'</td></font>';
		}
		echo '</tr>';
	}
}
?>
</table>
<br /><br /><hr /><br /><br />
<!-- <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.1.1/jquery.min.js"></script> -->
<script src="jquery.min.js"></script>
<script>
$(document).ready(function() {
   $('textarea[maxlength]').keyup(function(){
      var max = parseInt($(this).attr('maxlength'));
 
      if($(this).val().length > max){
         $(this).val($(this).val().substr(0, max));
      }
      
      $(this).parent().find('.charleft').html(max - $(this).val().length);
   });  
});
$(document).ready(function() {
   $('input[maxlength]').keyup(function(){
      var max = parseInt($(this).attr('maxlength'));
 
      if($(this).val().length > max){
         $(this).val($(this).val().substr(0, max));
      }
      
      $(this).parent().find('.charleft').html(max - $(this).val().length);
   });  
});
</script>
<form target="_self" name="form1" id="form1" method="post" action="site.php?site=since_forum_forumpost.php&type=topics&id=<? echo $id; ?>&idtopic=<? echo $idtopic; ?>">
<b>Add Topic</b><br /><br />
<input maxlength="25" name="subject" id="subject" type="text" size="25" placeholder="Topic (25)" /><br /><br />
<textarea maxlength="480" name="message" id="message" cols="94" placeholder="Description (480)" rows="5"></textarea><br /><br />
<input maxlength="15" name="poster" id="poster" type="text" size="15" placeholder="Name (15)" /><br /><br />
<input type="submit" name="submit" id="submit" value="Submit" />
</form>
</body>
</html>