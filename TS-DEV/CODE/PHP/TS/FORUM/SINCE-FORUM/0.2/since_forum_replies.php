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
</style></head>
<body>
<font size="2" color="red"><b>Forum<font size="2" color="white">:&nbsp;since</font></b></font><a target="_self" href="site.php?site=since_forum_topics.php&id=<? echo $idtopic; ?>">&nbsp;<img border="0" src="images/up.gif" width="11" height="9"></a><br><br>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#808080" width="100%">
<?php
include 'since_forum_connect.php';
$query1 = mysql_query("SELECT * FROM since_forum_replies WHERE topicid = $id ORDER BY id ASC");
$query2 = mysql_num_rows($query1);
$output2 = mysql_fetch_assoc(mysql_query("SELECT * FROM since_forum_topics WHERE id = $id"));
echo '<tr><td><font size="2"><b>'.$output2['subject'].' - Posted by <strong>'.$output2['poster'].'</strong></b></font></td>';
echo '<td><font size="2">'.date('D-m-y G:i', $output2['date']).'</font></td></tr>';
?>
<tr>
<td>&nbsp;</td>
<td>&nbsp;</td>
</tr>
<?php
echo '<tr><td colspan="2"><font size="2"><i>'.wordwrap( $output2['message'], 96, "\n", true ).'</i></font></td></tr>';
echo '<tr><font size="2"><td colspan="2">&nbsp;</td></tr>';
if ($query2 == 0)
	echo '<td colspan="2"><font size="2">No Replies</font></td>';
else
{
	while ($output = mysql_fetch_assoc($query1))
	{
		echo '<tr><td><font size="2" color="red"><b>'.wordwrap( $output['subject'], 82, "\n", true ).'</b></font></td><td><font color="white" " size="2">'.date('d-m-y G:i', $output['date']).'</font></td></tr>';
		$row_message =  wordwrap( $output['message'], 96, "\n", true ); #zeilenumbruch für message
		echo '<tr><td colspan="2"><font size="2"><b>'.$row_message.'<br /><strong>Posted by '.$output['poster'].'</strong></b></font></td></tr>';
echo "<tr>";
echo "<td>&nbsp;</td>";
echo "<td>&nbsp;</td>";
echo "</tr>";
	}
}
?>
</table>
<br /><hr /><br /><br />
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
<form target="_self" name="form1" id="form1" method="post" action="site.php?site=since_forum_forumpost.php&type=replies&id=<? echo $id; ?>&idtopic=<? echo $idtopic; ?>">
<b><font size="2" color="red">Add Reply</font></b><br /><br />
<input maxlength="82" name="subject" id="subject" type="text" size="82" placeholder="Title (82)" /><br /><br />
<textarea maxlength="960" name="message" id="message" cols="94" rows="10" placeholder="Message (960)"></textarea><br /><br />
<input maxlength="15" name="poster" id="poster" type="text" size="15" placeholder="Name (15)" /><br /><br />
<input type="submit" name="submit" id="submit" value="Submit" />
</form>
</body>
</html>