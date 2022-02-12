<!-- since forum v0.3 (thomas.schilb@live.de) -->
<html>
<head>

<?php
ob_start();
$id = (int) $_GET['id'];
$idtopic = (int) $_GET['idtopic'];
$type = $_GET['type'];
if ($id < 1 || ($type != 'replies' && $type != 'topics'))
{
	header('Location: since_forum.mysqli.php');
	exit();
}
ob_end_clean();
?>
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
<body>
<?php
function clear($message)
{
	if(!get_magic_quotes_gpc())
		$message = addslashes($message);
	$message = strip_tags($message);
	$message = htmlentities($message);
	return trim($message);
}
if ($_POST['submit'])
{
include 'since_forum_connect.mysqli.php';
	$message0 = clear($_POST['message']);
	$message= nl2br($message0); # zeilenumbruch speichern
	$subject = clear($_POST['subject']);
	$poster = clear($_POST['poster']);
	$date = time();
	if($type == 'topics')
	{
		$query = mysqli_fetch_assoc(mysqli_query($con, "SELECT topics FROM since_forum_main WHERE id = '$id'"));
		$topics = $query['topics'] + 1;
		mysqli_query($con, "UPDATE since_forum_main SET topics = '$topics', lastposter = '$poster', lastpostdate = '$date' WHERE id = '$id'");
		mysqli_query($con, "INSERT INTO since_forum_topics (id , forumid , message , subject, poster, date, lastposter, lastpostdate, replies) VALUES ('', '$id', '$message', '$subject','$poster', '$date', '', '', '0')");
		echo '<br><br><center><font color="#FF0000"><b>Topic Posted. <a target="_self" href="site.php?site=since_forum_topics.php&id='.$id.'">View Topic</a></b></font></center>';
		echo '<meta http-equiv="refresh" content="3; URL=site.php?site=since_forum_topics.php&id='.$id.'">';
	}
	else
	{
		$query = mysqli_fetch_assoc($con, mysqli_query("SELECT replies, forumid FROM since_forum_topics WHERE id = '$id'"));
		$replies = $query['replies'] + 1;
		$id2 =  $query['forumid'];
		mysqli_query($con, "UPDATE since_forum_topics SET replies = '$replies', lastposter = '$poster', lastpostdate = '$date' WHERE id = '$id'");
		$query = mysqli_fetch_array(mysqli_query($con, "SELECT replies FROM since_forum_main WHERE id = '$id2'"));
		$replies = $query['replies'] + 1;
		mysqli_query($con, "UPDATE since_forum_main SET replies = '$replies', lastposter = '$poster', lastpostdate = '$date' WHERE id = '$id2'");
		mysqli_query($con, "INSERT INTO since_forum_replies (id , topicid, message, subject, poster, date) VALUES ('', '$id', '$message', '$subject','$poster', '$date')");
		echo '<br><br><center><font color="#FF0000"><b>Reply Posted. <a target="_self" href="site.php?site=since_forum_replies.mysqli.php&id='.$id.'&idtopic='.$idtopic.'">View Reply</a></b></font></center>';
		echo '<meta http-equiv="refresh" content="3; URL=site.php?site=since_forum_replies.mysqli.php&id='.$id.'&idtopic='.$idtopic.'">';

	}
}
?>
</head>
</body>
</html>