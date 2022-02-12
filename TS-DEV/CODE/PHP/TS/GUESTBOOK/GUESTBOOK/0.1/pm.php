<!-- Guestbook v0.1 PM (thomas.schilb@live.de) -->
<html>
<head>
<font face="Courier" size="2" color="red"><b>PM</b></font><br><br>
<form method="post" action="site.php?site=pm.process.php" target="_self">
<br>
Name:<br>
<input type="text" name="user_name" size="33" placeholder="Enter Your Name"><br><br>
Email:<br>
<input type="email" name="user_email" size="33" placeholder="Enter Your Email"><br><br>
Message:<br>
<textarea name="user_text" rows="3" cols="94" placeholder="Enter Your Message"></textarea><br><br>
<br>
<input type="submit" value="Submit">
</form>
<br>
<hr noshade color="#FFFFFF" width="776" size="1"><br>
<br>
<?php
include 'pm.view.table.php';
?>
</head>
</html>