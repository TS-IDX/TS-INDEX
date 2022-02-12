<!-- Counter v0.1 (thomas.schilb@live.de)--> 
<html>
<head>
<link rel="icon" 
      type="image/png" 
      href="favicon.png">
<title>Counter v0.1</title>

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
	font-family: Courier;
    font-size: 12px;
}

input, select, button{
    font-family: inherit;
    font-size: inherit;
}
</style>
</head>

<body link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF" topmargin="0" leftmargin="0" text="#FFFFFF" bgcolor="#000000">

<?php

# Connect to mysql database

include 'counter.connect.php';

# Get host IP

$ip = $_SERVER['REMOTE_ADDR'];

# Test if IP is given in database or insert into

$test = mysql_query("SELECT * FROM counter WHERE ip='$ip'");

# Otherwise make new entry

if(mysql_num_rows($test) != 1) {
 mysql_query("INSERT INTO counter SET ip='$ip'");
}

# Visitor number

$count = mysql_num_rows(mysql_query("SELECT cid FROM counter"));

# Show count and current IP

echo "count: " . $count . "</br>ip: " . $ip ."<br /><br>iplist:<br><br>";

# Current IPs

$test2 = mysql_query("SELECT * FROM counter");
while ($row = mysql_fetch_array($test2)) {
$showip = $row['ip'];
echo "<a href='http://ip-lookup.net/?ip=".$showip."'>".$showip."</a><br>";
}
?>

</body>

</html>