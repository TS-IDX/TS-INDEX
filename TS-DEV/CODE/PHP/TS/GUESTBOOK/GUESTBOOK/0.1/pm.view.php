<!-- Guestbook v0.1 PM (thomas.schilb@live.de) -->
<?php
include 'pm.connect.php';
$mysqli = mysqli_connect($mysql_host, $mysql_username, $mysql_password, $mysql_database);

$query = mysqli_query($mysqli, "SELECT * FROM pm");
#	$query_delete = mysqli_query($mysqli, "DELETE * FROM since_users_data WHERE id='3'");

while ($row = mysqli_fetch_array($query)) {
echo $row['id'] . ". <a href='mailto:" . $row['user_email'] . "'>".$row['user_name']."</a><br>";
echo $row['user_message'] . "<br><br>";
}
mysqli_close($mysqli);
?>