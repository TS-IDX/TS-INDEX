<!-- Guestbook v0.1 PM (thomas.schilb@live.de) -->
<?php
include 'pm.connect.php';
$mysqli = mysqli_connect($mysql_host, $mysql_username, $mysql_password, $mysql_database);

$query = mysqli_query($mysqli, "SELECT * FROM pm");

while ($row = mysqli_fetch_array($query)) {




#tabelle
echo "<table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' bordercolor='#808080' width='776'>
  <tr>
    <td>";
echo "<font style='Courier' size='2'>". $row['id'] . ". <a href='mailto:" . $row['user_email'] . "'>".$row['user_name']."</a></font><br>";

echo "</td>
	</tr>
  <tr>
    <td>
	";
$row_user_message =  wordwrap( $row['user_message'], 96, "\n", true ); #zeilenumbruch für user_message
echo "<font style='Courier' size='2'>" . "$row_user_message\n" . "</font><br><br>";

echo "</td>
  </tr>
</table></font>";
}
mysqli_close($mysqli);
?>