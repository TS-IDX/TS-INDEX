<!-- since forum v0.3 (thomas.schilb@live.de) -->
<?php
# 0.0

$db_location = "localhost"; //Ort der Datenbank (normalerweise "localhost")
$db_username = "web378"; //Benutzername für DB-Zugriff
$db_passwort = "A1b2c3d4e5"; //Passwort für DB-Zugriff
$db_name = "usr_web378_1"; //Name der Datenbank

# 0.1
$con = mysqli_connect($db_location, $db_username, $db_passwort) or die ('OFFLINE'); 
mysqli_select_db($con, $db_name) or die ('DATABASE DOES NOT EXISTS'); 

# 0.2
// mysqli?
?>