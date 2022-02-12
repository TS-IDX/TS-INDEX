<!-- since forum v0.3 (thomas.schilb@live.de) -->
<?php

# 0.0

$db_location = "localhost"; //Ort der Datenbank (normalerweise "localhost")
$db_username = "username"; //Benutzername für DB-Zugriff
$db_passwort = "password"; //Passwort für DB-Zugriff
$db_name = "databasename"; //Name der Datenbank

# 0.1
$con = mysqli_connect($db_location, $db_username, $db_passwort) or die ('OFFLINE'); 
mysqli_select_db($con, $db_name) or die ('DATABASE DOES NOT EXISTS'); 

?>