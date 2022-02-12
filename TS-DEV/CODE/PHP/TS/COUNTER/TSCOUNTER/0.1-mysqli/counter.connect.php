<!-- Counter v0.1 (thomas.schilb@live.de)--> 
<?php

$db_location = "localhost"; //Ort der Datenbank (normalerweise "localhost")
$db_username = "root"; //Benutzername für DB-Zugriff
$db_passwort = ""; //Passwort für DB-Zugriff
$db_name = "counter"; //Name der Datenbank

@mysql_connect($db_location, $db_username, $db_passwort) or die ("OFFLINE<br />"); 
@mysql_select_db($db_name) or die ("DATABASE DOES NOT EXISTS<br />"); 

?>