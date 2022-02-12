<?php

//
// TS-FIGLET 1.0
//
// This is a Clone of Text_Figlet written in PHP4.
//
// Any other ideas? mailto:thomas_schilb@outlook.com
//
// © 2021 TS. MIT. ALL RIGHTS RESERVED.
//

require 'Figlet.php';

function render($font, $font2, $text) {
    try {
        $font = "flf/$font.flf";
        if (!file_exists($font)) {
            throw new Exception('File not found');
        }
        $figlet = new Text_Figlet();
        $figlet->LoadFont($font);
        $figtext = $figlet->LineEcho($text) ;
        $_SESSION['text'] = $figtext . PHP_EOL;
        $_SESSION['font'] = $font2;
        return $figtext;
    } catch (Exception $e) {
        die($e->getMessage());
    }
}

?>

<!doctype html>
<head>
    <link rel="stylesheet" href="css.css" />
    <LINK REL="ICON" TYPE="IMAGE/X-ICON" HREF="TS.ICO">
    <title>TS | FIGLET</title>
</head>
<font>
<pre class='text'>
 ______   ______        ______   __     ______     __         ______     ______  
/\__  _\ /\  ___\      /\  ___\ /\ \   /\  ___\   /\ \       /\  ___\   /\__  _\ 
\/_/\ \/ \ \___  \     \ \  __\ \ \ \  \ \ \__ \  \ \ \____  \ \  __\   \/_/\ \/ 
   \ \_\  \/\_____\     \ \_\    \ \_\  \ \_____\  \ \_____\  \ \_____\    \ \_\ 
    \/_/   \/_____/      \/_/     \/_/   \/_____/   \/_____/   \/_____/     \/_/ 
</pre>
<br><hr>
<?php
session_start();
if (!isset($_GET['text'])) {
    $text = 'TS';
    $font = 'sub-zero';
    $font2 = 'Source Code Pro';
} else {
    $text = $_GET['text'];
    $font = $_GET['font'];
    $font2 = 'SourceCodePro-Regular';
    #$font2 = $_GET['font2'];
    if (!isset($_SESSION['text'])) {
        echo "<h2 class='error'>** Please enable cookies. **</h2>";
    }
}

?>
<br><br>
<form action="./index.php" method="get">
&nbsp;&nbsp;&nbsp;Enter any text: <input type="text" size="25" name="text" value="<?php echo $text ?>"><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Figlet Font: <select name='font'>
<?php
foreach (glob('flf/*.flf') as $file) {
    $file = pathinfo($file);
    $filename = $file['filename'];
    $selected = $font == $filename ? 'selected' : null;
    echo "<option $selected value='$filename'>$filename</option>";
}
?>
</select><br><br><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" value="Generate">
</form>
<br><br><hr><br>

<br/>
<pre class='text'>
<?php echo render($font, $font2, $text) ?>
</pre>