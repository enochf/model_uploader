<?php
// ini_set('memory_limit','256M');
// ini_set('display_errors',E_ERROR);
// error_reporting(E_ALL);
// ======================= //
// function getTime() { 
    // $a = explode (' ',microtime()); 
    // return(double) $a[0] + $a[1]; 
// } 
// $Start = getTime(); 
// ======================= //
session_start();
echo $_SESSION['message'];
unset($_SESSION['message']);
// ======================= //
// $End = getTime(); 
// echo "\n\n\n\nTime taken = ".number_format(($End - $Start),2)." secs to process the following formula ".$num." times.\n\n".$str."\n\nPHP Code after conversion:\n\n".$_SESSION['code']['cell'].$_SESSION['code']['php']; 
// ======================= //
?>
<form action="handle_uploader.php" method="post" id="import" name="import" enctype="multipart/form-data">
Spreadsheet: <input type="file" name="spreadsheet" class="file" /><br /><br />
<input type="submit" value="Upload" class="btn button">
</form><br /><br />
<a href="handle_getscript.php">Generate Script</a>