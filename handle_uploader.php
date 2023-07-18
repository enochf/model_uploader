<?php
ini_set('memory_limit','1000M');
// ini_set('display_errors',E_ERROR);
// error_reporting(E_ALL);
// ======================= //
function getTime() { 
    $a = explode (' ',microtime()); 
    return(double) $a[0] + $a[1]; 
} 
$Start = getTime(); 
// ======================= //
session_start();
unset($_SESSION);
require('inc_functions.php');
require('inc_arrays.php');
require('inc_myinputs.php');
$spreadsheet = $_FILES['spreadsheet']['tmp_name'];
$_SESSION['count'] = 0;
if ($spreadsheet) {
	$destFile = 'files/test.txt';
	if (!copy($spreadsheet, $destFile)) {
		$_SESSION['message'] = '<div id="message">-- Your files did not upload properly 1 --</div>';
		header("Location:".$root."uploader.php");
		exit;
	}
	echo "Upload Named Cells: ".date('g:i:s')."\n\n";
	// get named cells
	$file = fopen($destFile, "r") or exit("Unable to open file! Try #1");
	echo "Upload Successful: ".date('g:i:s')."\n\n";
	while (($data = fgetcsv($file, 0, ",")) !== FALSE) {
		if($data[0] != 'Book' AND $data[1] != 'Sheet') {
			if($data[2] != '') {
				$namedCells[] = array($data[2],$data[1],$data[3]);
			}
		}
	}
	// convert cells to PHP
	echo "Start Convert Cells: ".date('g:i:s')."\n\n";
	$file = fopen($destFile, "r") or exit("Unable to open file! Try #2");
	echo "Upload Successful: ".date('g:i:s')."\n\n";
	while (($data = fgetcsv($file, 0, ",")) !== FALSE) {
		convertCell($data[0],$data[1],$data[2],$data[3],$data[4],$data[5]);
		// echo count($_SESSION['notdone'],1)."<br />";
	}
	$lookups = 1;
	// print_r($_SESSION['lookups']);
	// exit;
	// foreach($_SESSION['lookups'] as $k => $v) {
		// convertCell($v[0],$v[1],$v[2],$v[3],$v[4],$v[5]);
	// }
	echo "Start Finish: ".date('g:i:s')."\n\n";
	// echo "\nNot Done:\n";
	// print_r($_SESSION['notdone']);
	// exit;
	finishTree();
	// echo "BLANKS\n\n";
	// print_r($_SESSION['blanks']);
	// echo "\n\nEND BLANKS\n\n";
	// exit;
	setBlanks();
	// finishTree();
	createTree();
	// echo "\nDone:\n";
	$order = array_flip($_SESSION['done']);
	foreach($order as $k => $v) {
		$id = $_SESSION['inserted'][$v];
		$vals = array($k,$id);
		$q = "UPDATE cells SET runOrder=? WHERE cID=?";
		query($q,$vals,'u','UPDATE cell order');
	}
	// print_r($_SESSION['done']);
	echo "\nNot Done:\n";
	print_r($_SESSION['notdone']);
	// echo "\nCompleted:\n";
	// print_r($_SESSION['completed']);
	fclose($file);
	unset($_SESSION['done'],$_SESSION['notdone'],$_SESSION['completed'],$_SESSION['inserted']);
}
// ======================= //
$End = getTime(); 
echo "\n\n\n\nTime taken = ".number_format(($End - $Start),2)." secs to process the formula.\n\n\n\n"; 
// ======================= //
?>