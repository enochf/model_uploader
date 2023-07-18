<?php
session_start();
include("inc_functions.php");
$vars = array();
$q = "SELECT sheet,cell,code,input FROM cells WHERE vID=11 ORDER BY runOrder";
foreach(query($q,$vars,'s','Get all cells in order') as $rs) {
	if($rs['input'] == 1) {
		$code.= 'if($_SESSION[\'inputs\'][$sid][\''.$rs['sheet'].'\'][\''.$rs['cell'].'\']) { $cell[\''.$rs['sheet'].'\'][\''.$rs['cell'].'\'] = $_SESSION[\'inputs\'][$sid][\''.$rs['sheet'].'\'][\''.$rs['cell'].'\']; } else { '.$rs['code']." }\n";
	} else {
		$code.= $rs['code']."\n";
	}
}
echo $code;
?>