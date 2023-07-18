<?php
include("inc_conn.php");
function query($query,$vars,$type,$qname) {
	// echo "debug query($query)\n";
	// =============== EXAMPLE QUERIES =============== //
		/*
		// Insert
		$query = "INSERT INTO table VALUES('0',?,?,?,NOW(),'','1')";
		$newID = query($query,array($var1,$var2,$var3),'i','My Query Description');

		// Select a single row
		$query = "SELECT * FROM x_mu9";
		$rs = query($query,array(),'s','My Query Description');
		echo $rs[0]['name'];

		// Select multiple rows
		$query = "SELECT * FROM x_mu9";
		foreach(query($query,array(),'s','My Query Description') as $rs) {
			echo $rs['name'];
		}
		// Update
		$query = "UPDATE table SET name='?' WHERE id=?";
		$newID = query($query,array($name,$id),'u','My Query Description');

		// Delete
		$query = "DELETE FROM table WHERE id=?";
		$newID = query($query,array($id),'d','My Query Description');
		*/
	// =============================================== //
	global $dbh;
	$stmt = $dbh->prepare($query);
	if($stmt->execute($vars)) {
		if ($type == 'i') {
			return $dbh->lastInsertId();
		} else if ($type == 's') {
			return $stmt->fetchALL(PDO::FETCH_ASSOC);
		} else if ($type == 'u') {
			return true;
		} else if ($type == 'd') {
			return true;
		}
	} else {
		$err = $stmt->errorInfo();
		$num = 1;
		foreach($vars as $value) {
			$varStr.= '['.$num++.'] '.$value.', ';
		}
		$dt = date("ymdHis");
		$msg = $qname."\n\nQUERY: ".$query."\n\nVARIABLES: ".$varStr."\n\nERROR: ".$err[2];
		query('INSERT INTO errors (msg, dtAdded) VALUES (?,?)',array($msg,$dt),'i','WHY NOT WORKING');
		echo "Requested page is temporarily unavailable, please try again later.";
		exit;
	}
}

// Main function that initializes the conversion of an Excel cell to php code
function convertCell($data0,$data1,$data2,$data3,$data4,$data5) {
	global $myinputs;
	if($data0 != 'Book' AND $data1 != 'Sheet') {
		// if ($lookups != 1 AND preg_match('/(HLOOKUP|VLOOKUP)\(/',$data5)) {
			// $_SESSION['lookups'][] = array($data0,$data1,$data2,$data3,$data4,$data5);
			// return;
		// }
		$_SESSION['code']['sheet'] = $data1;
		$_SESSION['code']['cell'] = $data3;
		if($data5 == '') {
			$str = $data4;
			$offset = 0;
			$_SESSION['done'][$_SESSION['code']['sheet'].'!'.$_SESSION['code']['cell']] = $_SESSION['count']++; //count($_SESSION['done'])+1;
		} else {
			$str = $data5;
			$str = findConvertRefs($str,$namedCells);
			$offset = 1;
		}
		parseCell($str); // Begin cell conversion
		$_SESSION['code']['php'] = "\$cell['".$_SESSION['code']['sheet']."']['".$_SESSION['code']['cell']."'] =".$_SESSION['code']['php']; // add cell reference to begin of code
		$_SESSION['code']['php'].= ';'; // close the code
		foreach(array_reverse($_SESSION['code']['subformulas']) as $v) { // write nested statements and functions to the ['php'] variable
			$_SESSION['code']['php'] = $v."\n".$_SESSION['code']['php'];
		}
		// echo $_SESSION['code']['php']."\n"; // print out the final php code
		// echo "// ================== ".$_SESSION['code']['sheet']."!".$_SESSION['code']['cell']."\n".$_SESSION['code']['php']."\n\n"; // print out the final php code
		if (in_array($_SESSION['code']['cell'],$myinputs[$_SESSION['code']['sheet']])) {
			$input = 1;
		} else {
			$input = 0;
		}
		$vals = array($data1,$data3,$data2,$_SESSION['code']['php'],$input);
		$q = "INSERT INTO cells VALUES('0','11',?,?,?,?,'',?)";
		$newID = query($q,$vals,'i','INSERT Cell Code'); // print out the final php code
		$_SESSION['inserted'][$_SESSION['code']['sheet'].'!'.$_SESSION['code']['cell']] = $newID;
		unset($_SESSION['code']);
		createTree();
	}
}

// Begins parsing cell and breaking it into it's different components
function parseCell($str) {
	// echo "debug parseCell()\n";
	// echo "STRING: $str\n\n";
	if (substr($str,0,2) == '=+') {
		getVals(substr($str,2));
	} else if (substr($str,0,1) == '=') {
		getVals(substr($str,1));
	} else {
		if(is_numeric($str)) {
			$_SESSION['code']['php'].= ' '.$str;
		} else {
			$str = str_replace('"','\"',$str);
			$_SESSION['code']['php'].= ' "'.$str.'"';
		}
	}
}

// Break up a string into individual values while maintaining the highest level formulas
function getVals($str,$set) { // $set is an optional param that indicates an internal set initiated from the parseFormula() function
	// echo "debug getVals()\n";
	// echo "STRING: $str\n\n";
	// $vals = preg_split('/(([^\*|\+|\-|\/]*(?:\([^)]*\))[^\*|\+|\-|\/]*)+|\*|\+|\-|\/)/', $str, -1, PREG_SPLIT_NO_EMPTY | PREG_SPLIT_DELIM_CAPTURE | PREG_SPLIT_OFFSET_CAPTURE);
	getBracesPos($str,'(',')',0); // get initial braces in order to get $start and $type
	$start = $_SESSION['braces'][0][0]; // the first parenthese in the outer formula
	$delims = getDelims($str,'\+|-|\*|\/|<>|<|>'); // get all of the delims in $formVals
	foreach($delims[0] as $k => $v) { // loop through the delims and find all that are not within parentheses
		foreach($_SESSION['braces'] as $b) { // loop through braces
			if($v[1] > $b[0] AND $v[1] < $b[1]) { // check if dilim is inside of braces
				unset($delims[0][$k]); // unset delims inside of braces
			}
		}
	}
	if (count($delims[0]) != 0) { // check to see if there are delims
		foreach($delims[0] as $k => $v) { // loop through delims
			$vals[] = substr($str, $startPos, $v[1]-$startPos); // add value inside of delim to vals array
			$vals[] = substr($str, $delims[0][$k][1], 1); // add delim to vals array
			$startPos = $delims[0][$k][1]+1; // set new starting position for next value
		}
	}
	$vals[] = substr($str, $startPos); // add left over section of string as the last value of vals array
	if($set == 1) {
		$_SESSION['code']['php'].= '('; // add set start parenthesis
	}
	foreach($vals as $k => $v) { // loop through vals
		if(checkFormula($v) == true) { // check if the val is a formula or not
			$token = md5(rand(1000,9999)); // create a placeholder token for the formula
			parseFormula($v,$token); // parse the formula
			if($v[0] != '(') {
				$_SESSION['code']['php'].= ' '.'$val_'.substr($token,0,12); // add the placeholder code to the session
			}
		} else {
			if(preg_match_all('/(10\^[0-9]+)/',$v,$matches)) {
				$v = convertPow($v,$matches);
			}
			$_SESSION['code']['php'].= ' '.$v; // add value to the session
		}
	}
	if($set == 1) {
		$_SESSION['code']['php'].= ' )'; // add set end parenthesis
	}
}

// Removes duplicate function entries sesulting from getVals and the matching delimiters
function cleanVals($arr) {
	// echo "debug cleanVals()\n";
	foreach($arr as $k => $v) {
		if($arr[$k][1] != $arr[$k+1][1]) {
			$newArr[] = $v[0];
		}
	}
	return $newArr;
}

// Checks to see if the string includes a formula
function checkFormula($str) {
	// echo "debug checkFormula()\n";
	// if(preg_match('/IF\(|SUM\(|HLOOKUP\(|AVERAGE\(|ROUND\(|UPPER\(|CONCATENATE\(/', $str)) {
	if(preg_match('/(IF|SUM|HLOOKUP|AVERAGE|ROUND|UPPER|CONCATENATE|PPMT|ISPMT)\(/', $str)) {
		return true;
	} else {
		return false;
	}
}

// Break up a formula into the different components
function parseFormula($str,$token) {
	// echo "debug parseFormula()\n";
	// echo "STRING: $str\n\n";
	getBracesPos($str,'(',')',0); // get initial braces in order to get $start and $type
	$start = $_SESSION['braces'][0][0]; // the first parenthese in the outer formula
	$type = substr($str, 0, $start); // type of outer formula
	$formVals = substr($str, $start+1, -1); // get the formula without the formula type or trailing ")"
	getBracesPos($formVals,'(',')',0); // get all of the outer parentheses in $formVals
	$quotes = getQuotes($formVals);
	foreach($_SESSION['braces'] as $k => $v) {
		foreach($quotes as $k2 => $v2) {
			if($v[0] > $v2[0] AND $v[0] < $v2[1]) {
				unset($_SESSION['braces'][$k]);
			}
		}
	}
	if($str[0] == '(') { // check for set instead of formula
		// $delims = getDelims($str,'\+|-|\*|\/|<>|<|>'); // get all of the delims in $formVals
		getVals($formVals,1);
		return;
	} else {
		$delims = getDelims($formVals,','); // get all of the delims in $formVals
	}
	foreach($delims[0] as $k => $v) { // loop through the delims and find all that are not within parentheses
		foreach($_SESSION['braces'] as $b) {
			if($v[1] > $b[0] AND $v[1] < $b[1]) {
				unset($delims[0][$k]);
			}
		}
	}
	$startPos = 0;
	if(count($delims[0]) != 0) {
		$startPos = 0;
		foreach($delims[0] as $k => $v) {
			$vals[] = substr($formVals, $startPos, $v[1]-$startPos);
			$startPos = $delims[0][$k][1]+1;
		}
	}
	$vals[] = substr($formVals, $startPos);
	if($type == 'IF') { // IF statement
		$formula = funcIF($vals,$token);
	} else if($type == 'SUM') {
		$formula = funcSUM($vals,$token);
	} else if($type == 'ROUND') {
		$formula = funcROUND($vals,$token);
	} else if($type == 'UPPER') {
		$formula = funcUPPER($vals,$token);
	} else if($type == 'HLOOKUP') {
		$formula = funcHLOOKUP($vals,$token);
	} else if($type == 'CONCATENATE') {
		$formula = funcCONCATENATE($vals,$token);
	} else if($type == 'PPMT' OR $type == '-PPMT') {
		$formula = funcPPMT($vals,$token,$type[0]);
	} else if($type == 'ISPMT' OR $type == '-ISPMT') {
		$formula = funcISPMT($vals,$token,$type[0]);
	}// else if($type == '') {
		// $formula = funcSET($vals,$token);
	// }
	$_SESSION['code']['subformulas'][$token] = $formula;
}

// Get position of all delims in a string
function getDelims($str,$del) {
	// echo "debug getDelims()\n";
	preg_match_all('/'.$del.'/',$str,$matches,PREG_OFFSET_CAPTURE);
	return $matches;
}

// Get position of all delims in a string
function getQuotes($str) {
	// echo "debug getQuotes()\n";
	$quotes = getDelims($str,'(?<!\\\\)"');
	$num = 0;
	foreach($quotes[0] as $k => $v) {
		if ($num == 0) {
			$quoteSets[] = array($quotes[0][$k][1],$quotes[0][$k+1][1]);
			$num = 1;
		} else {
			$num = 0;
		}
	}
	return $quoteSets;
}

// Take the values and write the PHP version
function writeCode($str) {
	// echo "debug writeCode()\n";
	if(is_numeric($str)) {
		$_SESSION['code']['php'].= ' '.$str;
	} else {
		$_SESSION['code']['php'].= ' "'.str_replace('"','\"',$str).'"';
	}
}

// Find the start and close braces
function getBracesPos($source, $oB, $eB, $start) {
	// echo "debug getBracesPos()\n";
	if ($start == 0) {
		unset($_SESSION['braces']);
	}
	if (preg_match("/\\$oB.*\\$eB/", $source) > 0) {
		$open = 0;
		$length = strlen($source);
		for ($i = $start; $i < $length; $i++) {
			$currentChar = substr($source, $i, 1);
			if ($currentChar == $oB) {
				$open++;
				if ($open == 1) { // First open brace
					$firstOpenBrace = $i;
				}
			} else if ($currentChar == $eB) {
				$open--;
				if ($open == 0) { //time to wrap the roots
					$lastCloseBrace = $i;
					$_SESSION['braces'][] = array($firstOpenBrace, $lastCloseBrace);
					getBracesPos($source, $oB, $eB, $lastCloseBrace);
				}
			}
		}
	}
} 

// Not sure what this does yet
function cleanBraces($source, $oB, $eB) {
    // echo "debug cleanBraces()\n";
	$finalText = "";
    if (preg_match("/\\$oB.*\\$eB/", $source) > 0) {
        while (preg_match("/\\$oB.*\\$eB/", $source) > 0) {
            $brace = getBracesPos($source, $oB, $eB);
            $finalText .= substr($source, 0, $brace[0]);
            $source = substr($source, $brace[1] + 1, strlen($source) - $brace[1]);
        }
        $finalText .= $source;
    } else {
        $finalText = $source;
    }
    return $finalText;
}

// Convert the excel formatted cell name to an array
function convertRef($ref) {
	// echo "debug convertRef()\n";
	list($sheet,$cell) = explode('!', $ref);
	$_SESSION['code']['sheet'] = $sheet;
	$_SESSION['code']['cell'] = $cell;
}

// Match excel cell names and convert them into php array names
function findConvertRefs($str,$namedCells) {
	// echo "debug findConvertRefs()\n";
	// replace named cells with cell references
	foreach($namedCells as $k => $v) {
		if ($v[2] == 'Print_Area') {
			$str = str_replace($v[0],$v[2],$str);
		} else {
			$str = str_replace($v[0],$v[1].'!'.$v[2],$str);
		}
	}
	// convert cell references to PHP array reference
	$orginal = $str;
	preg_match_all('/(([A-Za-z])*\!?\$?([A-Za-z]{1,2})+\$?([0-9]{1,5})+)/',$str,$matches,PREG_OFFSET_CAPTURE);
	$matches = $matches[0];
	// foreach($matches[0] as $k => $v) {  NOT SURE IF THIS IS NECESSARY SINCE I ADDED THE OFFSET INTO THE ARRAY
		// if (preg_match('/!/',$v[0])) {
			// $matches[0][] = $v;
			// unset($matches[0][$k]);
		// }
	// }
	foreach($matches as $k => $v) {
		if (preg_match('/!/',$v[0])) {
			list($sheet,$cell) = explode('!',$v[0]);
			$noSheet = '';
		} else if (substr($orginal,$v[1]-1,1) == ':') {
			$sheet = $sheet;
			$cell = $v[0];
			$noSheet = '';
			$noSheet = "(?<=[^\['])";
		} else {			
			$sheet = $_SESSION['code']['sheet'];
			$cell = $v[0];
			$noSheet = "(?<=[^!|\$|\['])";
		}
		$new = "\$cell['".$sheet."']['".str_replace('$','',$cell)."']";
		// $str = preg_replace('/'.$v.'/',$new,$str,1);
		$search = '/'.$noSheet.'('.str_replace('$','\$',$v[0]).')/';
		$str = preg_replace($search,$new,$str,-1);
		$_SESSION['notdone'][$_SESSION['code']['sheet']][$_SESSION['code']['cell']][] = array($sheet,str_replace('$','',$cell)); 
	}
	return $str;
}

// Create dependency tree
function createTree() {
	// echo "debug createTree()\n";
	$redo = 0;
	foreach($_SESSION['notdone'] as $k1 => $sheet) {
		foreach($sheet as $k2 => $cell) {
			foreach($cell as $k3 => $dependent) {
				if(!$_SESSION['done'][$dependent[0].'!'.$dependent[1]]) {
					$check = 1;
					break;
				} else {
					unset($_SESSION['notdone'][$k1][$k2][$k3]);
				}
			}
			if ($check != 1) {
				$_SESSION['done'][$k1.'!'.$k2] = $_SESSION['count']++; //count($_SESSION['done']);
				$_SESSION['completed'][] = $k1.'!'.$k2;
				unset($_SESSION['notdone'][$k1][$k2]);
				$redo = 1;
			}
			$check = 0;
		}
	}
	if($redo == 1) {
		createTree();
	}
}

// Clean out "Not Done" array for any dependents that done
function finishTree() {
	// echo "debug finishTree()\n";
	foreach($_SESSION['notdone'] as $k1 => $sheet) {
		foreach($sheet as $k2 => $cell) {
			foreach($cell as $k3 => $dependent) {
				if(!$_SESSION['notdone'][$dependent[0]][$dependent[1]]) {
					// print_r($dependent);
					$_SESSION['blanks'][$dependent[0].'!'.$dependent[1]] = 1;
					// $_SESSION['blanks'][] = array($dependent[0],$dependent[1]);
					unset($_SESSION['notdone'][$k1][$k2][$k3]);
				}
			}
		}
	}
	// foreach($_SESSION['notdone'] as $k => $v) {
		// if(count($k) == 0) {
			// unset($_SESSION['notdone'][$k]);
		// } else {
			// finishTree();
		// }
	// }
}

// Create code for blank cells
function setBlanks() {
	global $myinputs;
	foreach($_SESSION['blanks'] as $k => $v) {
		list($sheet,$cell) = explode('!',$k);
		if (in_array($cell,$myinputs[$sheet])) {
			$input = 1;
		} else {
			$input = 0;
		}
		$code = "\$cell['".$sheet."']['".$cell."'] = '';";
		$vals = array($sheet,$cell,$code,$input);
		$q = "INSERT INTO cells VALUES('0','7',?,?,'',?,'0',?)";
		query($q,$vals,'i','INSERT Unreferenced Cell Code'); // print out the final php code
	}
}

// Create a list of cells in a set ie. B45:C50
function getCellSet($str) {
	// echo "debug getCellSet()\n";
	global $letters;
	list($first,$last) = explode(':', $str); // get the first and last cell in the set
	if ($first == $last) {
		$set[] = $first;
		return $set;
	}
	$cellFirst = explode("']['", $first); // extract cell reference from variable and sheet
	$sheet = explode("'",$cellFirst[0]); // extract sheet from remaining string
	$sheet = $sheet[1]; // assign sheet
	$first = substr($cellFirst[1],0,-2); // clean off reference
	$cellLast = explode("']['", $last); // extract cell reference from variable and sheet
	$last = substr($cellLast[1],0,-2); // clean off reference
	preg_match('/[1-9]/', $first, $firstPos, PREG_OFFSET_CAPTURE); // get position of first number in cell
	preg_match('/[1-9]/', $last, $lastPos, PREG_OFFSET_CAPTURE); // get position of first number in cell
	$firstCol = substr($first,0,$firstPos[0][1]); // get first column
	$firstRow = substr($first,$firstPos[0][1]); // get first row
	$lastCol = substr($last,0,$lastPos[0][1]); // get row #
	$lastRow = substr($last,$lastPos[0][1]); // get row #
	$firstChars = preg_split('//', $firstCol, -1, PREG_SPLIT_NO_EMPTY); // break column into array by individual characters
	$lastChars = preg_split('//', $lastCol, -1, PREG_SPLIT_NO_EMPTY); // break column into array by individual characters
	if ($firstCol == $lastCol) { // if both column names are the same
		$col = $firstCol; // set column
	} else {
		$colStart = array_search($firstCol,$letters); // set the start of a column difference
		$colEnd = array_search($lastCol,$letters); // set the end of a column difference
	}
	if ($colStart) { // check if there's a difference in columns
		for($i = $colStart; $i <= $colEnd; $i++) { // loop through letters to get the set full set of columns in range
			$columns[] = $col.$letters[$i]; // assign columns
		}
	} else {
		$columns[] = $col; // set the column range as a single column
	}
	foreach($columns as $k => $v) { // loop through column range
		for($i = $firstRow; $i <= $lastRow; $i++) { // loop through amount of rows
			$set[] = "\$cell['".$sheet."']['".$v.$i."']"; // set final names of all cells in range
			$_SESSION['notdone'][$_SESSION['code']['sheet']][$_SESSION['code']['cell']][] = array($sheet,$v.$i); 
		}
	}
	return $set;
}

// Create a list of cells in a set ie. B45:C50
function getHaystack($str) {
	// echo "debug getHaystack()\n";
	global $letters;
	list($first,$last) = explode(':', $str); // get the first and last cell in the set
	if ($first == $last) {
		$set[] = $first;
		return $set;
	}
	$cellFirst = explode("']['", $first); // extract cell reference from variable and sheet
	$sheet = explode("'",$cellFirst[0]); // extract sheet from remaining string
	$sheet = $sheet[1]; // assign sheet
	$first = substr($cellFirst[1],0,-2); // clean off reference
	$cellLast = explode("']['", $last); // extract cell reference from variable and sheet
	$last = substr($cellLast[1],0,-2); // clean off reference
	preg_match('/[1-9]/', $first, $firstPos, PREG_OFFSET_CAPTURE); // get position of first number in cell
	preg_match('/[1-9]/', $last, $lastPos, PREG_OFFSET_CAPTURE); // get position of first number in cell
	$firstCol = substr($first,0,$firstPos[0][1]); // get first column
	$firstRow = substr($first,$firstPos[0][1]); // get first row
	$lastCol = substr($last,0,$lastPos[0][1]); // get row #
	$lastRow = substr($last,$lastPos[0][1]); // get row #
	$firstChars = preg_split('//', $firstCol, -1, PREG_SPLIT_NO_EMPTY); // break column into array by individual characters
	$lastChars = preg_split('//', $lastCol, -1, PREG_SPLIT_NO_EMPTY); // break column into array by individual characters
	if ($firstCol == $lastCol) { // if both column names are the same
		$col = $firstCol; // set column
	} else {
		$colStart = array_search($firstCol,$letters); // set the start of a column difference
		$colEnd = array_search($lastCol,$letters); // set the end of a column difference
	}
	if ($colStart) { // check if there's a difference in columns
		for($i = $colStart; $i <= $colEnd; $i++) { // loop through letters to get the set full set of columns in range
			$columns[] = $col.$letters[$i]; // assign columns
		}
	} else {
		$columns[] = $col; // set the column range as a single column
	}
	foreach($columns as $k => $v) { // loop through column range
		$set['cells'][] = "\$cell[\"".$sheet."\"][\"".$v.$firstRow."\"]"; // set final names of all cells in range
		$_SESSION['notdone'][$_SESSION['code']['sheet']][$_SESSION['code']['cell']][] = array($sheet,$v.$firstRow); // add cell to "notdone" list
		// for($i = $firstRow; $i <= $lastRow; $i++) { // loop through amount of rows
			// $_SESSION['notdone'][$_SESSION['code']['sheet']][$_SESSION['code']['cell']][] = array($sheet,$v.$i); // add cell to "notdone" list
		// }
	}
	$set['settings'] = array($sheet,$firstCol,$firstRow);
	return $set;
}

// Convert 10^6 to pow(10,6)
function convertPow($str,$matches) {
	// echo "convertPow()\n";
	foreach($matches[0] as $k => $v) {
		$parts = explode('^',$v);
		$str = str_replace($v,'pow('.$parts[0].','.$parts[1].')',$str);
	}
	return $str;
}

/*
EXCEL FUNCTIONS:	CONCATENATE
					VLOOKUP
					ISPMT
					PPMT
					UPPER
					SUM
					ROUND
					HLOOKUP
					IF
*/

// Excel: () Set
function funcSET($vals,$token) {
	// echo "debug funcSET()\n";
	// print_r($vals);
	foreach($vals as $k => $v) {
		if(checkFormula($v) == true) {
			$tokenNew = md5(rand(1000,9999));
			parseFormula($v,$tokenNew);
			$formula.= '$val_'.substr($token,0,12).' = $val_'.substr($tokenNew,0,12);
		} else {
			if($v == '') {
				$v = '""';
			}
			$formula.= '$val_'.substr($token,0,12).' = '.$v;
		}
	}
	return $formula.';';
}

// Excel: IF() Statement
function funcIF($vals,$token) {
	// echo "debug funcIF()\n";
	$formula.= 'if(';
	if(checkFormula($vals[0]) == true) {
		$tokenNew = md5(rand(1000,9999));
		parseFormula($vals[0],$tokenNew);
		$formula.= '$val_'.substr($token,0,12).' = $val_'.substr($tokenNew,0,12);
	} else {
		// $formula.= str_replace(array('=','<>'), array('==','!='), $vals[0]);
		if(preg_match('/(<>)/',$vals[0])) {
			$vals[0] = str_replace('<>','!=',$vals[0]);
		} else if(!preg_match('/(>=|<=)/',$vals[0])) {
			$vals[0] = str_replace('=','==',$vals[0]);
		}
		$ifVals = preg_split('/(==|!=|>=|<=|<|>)/',$vals[0], -1, PREG_SPLIT_NO_EMPTY | PREG_SPLIT_DELIM_CAPTURE); // Break up the conditional values by the operator
		// if($ifVals[0][0] == '"') {
			// $ifVals[0] = "strtolower(".$ifVals[0].")";
		// }
		// if($ifVals[2][0] == '"') {
			// $ifVals[2] = "strtolower(".$ifVals[2].")";
		// }
		$formula.= 'trim(strtolower('.$ifVals[0].'))'.$ifVals[1].'trim(strtolower('.$ifVals[2].'))';
	}
	$formula.= ') { ';
	if(checkFormula($vals[1]) == true) {
		$tokenNew = md5(rand(1000,9999));
		parseFormula($vals[1],$tokenNew);
		$formula.= '$val_'.substr($token,0,12).' = $val_'.substr($tokenNew,0,12);
	} else {
		if($vals[1] == '') {
			$vals[1] = '""';
		}
		$formula.= '$val_'.substr($token,0,12).' = '.$vals[1];
	}
	$formula.= '; } else { ';
	if(checkFormula($vals[2]) == true) {
		$tokenNew = md5(rand(1000,9999));
		parseFormula($vals[2],$tokenNew);
		$formula.= '$val_'.substr($token,0,12).' = $val_'.substr($tokenNew,0,12);
	} else {
		if($vals[2] == '') {
			$vals[2] = '""';
		}
		$formula.= '$val_'.substr($token,0,12).' = '.$vals[2];
	}
	return $formula.'; }';
}

// Excel: SUM() Formula
function funcSUM($vals,$token) {
	// echo "debug funcSUM()\n";
	// print_r($vals);
	$formula.= '$val_'.substr($token,0,12).' = ';
	foreach($vals as $k => $v) {
		if(checkFormula($v) == true) {
			$tokenNew = md5(rand(1000,9999));
			parseFormula($v,$tokenNew);
			$formula.= '$val_'.substr($tokenNew,0,12).' + ';
		} else {
			if(preg_match('/:/',$v)) {
				$set = getCellSet($v);
				foreach($set as $k => $v) {
					$formula.= $v.' + ';
				}
			} else {
				$formula.= $v.' + ';
			}
		}
	}
	return substr($formula,0,-3).";";
}

// Excel: ROUND() Formula
function funcROUND($vals,$token) {
	// echo "debug funcROUND()\n";
	$formula = '$val_'.substr($token,0,12).' = round('.$vals[0].','.$vals[1].');';
	return $formula;
}

// Excel: UPPER() Formula
function funcUPPER($vals,$token) {
	// echo "debug funcUPPER()\n";
	$formula = '$val_'.substr($token,0,12).' = strtoupper('.$vals[0].');';
	return $formula;
}

// Excel: CONCATENATE() Formula
function funcCONCATENATE($vals,$token) {
	// echo "debug funcCONCATENATE()\n";
	$formula = '$val_'.substr($token,0,12).' = ';
	foreach($vals as $k => $v) {
		if(checkFormula($v) == true) {
			$tokenNew = md5(rand(1000,9999));
			parseFormula($v,$tokenNew);
			$formula.= '$val_'.substr($token,0,12).' = $val_'.substr($tokenNew,0,12);
		} else {
			$formula.= $v.'.';
		}
	}
	return rtrim($formula,'.').';';
}

// Excel: HLOOKUP() Formula
function funcHLOOKUP($vals,$token) {
	// echo "debug funcHLOOKUP()\n";
	$token = substr($token,0,12);
	// GET NEEDLE
	if(checkFormula($vals[0]) == true) {
		$tokenNew = md5(rand(1000,9999));
		parseFormula($v,$tokenNew);
		$formula.= '$val_'.substr($tokenNew,0,12);
	} else {
		$formula.= '$needle_'.$token.' = '.$vals[0].";\n";
	}
	// GET HAYSTACK
	$formula.= '$haystack_'.$token.' = \''.serialize(getHaystack($vals[1]))."';\n";
	$formula.= '$haystack = unserialize($haystack_'.$token.");\n";
	$formula.= 'foreach($haystack[\'cells\'] as $k => $v) {
	eval("\$v = ".$v.";");
	if(trim(strtolower($v)) == trim(strtolower($needle_'.$token.'))) {
		$key = array_search($haystack[\'settings\'][1],$letters);
		$col = $letters[$key+$k];
		$val_'.$token.' = $cell[$haystack[\'settings\'][0]][$col.($haystack[\'settings\'][2]+('.$vals[2].'-1))];
		break;
	}
}
unset($haystack_'.$token.',$haystack);';
	return $formula;
}

// Excel: PPMT() Formula
function funcPPMT($vals,$token,$av) {
	// echo "debug funcPPMT()\n";
	if ($av == '-') {
		$av1 = 'abs(';
		$av2 = ')';
	}
	if ($vals[4]) {
		$opt1 = ', '.$vals[4];
	}
	if ($vals[5]) {
		$opt2 = ', '.$vals[5];
	}
	$formula = '$val_'.substr($token,0,12).' = '.$av1.'PPMT('.$vals[0].', '.$vals[1].', '.$vals[2].', '.$vals[3].$opt1.$opt2.')'.$av2.';';
	return $formula;
}

// Excel: ISPMT() Formula
function funcISPMT($vals,$token,$av) {
	// echo "debug funcISPMT()\n";
	if ($av == '-') {
		$av1 = 'abs(';
		$av2 = ')';
	}
	if ($vals[4]) {
		$opt1 = ', '.$vals[4];
	}
	if ($vals[5]) {
		$opt2 = ', '.$vals[5];
	}
	$formula = '$val_'.substr($token,0,12).' = '.$av1.'ISPMT('.$vals[0].', '.$vals[1].', '.$vals[2].', '.$vals[3].$opt1.$opt2.')'.$av2.';';
	return $formula;
}

/*






















*/
?>