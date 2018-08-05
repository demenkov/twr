<?php

require('vendor/autoload.php');

$fileName = 'sample.xlsx';
//$fileName = 'test.xlsx';

$reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReaderForFile($fileName);
$reader->setReadDataOnly(true);
$spreadsheet = $reader->load($fileName);

$sheets = [];

foreach ($spreadsheet->getSheetNames() as $index => $sheetName) {
	$sheets[$index] = [
		'title' => $sheetName,
		'data' => [],
		'total' => 1,
		'weekly' => [],
		'monthly' => [],
		'quarter' => [],
	];
	$worksheet = $spreadsheet->getSheetByName($sheetName);
	$highestRow = $worksheet->getHighestRow();
	$highestColumn = $worksheet->getHighestColumn();
	$highestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColumn);
	for ($row = 2; $row <= $highestRow; ++$row) {
	    	$worksheet->getStyle('A' . $row)->getNumberFormat()->setFormatCode(
		        \PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_DATE_YYYYMMDD2
		    );
	    for ($col = 1; $col <= $highestColumnIndex; ++$col) {
	        $sheets[$index]['data'][$row][$col] = $worksheet->getCellByColumnAndRow($col, $row)->getFormattedValue();
	    }

	    $sheets[$index]['data'][$row][1] = strtotime($sheets[$index]['data'][$row][1]);
	    if (!isset($sheets[$index]['data'][$row-1]) || !is_numeric($sheets[$index]['data'][$row-1][2])) {
	    	continue;
	    }
	    $begin = $sheets[$index]['data'][$row-1][2] ?: 1;
	    $curr = ($sheets[$index]['data'][$row][2] + $sheets[$index]['data'][$row][3]);
	    $sheets[$index]['data'][$row][4] = ($curr / $begin) - 1;

	    // echo implode(' ', ['(', $curr, '/', $begin, ')', '-1', '=', $sheets[$index]['data'][$row][4]]), PHP_EOL;
	    $dayTwr = $sheets[$index]['data'][$row][4] + 1;
	    //multiple total
	    $sheets[$index]['total'] *= $dayTwr;

	    $week = date('W', $sheets[$index]['data'][$row][1]);
	    if (!isset($sheets[$index]['weekly'][$week])) {
	    	$sheets[$index]['weekly'][$week] = 1;
	    }
		$sheets[$index]['weekly'][$week] *= $dayTwr;

		$month = date('M', $sheets[$index]['data'][$row][1]);
		if (!isset($sheets[$index]['monthly'][$month])) {
	    	$sheets[$index]['monthly'][$month] = 1;
	    }
		$sheets[$index]['monthly'][$month] *= $dayTwr;

		$quarter = intval((date('n', $sheets[$index]['data'][$row][1])+2)/3);
		if (!isset($sheets[$index]['quarter'][$quarter])) {
	    	$sheets[$index]['quarter'][$quarter] = 1;
	    }
		$sheets[$index]['quarter'][$quarter] *= $dayTwr;
	}

	foreach ($sheets[$index]['data'] as $row) {
		if (!isset($row[4])) {
			continue;
		}
	}

	echo PHP_EOL, $sheetName, PHP_EOL;
	echo '===========', PHP_EOL;
	echo 'Weekly TWR:', PHP_EOL;
	foreach ($sheets[$index]['weekly'] as $week => $result) {
		echo $week, ' ', (round($result - 1, 4) * 100) . '%', PHP_EOL;
	}
	echo '===========', PHP_EOL;
	echo 'Monthly TWR:', PHP_EOL;
	foreach ($sheets[$index]['monthly'] as $month => $result) {
		echo $month, ' ', (round($result - 1, 4) * 100) . '%', PHP_EOL;
	}
	echo '===========', PHP_EOL;
	echo 'Quarter TWR:', PHP_EOL;
	foreach ($sheets[$index]['quarter'] as $quarter => $result) {
		echo $quarter, ' ', (round($result - 1, 4) * 100) . '%', PHP_EOL;
	}
	echo '===========', PHP_EOL;
	echo 'Half-year TWR: ', (round($sheets[$index]['total'] - 1, 4) * 100) . '%', PHP_EOL, PHP_EOL;
}
//print_r($sheets);


