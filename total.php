<?php

require __DIR__ . '/vendor/autoload.php';

use mikehaertl\pdftk\Pdf;

$reports = array(
	"AAA.pdf" => array("00", "00", "00", "00", "00",""),
	"BBB.pdf" => array("00", "00", "00", "00", "00","")
);
$columns = array("Place", "Video", "Hours", "RV", "Studies");
$average = 1;
if ($handle = opendir('./')) {
	while (false !== ($entry = readdir($handle))) {
		if (preg_match('/\.pdf$/', $entry) && isset($reports[$entry])) {
			$total = array();
			$pdfReader = new Pdf($entry);
			$fields = $pdfReader->getDataFields();
			foreach($fields as $field) {
				foreach($columns as $column) {
					if(!isset($total[$column])) {
						$total[$column] = 0;
					}
					for ($i = 1; $i <= 12; $i++) {
						if($field['FieldName'] == "1-{$column}_{$i}"
						&& isset($field['FieldValue'])
						&& is_numeric($field['FieldValue'])) {
							$total[$column] = $total[$column] + intval($field['FieldValue']);
						}
					}
				}
			}
			$filler = array();
			foreach($columns as $column) {
				$valueTotal = intval($total[$column]);
				$valueAverage = $valueTotal / $average;
				$filler = array_merge(
					$filler,
					array("1-{$column}_Total" => $valueTotal),
					array("1-{$column}_Average" => round($valueAverage, 2))
				);
			}
			$pdfFiller = new Pdf($entry);
			$pdfFiller->fillForm($filler);
			if (!$pdfFiller->saveAs("./{$entry}")) {
				die($pdfFiller->getError());
			}
			var_dump($entry);
			unset($total);
		}
	}
	closedir($handle);
}
