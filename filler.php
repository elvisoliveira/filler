<?php

require __DIR__ . '/vendor/autoload.php';

use mikehaertl\pdftk\Pdf;

$reports = array(
	"BBB.pdf" => array("00", "00", "00", "00", "00",""),
	"BBB.pdf" => array("00", "00", "00", "00", "00","")
);
$month = ;
$monthName = "";
$columns = array("Place", "Video", "Hours", "RV", "Studies");
if ($handle = opendir('./')) {
	while (false !== ($entry = readdir($handle))) {
		if (preg_match('/\.pdf$/', $entry) && isset($reports[$entry])) {
			$pdf = new Pdf($entry);
			$pdf->fillForm([
				"1-Place_{$month}"     => intval($reports[$entry][0]),
				"1-Video_{$month}"     => intval($reports[$entry][1]),
				"1-Hours_{$month}"     => intval($reports[$entry][2]),
				"1-RV_{$month}"        => intval($reports[$entry][3]),
				"1-Studies_{$month}"   => intval($reports[$entry][4]),
				"Remarks{$monthName}"  => $reports[$entry][5]
			]);
			if (!$pdf->saveAs("./{$entry}")) {
				die($pdf->getError());
			}
			var_dump($entry);
		}
	}
	closedir($handle);
}
