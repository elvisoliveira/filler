<?php
require __DIR__ . '/vendor/autoload.php';

use mikehaertl\pdftk\Pdf;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Shared\Date;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;

$month = readline(lang('MONTH_INPUT'));
if(!is_numeric($month) || $month < 1 || $month > 12) die(lang('MONTH_INVALID'));

$processPDF = readline(lang('PDF_PROCESS'));
$processPDF = empty($processPDF) ? 0 : (!is_numeric($processPDF) || $processPDF > 1 || $processPDF < 0 ? die(lang('PDF_INVALID')) : $processPDF);

$months = [
    1 => 'September',
    2 => 'October',
    3 => 'November',
    4 => 'December',
    5 => 'January',
    6 => 'February',
    7 => 'March',
    8 => 'April',
    9 => 'May',
    10 => 'June',
    11 => 'July',
    12 => 'August'
);
$columns = array(
    "Place",
    "Video",
    "Hours",
    "RV",
    "Studies",
    "Remarks"
);
$monthName = $months[$month];
$serviceYear = 2022;
$directory = sprintf("%s/pdf/%s", getcwd(), lang('FOLDER_PUBLISHER'));
$prefix = 1;
$suffix = $prefix > 1 ? "_{$prefix}" : '';

// Placements
// Video Showings
// Hours
// Return Visits
// Bible Studies
// Observation
$reports = [];
$reportsFile = sprintf("%s/reports/%s-%s.csv", getcwd(), $serviceYear, $month);
if (!file_exists($reportsFile)) {
    die(lang('NOT_FOUND_REPORT'));
}
if (($handle = fopen($reportsFile, 'r')) !== FALSE) {
    while (($data = fgetcsv($handle, 1000, ',')) !== FALSE) {
        $name = trim($data[0]);
        if(empty($name)) {
            continue;
        }
        unset($data[0]);
        $reports["{$name}.pdf"] = array_map('trim', array_values($data));
    }
}

// Assistência Reunião
$meetings = [];
$meetingsFile = sprintf("%s/attendence/%s-%s.csv", getcwd(), $serviceYear, $month);
if (!file_exists($meetingsFile)) {
    die(lang('NOT_FOUND_ATTENDENCE'));
}
if (($handle = fopen($meetingsFile, 'r')) !== FALSE) {
    while (($data = fgetcsv($handle, 1000, ',')) !== FALSE) {
        array_push($meetings, array_map('trim', $data));
    }
}

uasort($reports, function ($one, $two) {
    return $one[0] <=> $two[0];
});

if ($handle = opendir($directory))
{
    $spreadsheet = new Spreadsheet();
    $spreadsheet->setActiveSheetIndex(0);
    $publisherSheet = $spreadsheet->getActiveSheet();
    $publisherSheet->setTitle(lang('FOLDER_PUBLISHER'));

    $index = 1;
    foreach($reports as $fileName => $report) {
        if (preg_match('/\.pdf$/', $fileName) && isset($report)) {
            $data = [
                "Service Year{$suffix}"             => $serviceYear,
                "{$prefix}-{$columns[0]}_{$month}"   => intval($report[1]),
                "{$prefix}-{$columns[1]}_{$month}"   => intval($report[2]),
                "{$prefix}-{$columns[2]}_{$month}"   => intval($report[3]),
                "{$prefix}-{$columns[3]}_{$month}"   => intval($report[4]),
                "{$prefix}-{$columns[4]}_{$month}"   => intval($report[5]),
                "{$columns[5]}{$monthName}{$suffix}" => $report[6]
            ];
            $assignment = $report[0];
            $row = array_merge(array($assignment, pathinfo($fileName, PATHINFO_FILENAME)), array_values($data));

            unset($row[2]);

            $publisherSheet->fromArray($row, null, "A{$index}");
            $index++;

            if($processPDF)
            {
                $pdf = new Pdf("{$directory}/{$fileName}");
                $pdf->fillForm($data);
                if (!$pdf->saveAs("{$directory}/{$fileName}")) {
                    die($pdf->getError());
                }
            }

            print $fileName . PHP_EOL;
        }
    }
    closedir($handle);

    foreach ($publisherSheet->getRowIterator() as $row) {
        $rowId = $row->getRowIndex();
        foreach($row->getCellIterator() as $column) {
            $columnId = $column->getColumn();
            $value = $publisherSheet->getCell("{$columnId}{$rowId}")->getValue();
            if($columnId == "A") {
                switch ($value) {
                    case "P":
                        $bg = Color::COLOR_YELLOW;
                        break;
                    case "R":
                        $bg = Color::COLOR_RED;
                        break;
                    case "A":
                        $bg = Color::COLOR_GREEN;
                        break;
                    default:
                        $bg = Color::COLOR_WHITE;
                }
            }
            $publisherSheet->getStyle("{$columnId}{$rowId}")->applyFromArray(getStyle($value, $bg));
            $publisherSheet->getColumnDimension($columnId)->setAutoSize(true);
        }
    }

    $spreadsheet->createSheet();
    $spreadsheet->setActiveSheetIndex(1);
    $attendenceSheet = $spreadsheet->getActiveSheet();
    $attendenceSheet->setTitle('Meeting Attendence');

    $index = 1;
    foreach($meetings as $meeting) {
        $attendenceSheet->fromArray($meeting, null, "A{$index}");
        $index++;
    }

    foreach ($attendenceSheet->getRowIterator() as $row) {
        $rowId = $row->getRowIndex();
        foreach($row->getCellIterator() as $column) {
            $columnId = $column->getColumn();
            $cell = "{$columnId}{$rowId}";
            $cellValue = $attendenceSheet->getCell($cell)->getValue();

            if($columnId == "A") {
                $date = DateTime::createFromFormat('Y-m-d', $cellValue);

                $attendenceSheet->setCellValue($cell, Date::PHPToExcel($date));
                $attendenceSheet->getStyle($cell)->getNumberFormat()->setFormatCode(true ? 'NNNNMMMM DD, YYYY' : NumberFormat::FORMAT_DATE_YYYYMMDD);

                switch ($date->format('l')) {
                    case "Sunday":
                    case "Saturday":
                        $bg = Color::COLOR_YELLOW;
                        break;
                    case "Monday":
                    case "Tuesday":
                    case "Wednesday":
                    case "Thursday":
                    case "Friday":
                        $bg = Color::COLOR_RED;
                        break;
                    default:
                        $bg = Color::COLOR_WHITE;
                }
            }

            $attendenceSheet->getStyle($cell)->applyFromArray(getStyle($cellValue, $bg));
            $attendenceSheet->getColumnDimension($columnId)->setAutoSize(true);
        }
    }

    $writer = IOFactory::createWriter($spreadsheet, 'Xls');
    $writer->save(getcwd() . "/excel/{$serviceYear}-{$month}.xlsx");
}

function getStyle($value, $bg) {
    return [
        'font' => [
            'size'  => 11,
            'name'  =>  'Arial',
            'color' => [
                'argb' => Color::COLOR_BLACK
            ]
        ],
        'alignment' => [
            'horizontal' => is_numeric($value) ? Alignment::HORIZONTAL_CENTER : Alignment::HORIZONTAL_LEFT
        ],
        'borders' => [
            'outline' => [
                'borderStyle' => Border::BORDER_THIN,
                'color' => [
                    'argb' => Color::COLOR_BLACK
                ]
            ]
        ],
        'fill' => [
            'fillType' => Fill::FILL_SOLID,
            'startColor' => [
                'argb' => $bg
            ]
        ]
    ];
}

function lang($phrase) {
    static $lang = [
        'NOT_FOUND_REPORT'     => 'Reports file not found',
        'NOT_FOUND_ATTENDENCE' => 'Attendence file not found',
        'MONTH_INPUT'          => "Input the service year month number [1-12]: ",
        'MONTH_INVALID'        => "Invalid month",
        'FOLDER_PUBLISHER'     => "Publisher Recordings",
        'PDF_PROCESS'          => "Process PDF [0 = No, 1 = Yes, default = 0]: ",
        'PDF_INVALID'          => "Invalid PDF param"
    ];
    return $lang[$phrase];
}
