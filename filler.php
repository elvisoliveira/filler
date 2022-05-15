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
];
$columns = [
    'Place' => 'Placements',
    'Video' => 'Video Showings',
    'Hours' => 'Hours',
    'RV' => 'Return Visits',
    'Studies' => 'Bible Studies',
    'Remarks' => 'Observation'
];
$segments = [
    'P' => [0, 0, 0, 0, 0, 0], // Publishers
    'R' => [0, 0, 0, 0, 0, 0], // Regular Pioneers
    'A' => [0, 0, 0, 0, 0, 0]  // Auxiliary Pioneers
];
$monthName = $months[$month];
$serviceYear = 2022;
$directory = sprintf("%s/pdf/%s", getcwd(), lang('FOLDER_PUBLISHER'));
$prefix = 1;
$suffix = $prefix > 1 ? "_{$prefix}" : '';

$reportsFile = sprintf("%s/reports/%s-%s.csv", getcwd(), $serviceYear, $month);
if (!file_exists($reportsFile)) {
    die(lang('NOT_FOUND_REPORT'));
}
$reports = [];
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

// Meeting Attendance
$meetingsFile = sprintf("%s/attendence/%s-%s.csv", getcwd(), $serviceYear, $month);
if (!file_exists($meetingsFile)) {
    die(lang('NOT_FOUND_ATTENDENCE'));
}
$meetings = [];
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
    $publisherSheet = $spreadsheet->getActiveSheet()->freezePane('A2');
    $publisherSheet->setTitle(lang('FOLDER_PUBLISHER'));
    $publisherSheet->fromArray(array_merge(['', ''], array_values($columns)), null, "A1");

    $index = 2;
    foreach($reports as $fileName => $report) {
        for ($i = 0; $i <= 4; $i++) {
            $segments[$report[0]][$i] += $report[$i + 1];
        }
        $segments[$report[0]][5]++;
        if (preg_match('/\.pdf$/', $fileName) && isset($report))
        {
            $data["Service Year{$suffix}"] = $serviceYear;
            $indexColumns = array_keys($columns);
            for ($i = 0; $i <= 4; $i++) {
                $data["{$prefix}-{$indexColumns[$i]}_{$month}"] = intval($report[$i + 1]);
            }
            $data["{$indexColumns[5]}{$monthName}{$suffix}"] = $report[6];

            $assignment = $report[0];
            $row = array_merge([$assignment, pathinfo($fileName, PATHINFO_FILENAME)], array_values($data));

            unset($row[2]);

            $publisherSheet->fromArray($row, null, "A{$index}");
            $index++;

            if($processPDF) {
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

    setAutoSize($publisherSheet);

    array_pop($columns);

    $spreadsheet->createSheet();
    $spreadsheet->setActiveSheetIndex(1);
    $publisherTotalSheet = $spreadsheet->getActiveSheet()->freezePane('A2');
    $publisherTotalSheet->fromArray(array_merge([''], array_values($columns), ['Number of Reports']), null, "A1");
    $publisherTotalSheet->setTitle(lang('TAB_TOTALS'));
    $index = 2;
    foreach($segments as $privilege => $data) {
        $publisherTotalSheet->fromArray(array_merge([$privilege], $data), null, "A{$index}");
        $index++;
    }

    setAutoSize($publisherTotalSheet);

    $spreadsheet->createSheet();
    $spreadsheet->setActiveSheetIndex(2);
    $attendenceSheet = $spreadsheet->getActiveSheet()->freezePane('A2');
    $attendenceSheet->fromArray(['', 'Total', 'Foreigners Only'], null, "A1");
    $attendenceSheet->setTitle(lang('TAB_ATTENDENCE'));

    $index = 2;
    foreach($meetings as $meeting) {
        $attendenceSheet->fromArray($meeting, null, "A{$index}");
        $index++;
    }

    foreach ($attendenceSheet->getRowIterator() as $row) {
        $rowId = $row->getRowIndex();
        foreach($row->getCellIterator() as $column) {
            $bg = Color::COLOR_WHITE;
            $columnId = $column->getColumn();
            $cell = "{$columnId}{$rowId}";
            $cellValue = $attendenceSheet->getCell($cell)->getValue();

            if($columnId == "A" && $cellValue <> NULL) {
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
                }
            }

            $attendenceSheet->getStyle($cell)->applyFromArray(getStyle($cellValue, $bg));
            $attendenceSheet->getColumnDimension($columnId)->setAutoSize(true);
        }
    }

    $writer = IOFactory::createWriter($spreadsheet, 'Xls');
    $writer->save(getcwd() . "/excel/{$serviceYear}-{$month}.xlsx");
}

function setAutoSize($sheet) {
    foreach ($sheet->getRowIterator() as $row) {
        $rowId = $row->getRowIndex();
        foreach($row->getCellIterator() as $column) {
            $columnId = $column->getColumn();
            $value = $sheet->getCell("{$columnId}{$rowId}")->getValue();
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
            $sheet->getStyle("{$columnId}{$rowId}")->applyFromArray(getStyle($value, $bg));
            $sheet->getColumnDimension($columnId)->setAutoSize(true);
        }
    }
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
        'MONTH_INPUT'          => 'Input the service year month number [1-12]: ',
        'MONTH_INVALID'        => 'Invalid month',
        'FOLDER_PUBLISHER'     => 'Publisher Recordings',
        'PDF_PROCESS'          => 'Process PDF [0 = No, 1 = Yes, default = 0]: ',
        'PDF_INVALID'          => 'Invalid PDF param',
        'TAB_TOTALS'           => 'Publisher Recordings Totals',
        'TAB_ATTENDENCE'       => 'Meeting Attendence'
    ];
    return $lang[$phrase];
}
