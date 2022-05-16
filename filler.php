<?php
require __DIR__ . '/vendor/autoload.php';

use mikehaertl\pdftk\Pdf;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Shared\Date;
use PhpOffice\PhpSpreadsheet\Shared\Font;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;

// Path to true-type fonts containig Arial
// Filename must be exactly: arial.ttf
Font::setTrueTypeFontPath('/usr/share/fonts/TTF/');
Font::setAutoSizeMethod(Font::AUTOSIZE_METHOD_EXACT);

// @TODO: Make month to be an ARG input as well
$month = readline(lang('MONTH_INPUT'));
if(!is_numeric($month) || $month < 1 || $month > 12) die(lang('MONTH_INVALID'));

$runPDF = readline(lang('PDF_PROCESS'));
$runPDF = empty($runPDF) ? 0 : (!is_numeric($runPDF) || $runPDF > 1 || $runPDF < 0 ? die(lang('PDF_INVALID')) : $runPDF);

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
    'P' => [0, 0, 0, 0, 0, 0, 'Publishers'],
    'R' => [0, 0, 0, 0, 0, 0, 'Regular Pioneers'],
    'A' => [0, 0, 0, 0, 0, 0, 'Auxiliary Pioneers']
];
$monthName = $months[$month];
$serviceYear = 2022; // @TODO: Make year to be an arg and user input
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

if ($handle = opendir($directory)) {
    $spreadsheet = new Spreadsheet();
    $spreadsheet->getDefaultStyle()->getFont()->setName('Arial');
    $spreadsheet->getDefaultStyle()->getFont()->setSize(11);
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
        if (preg_match('/\.pdf$/', $fileName) && isset($report)) {
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

            if($runPDF) {
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

    setSizeAndColors($publisherSheet);

    array_pop($columns);

    $spreadsheet->createSheet();
    $spreadsheet->setActiveSheetIndex(1);

    $publisherTotalSheet = $spreadsheet->getActiveSheet()->freezePane('A2');
    $publisherTotalSheet->fromArray(array_merge([''], array_values($columns), ['Number of Reports']), null, "A1");
    $publisherTotalSheet->setTitle(lang('TAB_TOTALS'));

    $index = 2;
    foreach($segments as $privilege => $data) {
        if($runPDF) {
            $fill = [];
            foreach(array_keys($columns) as $i => $column) {
                $fill["{$prefix}-{$column}_{$month}"] = intval($data[$i]);
            }
            $fill["{$indexColumns[5]}{$monthName}{$suffix}"] = intval($data[5]);

            $file = sprintf("%s/../%s/%s.pdf", $directory, lang('FOLDER_TOTALS'), $data[6]);
            $pdf = new Pdf($file);
            $pdf->fillForm($fill);
            if (!$pdf->saveAs($file)) {
                die($pdf->getError());
            }

            print $data[6] . PHP_EOL;
        }
        unset($data[6]);
        $publisherTotalSheet->fromArray(array_merge([$privilege], $data), null, "A{$index}");
        $index++;
    }

    setSizeAndColors($publisherTotalSheet);

    $spreadsheet->createSheet();
    $spreadsheet->setActiveSheetIndex(2);
    $attendenceSheet = $spreadsheet->getActiveSheet()->freezePane('A2');
    $attendenceSheet->fromArray(['', 'Total', 'Foreigners Only'], null, "A1");
    $attendenceSheet->setTitle(lang('TAB_ATTENDENCE'));

    // 0: Number of meetings
    // 1: Attendance (Total)
    // 2: Attendance (Foreigners Only)
    $weekend = [0, 0, 0];
    $midweek = [0, 0, 0];

    $index = 2;
    foreach($meetings as $meeting) {
        $date = DateTime::createFromFormat('Y-m-d', $meeting[0]);

        $isWeekend = isWeekend($date);
        if(!is_null($isWeekend)) {
            if($isWeekend) {
                $weekend[0]++;
                $weekend[1]+=$meeting[1];
                $weekend[2]+=$meeting[2];
            }
            else {
                $midweek[0]++;
                $midweek[1]+=$meeting[1];
                $midweek[2]+=$meeting[2];
            }
        }

        $attendenceSheet->fromArray($meeting, null, "A{$index}");
        $index++;
    }

    if($runPDF) {
        $_prefix = $prefix + 2;
        $attendance = [
            'Report of Meeting Attendance - Foreigners' => [
                "{$prefix}-Meeting_{$month}" => $midweek[0],
                "{$prefix}-Attendance_{$month}" => $midweek[2],
                "{$prefix}-Average_{$month}" => round($midweek[2] / $midweek[0], 2),
                "{$_prefix}-Meeting_{$month}" => $weekend[0],
                "{$_prefix}-Attendance_{$month}" => $weekend[2],
                "{$_prefix}-Average_{$month}" => round($weekend[2] / $weekend[0], 2)
            ],
            'Report of Meeting Attendance' => [
                "{$prefix}-Meeting_{$month}" => $midweek[0],
                "{$prefix}-Attendance_{$month}" => $midweek[1],
                "{$prefix}-Average_{$month}" => round($midweek[1] / $midweek[0], 2),
                "{$_prefix}-Meeting_{$month}" => $weekend[0],
                "{$_prefix}-Attendance_{$month}" => $weekend[1],
                "{$_prefix}-Average_{$month}" => round($weekend[1] / $weekend[0], 2)
            ]
        ];

        foreach($attendance as $file => $reports) {
            $path = sprintf("%s/../%s/%s.pdf", $directory, 'Meeting Attendence', $file);
            $pdf = new Pdf($path);
            $pdf->fillForm($reports);
            if (!$pdf->saveAs($path)) {
                die($pdf->getError());
            }
        }
    }

    setSizeAndColors($attendenceSheet);

    $writer = IOFactory::createWriter($spreadsheet, 'Xls');
    $writer->save(getcwd() . "/excel/{$serviceYear}-{$month}.xlsx");
}

function setSizeAndColors($sheet) {
    foreach ($sheet->getRowIterator() as $row) {
        $rowId = $row->getRowIndex();
        foreach($row->getCellIterator() as $column) {
            $columnId = $column->getColumn();
            $cell = "{$columnId}{$rowId}";
            $value = $sheet->getCell("{$cell}")->getValue();
            if($columnId == "A") {
                if($value == null) {
                    continue;
                }
                if(preg_match("/\b(A|P|R)\b/", $value)) {
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
                    }
                }
                else if(preg_match("/\d{4}\-\d{2}-\d{2}/", $value)) {
                    $date = DateTime::createFromFormat('Y-m-d', $value);
                    $format = true ? 'NNNNMMMM DD, YYYY' : NumberFormat::FORMAT_DATE_YYYYMMDD;

                    $sheet->setCellValue($cell, Date::PHPToExcel($date));
                    $sheet->getStyle($cell)->getNumberFormat()->setFormatCode($format);

                    $isWeekend = isWeekend($date);
                    $bg = is_null($isWeekend) ? Color::COLOR_WHITE : $isWeekend ? Color::COLOR_YELLOW : Color::COLOR_RED;
                }
            }
            $sheet->getStyle("{$cell}")->applyFromArray(getStyle($rowId == 1 ? 1 : $value, $bg ?? Color::COLOR_WHITE));
            $sheet->getColumnDimension($columnId)->setAutoSize(true);
        }
    }
}

function isWeekend($date) {
    switch ($date->format('l')) {
        case "Sunday":
        case "Saturday":
            return true;
            break;
        case "Monday":
        case "Tuesday":
        case "Wednesday":
        case "Thursday":
        case "Friday":
            return false;
            break;
    }
    return null;
}

function getStyle($value, $bg) {
    return [
        'font' => [
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
        'FOLDER_TOTALS'        => 'Congregation Totals',
        'PDF_PROCESS'          => 'Process PDF [0 = No, 1 = Yes, default = 0]: ',
        'PDF_INVALID'          => 'Invalid PDF param',
        'TAB_TOTALS'           => 'Publisher Recordings Totals',
        'TAB_ATTENDENCE'       => 'Meeting Attendence'
    ];
    return $lang[$phrase];
}
