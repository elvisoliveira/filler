<?php
require __DIR__ . '/vendor/autoload.php';
require __DIR__ . '/base.php';

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

$pdf = 'S-88_E';
$columns = [
    '',
    'Mid. Amount',
    'Mid. Total',
    'Mid. Average',
    'Week. Amount',
    'Week. Total',
    'Week. Average'
];

$fields = [
    'Meeting',
    'Attendance',
    'Average'
];

// Meeting Attendance
$meetingsFile = sprintf("%s/attendence/%s-%s.csv", getcwd(), $serviceYear, str_pad($month, 2, '0', STR_PAD_LEFT));
if (!file_exists($meetingsFile)) {
    die(lang('NOT_FOUND_ATTENDENCE'));
}
$meetings = [];
if (($handle = fopen($meetingsFile, 'r')) !== FALSE) {
    while (($data = fgetcsv($handle, 1000, ',')) !== FALSE) {
        array_push($meetings, array_map('trim', $data));
    }
}

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
    // Calc the totals
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

    // XLS
    $attendenceSheet->fromArray($meeting, null, "A{$index}");
    $index++;
}

setSizeAndColors($attendenceSheet);

$_prefix = $prefix + 2;
$attendance = [
    'Report of Meeting Attendance - Foreigners' => [
        "{$prefix}-{$fields[0]}_{$month}" => $midweek[0],
        "{$prefix}-{$fields[1]}_{$month}" => $midweek[2],
        "{$prefix}-{$fields[2]}_{$month}" => round($midweek[2] / $midweek[0], 2),
        "{$_prefix}-{$fields[0]}_{$month}" => $weekend[0],
        "{$_prefix}-{$fields[1]}_{$month}" => $weekend[2],
        "{$_prefix}-{$fields[2]}_{$month}" => round($weekend[2] / $weekend[0], 2)
    ],
    'Report of Meeting Attendance' => [
        "{$prefix}-{$fields[0]}_{$month}" => $midweek[0],
        "{$prefix}-{$fields[1]}_{$month}" => $midweek[1],
        "{$prefix}-{$fields[2]}_{$month}" => round($midweek[1] / $midweek[0], 2),
        "{$_prefix}-{$fields[0]}_{$month}" => $weekend[0],
        "{$_prefix}-{$fields[1]}_{$month}" => $weekend[1],
        "{$_prefix}-{$fields[2]}_{$month}" => round($weekend[1] / $weekend[0], 2)
    ]
];

$spreadsheet->createSheet();
$spreadsheet->setActiveSheetIndex(1);
$attendenceTotalsSheet = $spreadsheet->getActiveSheet()->freezePane('A2');
$attendenceTotalsSheet->fromArray($columns, null, "A1");
$attendenceTotalsSheet->setTitle(lang('TAB_ATTENDENCE_TOTAL'));

$index = 2;
foreach($attendance as $file => $reports) {
    // PDF
    if($runPDF) {
        savePDF(sprintf("%s/%s/%s.pdf", $directory, 'Meeting Attendence', $file), $reports);
    }
    // XLS
    $attendenceTotalsSheet->fromArray(array_merge([$file], array_values($reports)), null, "A{$index}");
    $index++;
}

setSizeAndColors($attendenceTotalsSheet);

$writer = IOFactory::createWriter($spreadsheet, 'Xls');
$writer->save(getcwd() . "/excel/{$pdf}-{$serviceYear}-{$month}.xlsx");
