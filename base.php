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

Font::setAutoSizeMethod(Font::AUTOSIZE_METHOD_APPROX);

// @TODO: Make month to be an ARG input as well
$month = readline(lang('MONTH_INPUT'));
if(!is_numeric($month) || $month < 1 || $month > 12) die(lang('MONTH_INVALID'));

$runPDF = readline(lang('PDF_PROCESS'));
$runPDF = empty($runPDF) ? 0 : (!is_numeric($runPDF) || $runPDF > 1 || $runPDF < 0 ? die(lang('PDF_INVALID')) : $runPDF);

$serviceYear = 2022; // @TODO: Make year to be an arg and user input
$directory = sprintf("%s/pdf", getcwd());
$prefix = 1;
$suffix = $prefix == 1 ? '' : '_2';

$spreadsheet = new Spreadsheet();
$spreadsheet->getDefaultStyle()->getFont()->setName('Arial');
$spreadsheet->getDefaultStyle()->getFont()->setSize(11);
$spreadsheet->setActiveSheetIndex(0);

function savePDF($file, $data) {
    $pdf = new Pdf($file);
    $pdf->fillForm($data);
    if (!$pdf->saveAs($file)) {
        die($pdf->getError());
    }
    print $file . PHP_EOL;
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
                    $bg = is_null($isWeekend) ? Color::COLOR_WHITE : ($isWeekend ? Color::COLOR_YELLOW : Color::COLOR_RED);
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
        'TAB_ATTENDENCE'       => 'Meeting Attendence',
        'TAB_ATTENDENCE_TOTAL' => 'Meeting Attendence Totals'
    ];
    return $lang[$phrase];
}
