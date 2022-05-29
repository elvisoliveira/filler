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

$pdf = 'S-21_E';
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

// Reports
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

uasort($reports, function ($one, $two) {
    return $one[0] <=> $two[0];
});

$publisherSheet = $spreadsheet->getActiveSheet()->freezePane('A2');
$publisherSheet->setTitle(lang('FOLDER_PUBLISHER'));
$publisherSheet->fromArray(array_merge(['', ''], array_values($columns)), null, "A1");

$index = 2;
foreach($reports as $file => $report) {
    // Calc totals
    for ($i = 0; $i <= 4; $i++) {
        $segments[$report[0]][$i] += $report[$i + 1];
    }
    $segments[$report[0]][5]++;
    // PDF
    if ($runPDF) {
        // $data["Service Year{$suffix}"] = $serviceYear;
        $indexColumns = array_keys($columns);
        for ($i = 0; $i <= 4; $i++) {
            $data["{$prefix}-{$indexColumns[$i]}_{$month}"] = intval($report[$i + 1]);
        }
        $data["{$indexColumns[5]}{$monthName}{$suffix}"] = $report[6];

        $file = sprintf("{$directory}/%s/{$file}", lang('FOLDER_PUBLISHER'));
        savePDF($file, $data);
        calcPDF($file);
        cleanPDF($file);
    }
    // XLS
    $assignment = $report[0];
    unset($report[0]);
    $values = array_merge([$assignment, pathinfo($file, PATHINFO_FILENAME)], array_values($report));
    $publisherSheet->fromArray($values, null, "A{$index}");
    $index++;
}

setSizeAndColors($publisherSheet);

// Remove Observation column for the next procedures
array_pop($columns);

$spreadsheet->createSheet();
$spreadsheet->setActiveSheetIndex(1);

$publisherTotalSheet = $spreadsheet->getActiveSheet()->freezePane('A2');
$publisherTotalSheet->fromArray(array_merge([''], array_values($columns), ['Number of Reports']), null, "A1");
$publisherTotalSheet->setTitle(lang('TAB_TOTALS'));

$index = 2;
foreach($segments as $privilege => $data) {
    // PDF
    if($runPDF) {
        $fill = [];
        foreach(array_keys($columns) as $i => $column) {
            $fill["{$prefix}-{$column}_{$month}"] = intval($data[$i]);
        }
        // Fill the amount of reports on Observation
        $fill["{$indexColumns[5]}{$monthName}{$suffix}"] = intval($data[5]);

        $file = sprintf("%s/%s/%s.pdf", $directory, lang('FOLDER_TOTALS'), $data[6]);

        savePDF($file, $fill);
        calcPDF($file);
        cleanPDF($file);
    }
    // Remove privilege labels
    unset($data[6]);
    $publisherTotalSheet->fromArray(array_merge([$privilege], $data), null, "A{$index}");
    $index++;
}

setSizeAndColors($publisherTotalSheet);

$writer = IOFactory::createWriter($spreadsheet, 'Xls');
$writer->save(getcwd() . "/excel/{$pdf}-{$serviceYear}-{$month}.xlsx");

function calcPDF($entry) {
    global $columns;
    global $prefix;
    global $suffix;

    $f = ['Remarks', 'Average', 'Total']; // meta fields
    $average = 0;
    $total = [];

    foreach((new Pdf($entry))->getDataFields() as $field) {
        $name = $field['FieldName'];
        $value = $field['FieldValue'] ?? 0;
        // Evaluate each month of each column
        foreach(array_keys($columns) as $column) {
            // Set initial value for each column calculated totals
            if(!isset($total[$column])) {
                $total[$column] = 0;
            }
            for ($i = 1; $i <= 12; $i++) {
                if($name == "{$prefix}-{$column}_{$i}" && is_numeric($value)) {
                    $total[$column] = $total[$column] + intval($value);
                    // Count each filled hour column to calc averages
                    if($column == 'Hours') {
                        $average++;
                    }
                }
            }
        }
        if(str_starts_with($name, $f[0]) && !str_contains($name, $f[1]) && !str_contains($name, $f[2])) {
            $endsWith = str_ends_with('_2', $name);
            if(empty($suffix) ? !$endsWith : $endsWith) {
                $int = (int) filter_var($value, FILTER_SANITIZE_NUMBER_INT);
                if(is_numeric($int)) {
                    $total[$f[0]] = $total[$f[0]] + $int;
                }
            }
        }
    }

    $data = [];
    foreach(array_keys($columns) as $column) {
        $valueTotal = intval($total[$column]);
        $valueAverage = round(($valueTotal / $average), 2);
        if($column == $f[0]) { // total and average of remarks column
            $data = array_merge($data,
                [$f[0] . $f[2] => $valueTotal],
                [$f[0] . $f[1] => $valueAverage]
            );
        }
        else { // total and average of each column (but remarks)
            $data = array_merge($data,
                ["{$prefix}-{$column}_" . $f[2] => $valueTotal],
                ["{$prefix}-{$column}_" . $f[1] => $valueAverage]
            );
        }
    }

    savePDF($entry, $data);
}

function cleanPDF($entry) {
    global $pdf;
    global $directory;

    $data = [];
    $birth = 0;
    $f = ['Date of birth',  'Date immersed']; // fields to calc

    foreach((new Pdf($entry))->getDataFields() as $field) {
        $name = $field['FieldName'];
        $value = $field['FieldValue'] ?? false;
        if($value) {
            if (preg_match('/^([0-9]{1,2})\\/([0-9]{1,2})\\/([0-9]{4})/', $value, $matches)) {
                $date = $matches[0];
                $elapsed = (new DateTime())->diff(DateTime::createFromFormat('d/m/Y', $date))->format('%yy');
                if($name == $f[0]) {
                    $birth = DateTime::createFromFormat('d/m/Y', $date);
                    $value = "{$date}; {$elapsed} of age";
                }
                if($name == $f[1]) {
                    $fromBirth  = ($birth)->diff(DateTime::createFromFormat('d/m/Y', $date))->format('%yy');
                    $value = "{$date}; {$elapsed} of baptism; baptized with {$fromBirth}";
                }
            }
            $data = array_merge($data, [$name => $value]);
        }
    }

    $file = new Pdf(sprintf("%s/%s", $directory, "{$pdf}.pdf"));
    $file->fillForm($data);
    if (!$file->saveAs($entry)) {
        die($file->getError());
    }

    print "Cleaned {$entry}" . PHP_EOL;
}
