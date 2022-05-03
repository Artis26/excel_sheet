<?php
require_once ('vendor/autoload.php');

use App\Models\Sheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$newSheet = new Sheet();
$newSheet = $newSheet->generateNew(2022, 6);

$writer = new xlsx($newSheet);
$writer->save("public/demo.xlsx");

echo 'FInd in public directory' . PHP_EOL;
