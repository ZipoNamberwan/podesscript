<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class AnomaliController extends Controller
{
    function generate()
    {
        $anomali = \PhpOffice\PhpSpreadsheet\IOFactory::load('anomali/Hasil Anomali Untuk Script.xlsx');
        $sourcesheet = $anomali->getActiveSheet();

        $spreadsheet =

            //Tabel 1.1.2
            $target = \PhpOffice\PhpSpreadsheet\IOFactory::load('anomali/anomali template.xlsx');
        $targetsheet = $spreadsheet->getSheet(0);

        $row_target = 2;
        for ($i = 10001; $i < 23013; $i++) {
            $anomalies = explode(";", $sourcesheet->getCell('K' . $i)->getValue());
            foreach ($anomalies as $a) {
                if ($a != "" | $a != null) {
                    $targetsheet->setCellValue('A' . $row_target, $sourcesheet->getCell('A' . $i)->getValue());
                    $targetsheet->setCellValue('B' . $row_target, $sourcesheet->getCell('B' . $i)->getValue());
                    $targetsheet->setCellValue('C' . $row_target, $sourcesheet->getCell('C' . $i)->getValue());
                    $targetsheet->setCellValue('D' . $row_target, '=CONCATENATE("3513",A' . $row_target . ',B' . $row_target . ',C' . $row_target . ',TEXT(F' . $row_target . ',REPT("0",3)))');
                    $targetsheet->setCellValue('E' . $row_target, $sourcesheet->getCell('E' . $i)->getValue());
                    $targetsheet->setCellValue('F' . $row_target, $sourcesheet->getCell('F' . $i)->getValue());
                    $targetsheet->setCellValue('G' . $row_target, $sourcesheet->getCell('G' . $i)->getValue());
                    $targetsheet->setCellValue('H' . $row_target, $sourcesheet->getCell('H' . $i)->getValue());
                    $targetsheet->setCellValue('I' . $row_target, $sourcesheet->getCell('I' . $i)->getValue());
                    $targetsheet->setCellValue('J' . $row_target, $sourcesheet->getCell('J' . $i)->getValue());
                    $targetsheet->setCellValue('K' . $row_target, $sourcesheet->getCell('K' . $i)->getValue());
                    $targetsheet->setCellValue('K' . $row_target, $a);
                    $targetsheet->setCellValue('L' . $row_target, '=VLOOKUP(K' . $row_target . ',Sheet2!$A$1:$B$105,2,FALSE)');
                    $row_target++;
                }
            }
        }

        $writer = new Xlsx($target);

        $writer->save("anomali/anomali result.xlsx");

        return 'done';
    }
}
