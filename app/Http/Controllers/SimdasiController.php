<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Illuminate\Support\Str;

class SimdasiController extends Controller
{
    public function transform()
    {
        $table = array(
            '5.4.3' => array(
                'D' => '2',
                'F' => '3',
                'AR' => '4',
                'L' => '5',
                'T' => '6',
                'X' => '7',
                'AN' => '8',
                'B' => '9',
                'AF' => '10',
                'AB' => '13',
                'AH' => '14',
                'AJ' => '15',
            ),
            '5.4.4' => array(
                'D' => '2',
                'F' => '3',
                'AR' => '4',
                'L' => '5',
                'T' => '6',
                'X' => '7',
                'AN' => '8',
                'B' => '9',
                'AF' => '10',
                'AB' => '13',
                'AH' => '14',
                'AJ' => '15',
            ),
            '5.4.5' => array(
                'B' => '1',
                'H' => '2',
                'J' => '3',
                'L' => '4',
                'X' => '5',
                'Z' => '6',
                'R' => '7',
                'AB' => '8',
            ),
            '5.4.6' => array(
                'B' => '1',
                'H' => '2',
                'J' => '3',
                'L' => '4',
                'X' => '5',
                'Z' => '6',
                'R' => '7',
                'AB' => '8',
            ),
            '5.4.7' => array(
                'AV' => '1',
                'J' => '2',
                'P' => '3',
                'X' => '4',
                'Z' => '5',
                'AB' => '6',
                'AF' => '7',
                'AH' => '8',
                'AJ' => '9',
                'AL' => '10',
                'AT' => '11',
            ),
            '5.4.8' => array(
                'AV' => '1',
                'J' => '2',
                'P' => '3',
                'X' => '4',
                'Z' => '5',
                'AB' => '6',
                'AF' => '7',
                'AH' => '8',
                'AJ' => '9',
                'AL' => '10',
                'AT' => '11',
            ),
            '5.4.9' => array(
                'B' => '2',
                'D' => '3',
                'H' => '4',
                'J' => '5',
                'L' => '6',
                'N' => '7',
                'P' => '8',
                'V' => '9',
                'X' => '10',
                'Z' => '11',
                'AB' => '12',
                'AD' => '13',
                'AH' => '14',
                'AL' => '15',
                'AN' => '16',
                'AP' => '17',
                'AR' => '18',
                'AT' => '19',
                'AJ' => '22',
            ),
            '5.4.10' => array(
                'B' => '3',
                'D' => '4',
                'F' => '5',
                'H' => '6',
                'J' => '7',
                'L' => '8',
                'N' => '12',
                'P' => '13',
            ),
            '5.4.11' => array(
                'B' => '3',
                'D' => '4',
                'F' => '5',
                'H' => '6',
                'J' => '7',
                'L' => '8',
                'N' => '12',
                'P' => '13',
            ),
            '5.4.12' => array(
                'B' => '1',
                'D' => '3',
                'F' => '4',
                'H' => '5',
            ),
            '5.4.13' => array(
                'B' => '1',
                'D' => '2',
                'F' => '3',
                'H' => '4',
                'J' => '5',
                'L' => '6',
                'N' => '7',
                'P' => '8',
            ),
            '5.4.14' => array(
                'B' => '1',
                'D' => '2',
                'F' => '3',
                'H' => '4',
                'J' => '5',
                'L' => '6',
                'N' => '7',
            ),
            '5.4.15' => array(
                'B' => '1',
                'D' => '2',
                'F' => '3',
                'H' => '4',
                'J' => '5',
                'L' => '6',
                'N' => '7',
            ),
            '5.4.16' => array(
                'B' => '2',
                'D' => '3',
                'F' => '5',
                'H' => '6',
                'J' => '7',
                'L' => '8',
                'N' => '9',
                'P' => '10',
                'R' => '11',
            ),
            '5.4.1' => array(
                'B' => '1',
                'D' => '2',
                'F' => '3',
                'H' => '4',
                'J' => '5',
                'L' => '6',
                'N' => '7',
                'O' => '8',
                'R' => '9',
            ),
            '5.4.2' => array(
                'B' => '1',
                'D' => '2',
                'F' => '3',
                'H' => '4',
                'J' => '5',
                'L' => '6',
                'N' => '7',
                'O' => '8',
                'R' => '9',
            ),
        );
        $kecamatan = [
            '10' => 'Sukapura',
            '20' => 'Sumber',
            '30' => 'Kuripan',
            '40' => 'Bantaran',
            '50' => 'Leces',
            '60' => 'Tegalsiwalan',
            '70' => 'Banyuanyar',
            '80' => 'Tiris',
            '90' => 'Krucil',
            '100' => 'Gading',
            '110' => 'Pakuniran',
            '120' => 'Kotaanyar',
            '130' => 'Paiton',
            '140' => 'Besuk',
            '150' => 'Kraksaan',
            '160' => 'Krejengan',
            '170' => 'Pajarakan',
            '180' => 'Maron',
            '190' => 'Gending',
            '200' => 'Dringu',
            '210' => 'Wonomerto',
            '220' => 'Lumbang',
            '230' => 'Tongas',
            '240' => 'Sumberasih'
        ];


        foreach ($kecamatan as $key => $namakec) {

            $target = \PhpOffice\PhpSpreadsheet\IOFactory::load('data/simdasi/template simdasi.xlsx');

            foreach ($table as $name => $transform) {
                $source543 = \PhpOffice\PhpSpreadsheet\IOFactory::load('data/simdasi/source/' . $name . '.xlsx');

                $sheetSource543 = $source543->getSheet(0);
                $sheetTarget = $target->getSheetByName($name);
                $title = $sheetTarget->getCell('A1')->getValue();
                $title = Str::replaceFirst('xxx', $namakec, $title);
                $sheetTarget->setCellValue('A1', $title);

                $startTargetRow = 4;
                $sourceRow = 0;
                for ($i = 9; $i < 33; $i++) {
                    if ($sheetSource543->getCell('A' . $i)->getValue() == $namakec) {
                        $sourceRow = $i;
                        break;
                    }
                }
                foreach ($transform as $sourceCol => $targetRow) {
                    $startTargetCol = 'B';
                    for ($j = 0; $j < 2; $j++) {
                        $value = $sheetSource543->getCell($sourceCol . $sourceRow)->getValue();
                        $sheetTarget->setCellValue($startTargetCol . ($targetRow + $startTargetRow), $value);
                        $sourceCol++;
                        $startTargetCol++;
                    }
                }
            }

            $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($target);
            $writer->save("data/simdasi/result/" . $namakec . ".xlsx");
        }

        return 'done';
    }
}
