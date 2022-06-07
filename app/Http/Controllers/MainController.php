<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use App\Models\Podes1;
use App\Models\Podes2;
use App\Models\Podes3;
use App\Models\Podes4;
use Illuminate\Support\Str;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

class MainController extends Controller
{
    public function import()
    {
        $kecamatan = ['10', '20', '30', '40', '50', '60', '70', '80', '90', '100', '110', '120', '130', '140', '150', '160', '170', '180', '190', '200', '210', '220', '230', '240'];

        //$kecamatan = ['10', '20', '30'];

        foreach ($kecamatan as $kec) {
            $podes1 = Podes1::where(['R103' => $kec])->get();
            $podes2 = Podes2::where(['R103' => $kec])->get();
            $podes3 = Podes3::where(['R103' => $kec])->get();
            $podes4 = Podes4::where(['R103' => $kec])->get();

            $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('template.xlsx');

            //Tabel T1
            $sheet = $spreadsheet->getSheetByName('T1');
            $title = $sheet->getCell('D1')->getValue();
            $title = Str::replaceFirst('xxx', ucwords(Str::lower($podes1[0]->R103N)), $title);
            $sheet->setCellValue('D1', $title);

            $startrow = 7;
            $row = 5;
            foreach ($podes1 as $record) {
                $sheet->insertNewRowBefore($startrow);
            }

            $index = 1;
            foreach ($podes1 as $record) {
                $sheet->setCellValue('A' . $row, $index . '.');
                $sheet->setCellValue('B' . $row, ucwords(Str::lower($record->R104N)));
                $sheet->setCellValue('E' . $row, $record->R501A1);
                $sheet->setCellValue('F' . $row, $record->R501A2);
                $sheet->setCellValue('G' . $row, $record->R501A1 + $record->R501A2);
                $sheet->setCellValue('H' . $row, $record->R501B);
                $sheet->mergeCells('B' . $row . ':D' . $row);
                $index++;
                $row++;
            }

            for ($i = 0; $i < 3; $i++) {
                $sheet->removeRow($row);
            }

            $total = $sheet->getCell('A' . $row)->getValue();
            $total = Str::replaceFirst('xxx', ucwords(Str::lower($podes1[0]->R103N)), $total);
            $sheet->setCellValue('A' . $row, $total);

            $sheet->setCellValue('E' . $row, '=SUM(E5:E' . ($row - 1) . ')');
            $sheet->setCellValue('F' . $row, '=SUM(F5:F' . ($row - 1) . ')');
            $sheet->setCellValue('G' . $row, '=SUM(G5:G' . ($row - 1) . ')');
            $sheet->setCellValue('H' . $row, '=SUM(H5:H' . ($row - 1) . ')');
            $sheet->getStyle('A' . $row . ':H' . $row)
                ->getBorders()->getTop()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
            //Tabel T1

            //Tabel T6
            $sheet = $spreadsheet->getSheetByName('T6');
            $title = $sheet->getCell('D1')->getValue();
            $title = Str::replaceFirst('xxx', ucwords(Str::lower($podes2[0]->R103N)), $title);
            $sheet->setCellValue('D1', $title);

            $startrow = 7;
            $row = 5;
            foreach ($podes2 as $record) {
                $sheet->insertNewRowBefore($startrow);
            }

            $index = 1;
            foreach ($podes2 as $record) {
                $sheet->setCellValue('A' . $row, $index . '.');
                $sheet->setCellValue('B' . $row, ucwords(Str::lower($record->R104N)));
                $sheet->setCellValue('E' . $row, $record->R701DK2);
                $sheet->setCellValue('F' . $row, $record->R701DK3);
                $sheet->setCellValue('G' . $row, $record->R701DK2 + $record->R701DK3);
                $sheet->mergeCells('B' . $row . ':D' . $row);
                $index++;
                $row++;
            }

            for ($i = 0; $i < 3; $i++) {
                $sheet->removeRow($row);
            }

            $total = $sheet->getCell('A' . $row)->getValue();
            $total = Str::replaceFirst('xxx', ucwords(Str::lower($podes1[0]->R103N)), $total);
            $sheet->setCellValue('A' . $row, $total);
            $sheet->unmergeCells('B' . $row . ':D' . $row);
            $sheet->mergeCells('A' . $row . ':D' . $row);

            $sheet->setCellValue('E' . $row, '=SUM(E5:E' . ($row - 1) . ')');
            $sheet->setCellValue('F' . $row, '=SUM(F5:F' . ($row - 1) . ')');
            $sheet->setCellValue('G' . $row, '=SUM(G5:G' . ($row - 1) . ')');
            $sheet->getStyle('A' . $row . ':G' . $row)
                ->getBorders()->getTop()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
            //Tabel T6

            //Tabel T7
            $sheet = $spreadsheet->getSheetByName('T7');
            $title = $sheet->getCell('D1')->getValue();
            $title = Str::replaceFirst('xxx', ucwords(Str::lower($podes2[0]->R103N)), $title);
            $sheet->setCellValue('D1', $title);

            $startrow = 7;
            $row = 5;
            foreach ($podes2 as $record) {
                $sheet->insertNewRowBefore($startrow);
            }

            $index = 1;
            foreach ($podes2 as $record) {
                $sheet->setCellValue('A' . $row, $index . '.');
                $sheet->setCellValue('B' . $row, ucwords(Str::lower($record->R104N)));
                $sheet->setCellValue('E' . $row, $record->R701EK2);
                $sheet->setCellValue('F' . $row, $record->R701EK3);
                $sheet->setCellValue('G' . $row, $record->R701EK2 + $record->R701EK3);
                $sheet->mergeCells('B' . $row . ':D' . $row);
                $index++;
                $row++;
            }

            for ($i = 0; $i < 3; $i++) {
                $sheet->removeRow($row);
            }

            $total = $sheet->getCell('A' . $row)->getValue();
            $total = Str::replaceFirst('xxx', ucwords(Str::lower($podes2[0]->R103N)), $total);
            $sheet->setCellValue('A' . $row, $total);
            $sheet->unmergeCells('B' . $row . ':D' . $row);
            $sheet->mergeCells('A' . $row . ':D' . $row);

            $sheet->setCellValue('E' . $row, '=SUM(E5:E' . ($row - 1) . ')');
            $sheet->setCellValue('F' . $row, '=SUM(F5:F' . ($row - 1) . ')');
            $sheet->setCellValue('G' . $row, '=SUM(G5:G' . ($row - 1) . ')');
            $sheet->getStyle('A' . $row . ':G' . $row)
                ->getBorders()->getTop()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
            //Tabel T7

            //Tabel T8
            $sheet = $spreadsheet->getSheetByName('T8');
            $title = $sheet->getCell('D1')->getValue();
            $title = Str::replaceFirst('xxx', ucwords(Str::lower($podes2[0]->R103N)), $title);
            $sheet->setCellValue('D1', $title);

            $startrow = 7;
            $row = 5;
            foreach ($podes2 as $record) {
                $sheet->insertNewRowBefore($startrow);
            }

            $index = 1;
            foreach ($podes2 as $record) {
                $sheet->setCellValue('A' . $row, $index . '.');
                $sheet->setCellValue('B' . $row, ucwords(Str::lower($record->R104N)));
                $sheet->setCellValue('E' . $row, $record->R701FK2);
                $sheet->setCellValue('F' . $row, $record->R701FK3);
                $sheet->setCellValue('G' . $row, $record->R701FK2 + $record->R701FK3);
                $sheet->mergeCells('B' . $row . ':D' . $row);
                $index++;
                $row++;
            }

            for ($i = 0; $i < 3; $i++) {
                $sheet->removeRow($row);
            }

            $total = $sheet->getCell('A' . $row)->getValue();
            $total = Str::replaceFirst('xxx', ucwords(Str::lower($podes2[0]->R103N)), $total);
            $sheet->setCellValue('A' . $row, $total);
            $sheet->unmergeCells('B' . $row . ':D' . $row);
            $sheet->mergeCells('A' . $row . ':D' . $row);

            $sheet->setCellValue('E' . $row, '=SUM(E5:E' . ($row - 1) . ')');
            $sheet->setCellValue('F' . $row, '=SUM(F5:F' . ($row - 1) . ')');
            $sheet->setCellValue('G' . $row, '=SUM(G5:G' . ($row - 1) . ')');
            $sheet->getStyle('A' . $row . ':G' . $row)
                ->getBorders()->getTop()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
            //Tabel T8

            //Tabel T9
            $sheet = $spreadsheet->getSheetByName('T9');
            $title = $sheet->getCell('D1')->getValue();
            $title = Str::replaceFirst('xxx', ucwords(Str::lower($podes2[0]->R103N)), $title);
            $sheet->setCellValue('D1', $title);

            $startrow = 7;
            $row = 5;
            foreach ($podes2 as $record) {
                $sheet->insertNewRowBefore($startrow);
            }

            $index = 1;
            foreach ($podes2 as $record) {
                $sheet->setCellValue('A' . $row, $index . '.');
                $sheet->setCellValue('B' . $row, ucwords(Str::lower($record->R104N)));
                $sheet->setCellValue('E' . $row, $record->R701GK2);
                $sheet->setCellValue('F' . $row, $record->R701GK3);
                $sheet->setCellValue('G' . $row, $record->R701GK2 + $record->R701GK3);
                $sheet->mergeCells('B' . $row . ':D' . $row);
                $index++;
                $row++;
            }

            for ($i = 0; $i < 3; $i++) {
                $sheet->removeRow($row);
            }

            $total = $sheet->getCell('A' . $row)->getValue();
            $total = Str::replaceFirst('xxx', ucwords(Str::lower($podes2[0]->R103N)), $total);
            $sheet->setCellValue('A' . $row, $total);
            $sheet->unmergeCells('B' . $row . ':D' . $row);
            $sheet->mergeCells('A' . $row . ':D' . $row);

            $sheet->setCellValue('E' . $row, '=SUM(E5:E' . ($row - 1) . ')');
            $sheet->setCellValue('F' . $row, '=SUM(F5:F' . ($row - 1) . ')');
            $sheet->setCellValue('G' . $row, '=SUM(G5:G' . ($row - 1) . ')');
            $sheet->getStyle('A' . $row . ':G' . $row)
                ->getBorders()->getTop()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
            //Tabel T9

            //Tabel T10
            $sheet = $spreadsheet->getSheetByName('T10');
            $title = $sheet->getCell('D1')->getValue();
            $title = Str::replaceFirst('xxx', ucwords(Str::lower($podes2[0]->R103N)), $title);
            $sheet->setCellValue('D1', $title);

            $startrow = 7;
            $row = 5;
            foreach ($podes2 as $record) {
                $sheet->insertNewRowBefore($startrow);
            }

            $index = 1;
            foreach ($podes2 as $record) {
                $sheet->setCellValue('A' . $row, $index . '.');
                $sheet->setCellValue('B' . $row, ucwords(Str::lower($record->R104N)));
                $sheet->setCellValue('E' . $row, $record->R701HK2);
                $sheet->setCellValue('F' . $row, $record->R701HK3);
                $sheet->setCellValue('G' . $row, $record->R701HK2 + $record->R701HK3);
                $sheet->mergeCells('B' . $row . ':D' . $row);
                $index++;
                $row++;
            }

            for ($i = 0; $i < 3; $i++) {
                $sheet->removeRow($row);
            }

            $total = $sheet->getCell('A' . $row)->getValue();
            $total = Str::replaceFirst('xxx', ucwords(Str::lower($podes2[0]->R103N)), $total);
            $sheet->setCellValue('A' . $row, $total);
            $sheet->unmergeCells('B' . $row . ':D' . $row);
            $sheet->mergeCells('A' . $row . ':D' . $row);

            $sheet->setCellValue('E' . $row, '=SUM(E5:E' . ($row - 1) . ')');
            $sheet->setCellValue('F' . $row, '=SUM(F5:F' . ($row - 1) . ')');
            $sheet->setCellValue('G' . $row, '=SUM(G5:G' . ($row - 1) . ')');
            $sheet->getStyle('A' . $row . ':G' . $row)
                ->getBorders()->getTop()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
            //Tabel T10

            //Tabel T11
            $sheet = $spreadsheet->getSheetByName('T11');
            $title = $sheet->getCell('D1')->getValue();
            $title = Str::replaceFirst('xxx', ucwords(Str::lower($podes2[0]->R103N)), $title);
            $sheet->setCellValue('D1', $title);

            $startrow = 7;
            $row = 5;
            foreach ($podes2 as $record) {
                $sheet->insertNewRowBefore($startrow);
            }

            $index = 1;
            foreach ($podes2 as $record) {
                $sheet->setCellValue('A' . $row, $index . '.');
                $sheet->setCellValue('B' . $row, ucwords(Str::lower($record->R104N)));
                $sheet->setCellValue('E' . $row, $record->R701IK2);
                $sheet->setCellValue('F' . $row, $record->R701IK3);
                $sheet->setCellValue('G' . $row, $record->R701IK2 + $record->R701IK3);
                $sheet->mergeCells('B' . $row . ':D' . $row);
                $index++;
                $row++;
            }

            for ($i = 0; $i < 3; $i++) {
                $sheet->removeRow($row);
            }

            $total = $sheet->getCell('A' . $row)->getValue();
            $total = Str::replaceFirst('xxx', ucwords(Str::lower($podes2[0]->R103N)), $total);
            $sheet->setCellValue('A' . $row, $total);
            $sheet->unmergeCells('B' . $row . ':D' . $row);
            $sheet->mergeCells('A' . $row . ':D' . $row);

            $sheet->setCellValue('E' . $row, '=SUM(E5:E' . ($row - 1) . ')');
            $sheet->setCellValue('F' . $row, '=SUM(F5:F' . ($row - 1) . ')');
            $sheet->setCellValue('G' . $row, '=SUM(G5:G' . ($row - 1) . ')');
            $sheet->getStyle('A' . $row . ':G' . $row)
                ->getBorders()->getTop()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
            //Tabel T11

            //Tabel T12
            $sheet = $spreadsheet->getSheetByName('T12');
            $title = $sheet->getCell('D1')->getValue();
            $title = Str::replaceFirst('xxx', ucwords(Str::lower($podes2[0]->R103N)), $title);
            $sheet->setCellValue('D1', $title);

            $startrow = 7;
            $row = 5;
            foreach ($podes2 as $record) {
                $sheet->insertNewRowBefore($startrow);
            }

            $index = 1;
            foreach ($podes2 as $record) {
                $sheet->setCellValue('A' . $row, $index . '.');
                $sheet->setCellValue('B' . $row, ucwords(Str::lower($record->R104N)));
                $sheet->setCellValue('E' . $row, $record->R701JK2);
                $sheet->setCellValue('F' . $row, $record->R701JK3);
                $sheet->setCellValue('G' . $row, $record->R701JK2 + $record->R701JK3);
                $sheet->mergeCells('B' . $row . ':D' . $row);
                $index++;
                $row++;
            }

            for ($i = 0; $i < 3; $i++) {
                $sheet->removeRow($row);
            }

            $total = $sheet->getCell('A' . $row)->getValue();
            $total = Str::replaceFirst('xxx', ucwords(Str::lower($podes2[0]->R103N)), $total);
            $sheet->setCellValue('A' . $row, $total);
            $sheet->unmergeCells('B' . $row . ':D' . $row);
            $sheet->mergeCells('A' . $row . ':D' . $row);

            $sheet->setCellValue('E' . $row, '=SUM(E5:E' . ($row - 1) . ')');
            $sheet->setCellValue('F' . $row, '=SUM(F5:F' . ($row - 1) . ')');
            $sheet->setCellValue('G' . $row, '=SUM(G5:G' . ($row - 1) . ')');
            $sheet->getStyle('A' . $row . ':G' . $row)
                ->getBorders()->getTop()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
            //Tabel T12

            //Tabel T13
            $sheet = $spreadsheet->getSheetByName('T13');
            $title = $sheet->getCell('D1')->getValue();
            $title = Str::replaceFirst('xxx', ucwords(Str::lower($podes2[0]->R103N)), $title);
            $sheet->setCellValue('D1', $title);

            $startrow = 7;
            $row = 5;
            foreach ($podes2 as $record) {
                $sheet->insertNewRowBefore($startrow);
            }

            $index = 1;
            foreach ($podes2 as $record) {
                $sheet->setCellValue('A' . $row, $index . '.');
                $sheet->setCellValue('B' . $row, ucwords(Str::lower($record->R104N)));
                $sheet->setCellValue('E' . $row, $record->R701KK2);
                $sheet->setCellValue('F' . $row, $record->R701KK3);
                $sheet->setCellValue('G' . $row, $record->R701KK2 + $record->R601KK3);
                $sheet->mergeCells('B' . $row . ':D' . $row);
                $index++;
                $row++;
            }

            for ($i = 0; $i < 3; $i++) {
                $sheet->removeRow($row);
            }

            $total = $sheet->getCell('A' . $row)->getValue();
            $total = Str::replaceFirst('xxx', ucwords(Str::lower($podes2[0]->R103N)), $total);
            $sheet->setCellValue('A' . $row, $total);
            $sheet->unmergeCells('B' . $row . ':D' . $row);
            $sheet->mergeCells('A' . $row . ':D' . $row);

            $sheet->setCellValue('E' . $row, '=SUM(E5:E' . ($row - 1) . ')');
            $sheet->setCellValue('F' . $row, '=SUM(F5:F' . ($row - 1) . ')');
            $sheet->setCellValue('G' . $row, '=SUM(G5:G' . ($row - 1) . ')');
            $sheet->getStyle('A' . $row . ':G' . $row)
                ->getBorders()->getTop()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
            //Tabel T13

            //Tabel T14
            $sheet = $spreadsheet->getSheetByName('T14');
            $title = $sheet->getCell('D1')->getValue();
            $title = Str::replaceFirst('xxx', ucwords(Str::lower($podes2[0]->R103N)), $title);
            $sheet->setCellValue('D1', $title);

            $startrow = 8;
            $row = 6;
            foreach ($podes2 as $record) {
                $sheet->insertNewRowBefore($startrow);
            }

            $index = 1;
            foreach ($podes2 as $record) {
                $sheet->setCellValue('A' . $row, $index . '.');
                $sheet->setCellValue('B' . $row, ucwords(Str::lower($record->R104N)));
                $sheet->setCellValue('E' . $row, '=VLOOKUP(' . $record->R701DK5 . ',Y1:Z5,2,FALSE)');
                $sheet->setCellValue('F' . $row, '=VLOOKUP(' . $record->R701EK5 . ',Y1:Z5,2,FALSE)');
                $sheet->setCellValue('G' . $row, '=VLOOKUP(' . $record->R701FK5 . ',Y1:Z5,2,FALSE)');
                $sheet->setCellValue('H' . $row, '=VLOOKUP(' . $record->R701GK5 . ',Y1:Z5,2,FALSE)');
                $sheet->setCellValue('I' . $row, $index . '.');
                $sheet->setCellValue('J' . $row, ucwords(Str::lower($record->R104N)));
                $sheet->mergeCells('B' . $row . ':D' . $row);
                $sheet->mergeCells('J' . $row . ':L' . $row);
                $sheet->setCellValue('M' . $row, '=VLOOKUP(' . $record->R701HK5 . ',Y1:Z5,2,FALSE)');
                $sheet->setCellValue('N' . $row, '=VLOOKUP(' . $record->R701IK5 . ',Y1:Z5,2,FALSE)');
                $sheet->setCellValue('O' . $row, '=VLOOKUP(' . $record->R701JK5 . ',Y1:Z5,2,FALSE)');
                $sheet->setCellValue('P' . $row, '=VLOOKUP(' . $record->R701KK5 . ',Y1:Z5,2,FALSE)');
                $index++;
                $row++;
            }

            for ($i = 0; $i < 3; $i++) {
                $sheet->removeRow($row);
            }
            //Tabel T14

            //Tabel T15
            $sheet = $spreadsheet->getSheetByName('T15');
            $title = $sheet->getCell('D1')->getValue();
            $title = Str::replaceFirst('xxx', ucwords(Str::lower($podes2[0]->R103N)), $title);
            $sheet->setCellValue('D1', $title);

            $startrow = 9;
            $row = 7;
            foreach ($podes2 as $record) {
                $sheet->insertNewRowBefore($startrow);
            }

            $index = 1;
            foreach ($podes2 as $record) {
                $sheet->setCellValue('A' . $row, $index . '.');
                $sheet->setCellValue('B' . $row, ucwords(Str::lower($record->R104N)));
                $sheet->setCellValue('E' . $row, $record->R704AK2);
                $sheet->setCellValue('F' . $row, $record->R704BK2);
                $sheet->setCellValue('G' . $row, $record->R704FK2);
                $sheet->setCellValue('H' . $row, $index . '.');
                $sheet->setCellValue('I' . $row, ucwords(Str::lower($record->R104N)));
                $sheet->mergeCells('B' . $row . ':D' . $row);
                $sheet->mergeCells('I' . $row . ':K' . $row);
                $sheet->setCellValue('L' . $row, $record->R704CK2);
                $sheet->setCellValue('M' . $row, $record->R704DK2);
                $sheet->setCellValue('N' . $row, $record->R704LK2);
                $index++;
                $row++;
            }

            for ($i = 0; $i < 3; $i++) {
                $sheet->removeRow($row);
            }

            $total = $sheet->getCell('A' . $row)->getValue();
            $total = Str::replaceFirst('xxx', ucwords(Str::lower($podes2[0]->R103N)), $total);
            $sheet->setCellValue('A' . $row, $total);
            $sheet->unmergeCells('B' . $row . ':D' . $row);
            $sheet->mergeCells('A' . $row . ':D' . $row);
            $sheet->setCellValue('H' . $row, $total);
            $sheet->unmergeCells('I' . $row . ':K' . $row);
            $sheet->mergeCells('H' . $row . ':K' . $row);

            $sheet->setCellValue('E' . $row, '=SUM(E7:E' . ($row - 1) . ')');
            $sheet->setCellValue('F' . $row, '=SUM(F7:F' . ($row - 1) . ')');
            $sheet->setCellValue('G' . $row, '=SUM(G7:G' . ($row - 1) . ')');
            $sheet->setCellValue('L' . $row, '=SUM(L7:L' . ($row - 1) . ')');
            $sheet->setCellValue('M' . $row, '=SUM(M7:M' . ($row - 1) . ')');
            $sheet->setCellValue('N' . $row, '=SUM(N7:N' . ($row - 1) . ')');
            $sheet->getStyle('A' . $row . ':N' . $row)
                ->getBorders()->getTop()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);

            //Tabel T15

            //Tabel T16
            $sheet = $spreadsheet->getSheetByName('T16');
            $title = $sheet->getCell('D1')->getValue();
            $title = Str::replaceFirst('xxx', ucwords(Str::lower($podes2[0]->R103N)), $title);
            $sheet->setCellValue('D1', $title);

            $startrow = 9;
            $row = 7;
            foreach ($podes2 as $record) {
                $sheet->insertNewRowBefore($startrow);
            }

            $index = 1;
            foreach ($podes2 as $record) {
                $sheet->setCellValue('A' . $row, $index . '.');
                $sheet->setCellValue('B' . $row, ucwords(Str::lower($record->R104N)));
                $sheet->setCellValue('E' . $row, '=VLOOKUP(' . $record->R704AK4 . ',R1:S5,2,FALSE)');
                $sheet->setCellValue('F' . $row, '=VLOOKUP(' . $record->R704BK4 . ',R1:S5,2,FALSE)');
                $sheet->setCellValue('G' . $row, '=VLOOKUP(' . $record->R704FK4 . ',R1:S5,2,FALSE)');
                $sheet->setCellValue('H' . $row, $index . '.');
                $sheet->setCellValue('I' . $row, ucwords(Str::lower($record->R104N)));
                $sheet->mergeCells('B' . $row . ':D' . $row);
                $sheet->mergeCells('I' . $row . ':K' . $row);
                $sheet->setCellValue('L' . $row, '=VLOOKUP(' . $record->R704CK4 . ',R1:S5,2,FALSE)');
                $sheet->setCellValue('M' . $row, '=VLOOKUP(' . $record->R704DK4 . ',R1:S5,2,FALSE)');
                $sheet->setCellValue('N' . $row, '=VLOOKUP(' . $record->R704LK4 . ',R1:S5,2,FALSE)');
                $index++;
                $row++;
            }

            for ($i = 0; $i < 3; $i++) {
                $sheet->removeRow($row);
            }
            //Tabel T16

            //Tabel T18
            $sheet = $spreadsheet->getSheetByName('T18');
            $title = $sheet->getCell('D1')->getValue();
            $title = Str::replaceFirst('xxx', ucwords(Str::lower($podes1[0]->R103N)), $title);
            $sheet->setCellValue('D1', $title);

            $startrow = 8;
            $row = 6;
            foreach ($podes1 as $record) {
                $sheet->insertNewRowBefore($startrow);
            }

            $index = 1;
            foreach ($podes1 as $record) {
                $sheet->setCellValue('A' . $row, $index . '.');
                $sheet->setCellValue('B' . $row, ucwords(Str::lower($record->R104N)));
                $sheet->setCellValue('E' . $row, $record->R601DK3);
                $sheet->setCellValue('F' . $row, $record->R601EK3);
                $sheet->setCellValue('G' . $row, $record->R601HK3);
                $sheet->setCellValue('H' . $row, $record->R601AK3);
                $sheet->setCellValue('I' . $row, $index . '.');
                $sheet->setCellValue('J' . $row, ucwords(Str::lower($record->R104N)));
                $sheet->setCellValue('M' . $row, $record->R601BK3);
                $sheet->setCellValue('N' . $row, $record->R601CK3);
                $sheet->setCellValue('O' . $row, $record->R601JK3);
                $sheet->setCellValue('P' . $row, $index . '.');
                $sheet->setCellValue('Q' . $row, ucwords(Str::lower($record->R104N)));
                $sheet->setCellValue('T' . $row, $record->R601IK3);
                $sheet->setCellValue('U' . $row, $record->R601GK3);
                $sheet->setCellValue('V' . $row, $record->R601FK3);
                $sheet->mergeCells('B' . $row . ':D' . $row);
                $sheet->mergeCells('J' . $row . ':L' . $row);
                $sheet->mergeCells('Q' . $row . ':S' . $row);
                $index++;
                $row++;
            }

            for ($i = 0; $i < 3; $i++) {
                $sheet->removeRow($row);
            }

            $total = $sheet->getCell('A' . $row)->getValue();
            $total = Str::replaceFirst('xxx', ucwords(Str::lower($podes1[0]->R103N)), $total);
            $sheet->setCellValue('A' . $row, $total);
            $sheet->setCellValue('I' . $row, $total);
            $sheet->setCellValue('P' . $row, $total);
            $sheet->unmergeCells('B' . $row . ':D' . $row);
            $sheet->mergeCells('A' . $row . ':D' . $row);
            $sheet->setCellValue('H' . $row, $total);
            $sheet->unmergeCells('J' . $row . ':L' . $row);
            $sheet->mergeCells('I' . $row . ':L' . $row);
            $sheet->unmergeCells('Q' . $row . ':S' . $row);
            $sheet->mergeCells('P' . $row . ':S' . $row);

            $sheet->setCellValue('E' . $row, '=SUM(E6:E' . ($row - 1) . ')');
            $sheet->setCellValue('F' . $row, '=SUM(F6:F' . ($row - 1) . ')');
            $sheet->setCellValue('G' . $row, '=SUM(G6:G' . ($row - 1) . ')');
            $sheet->setCellValue('H' . $row, '=SUM(H6:H' . ($row - 1) . ')');
            $sheet->setCellValue('M' . $row, '=SUM(M6:M' . ($row - 1) . ')');
            $sheet->setCellValue('N' . $row, '=SUM(N6:N' . ($row - 1) . ')');
            $sheet->setCellValue('O' . $row, '=SUM(O6:O' . ($row - 1) . ')');
            $sheet->setCellValue('T' . $row, '=SUM(T6:T' . ($row - 1) . ')');
            $sheet->setCellValue('U' . $row, '=SUM(U6:U' . ($row - 1) . ')');
            $sheet->setCellValue('V' . $row, '=SUM(V6:V' . ($row - 1) . ')');

            $sheet->getStyle('A' . $row . ':V' . $row)
                ->getBorders()->getTop()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);

            //Tabel T18

            //Tabel T19
            $sheet = $spreadsheet->getSheetByName('T19');
            $title = $sheet->getCell('D1')->getValue();
            $title = Str::replaceFirst('xxx', ucwords(Str::lower($podes1[0]->R103N)), $title);
            $sheet->setCellValue('D1', $title);

            $startrow = 8;
            $row = 6;
            foreach ($podes1 as $record) {
                $sheet->insertNewRowBefore($startrow);
            }

            $index = 1;
            foreach ($podes1 as $record) {
                $sheet->setCellValue('A' . $row, $index . '.');
                $sheet->setCellValue('B' . $row, ucwords(Str::lower($record->R104N)));
                $sheet->setCellValue('E' . $row, $record->R601DK4);
                $sheet->setCellValue('F' . $row, $record->R601EK4);
                $sheet->setCellValue('G' . $row, $record->R601HK4);
                $sheet->setCellValue('H' . $row, $record->R601AK4);
                $sheet->setCellValue('I' . $row, $index . '.');
                $sheet->setCellValue('J' . $row, ucwords(Str::lower($record->R104N)));
                $sheet->setCellValue('M' . $row, $record->R601BK4);
                $sheet->setCellValue('N' . $row, $record->R601CK4);
                $sheet->setCellValue('O' . $row, $record->R601JK4);
                $sheet->setCellValue('P' . $row, $index . '.');
                $sheet->setCellValue('Q' . $row, ucwords(Str::lower($record->R104N)));
                $sheet->setCellValue('T' . $row, $record->R601IK4);
                $sheet->setCellValue('U' . $row, $record->R601GK4);
                $sheet->setCellValue('V' . $row, $record->R601FK4);
                $sheet->mergeCells('B' . $row . ':D' . $row);
                $sheet->mergeCells('J' . $row . ':L' . $row);
                $sheet->mergeCells('Q' . $row . ':S' . $row);
                $index++;
                $row++;
            }

            for ($i = 0; $i < 3; $i++) {
                $sheet->removeRow($row);
            }

            $total = $sheet->getCell('A' . $row)->getValue();
            $total = Str::replaceFirst('xxx', ucwords(Str::lower($podes1[0]->R103N)), $total);
            $sheet->setCellValue('A' . $row, $total);
            $sheet->setCellValue('I' . $row, $total);
            $sheet->setCellValue('P' . $row, $total);
            $sheet->unmergeCells('B' . $row . ':D' . $row);
            $sheet->mergeCells('A' . $row . ':D' . $row);
            $sheet->setCellValue('H' . $row, $total);
            $sheet->unmergeCells('J' . $row . ':L' . $row);
            $sheet->mergeCells('I' . $row . ':L' . $row);
            $sheet->unmergeCells('Q' . $row . ':S' . $row);
            $sheet->mergeCells('P' . $row . ':S' . $row);

            $sheet->setCellValue('E' . $row, '=SUM(E6:E' . ($row - 1) . ')');
            $sheet->setCellValue('F' . $row, '=SUM(F6:F' . ($row - 1) . ')');
            $sheet->setCellValue('G' . $row, '=SUM(G6:G' . ($row - 1) . ')');
            $sheet->setCellValue('H' . $row, '=SUM(H6:H' . ($row - 1) . ')');
            $sheet->setCellValue('M' . $row, '=SUM(M6:M' . ($row - 1) . ')');
            $sheet->setCellValue('N' . $row, '=SUM(N6:N' . ($row - 1) . ')');
            $sheet->setCellValue('O' . $row, '=SUM(O6:O' . ($row - 1) . ')');
            $sheet->setCellValue('T' . $row, '=SUM(T6:T' . ($row - 1) . ')');
            $sheet->setCellValue('U' . $row, '=SUM(U6:U' . ($row - 1) . ')');
            $sheet->setCellValue('V' . $row, '=SUM(V6:V' . ($row - 1) . ')');

            $sheet->getStyle('A' . $row . ':V' . $row)
                ->getBorders()->getTop()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);

            //Tabel T19

            //Tabel T20
            $sheet = $spreadsheet->getSheetByName('T20');
            $title = $sheet->getCell('D1')->getValue();
            $title = Str::replaceFirst('xxx', ucwords(Str::lower($podes1[0]->R103N)), $title);
            $sheet->setCellValue('D1', $title);

            $startrow = 8;
            $row = 6;
            foreach ($podes1 as $record) {
                $sheet->insertNewRowBefore($startrow);
            }

            $index = 1;
            foreach ($podes1 as $record) {
                $sheet->setCellValue('A' . $row, $index . '.');
                $sheet->setCellValue('B' . $row, ucwords(Str::lower($record->R104N)));
                $sheet->setCellValue('E' . $row, '=VLOOKUP(' . $record->R604A . ',P1:Q2,2,FALSE)');
                $sheet->setCellValue('F' . $row, '=VLOOKUP(' . $record->R604B . ',T1:U3,2,FALSE)');
                $sheet->setCellValue('G' . $row, '=VLOOKUP(' . $record->R604C . ',R1:S2,2,FALSE)');
                $sheet->setCellValue('H' . $row, $index . '.');
                $sheet->setCellValue('I' . $row, ucwords(Str::lower($record->R104N)));
                $sheet->setCellValue('L' . $row, '=VLOOKUP(' . $record->R604D . ',V1:W2,2,FALSE)');
                $sheet->setCellValue('M' . $row, '=VLOOKUP(' . $record->R604E . ',P1:Q2,2,FALSE)');
                $sheet->mergeCells('B' . $row . ':D' . $row);
                $sheet->mergeCells('I' . $row . ':K' . $row);
                $index++;
                $row++;
            }

            for ($i = 0; $i < 3; $i++) {
                $sheet->removeRow($row);
            }

            //Tabel T20

            //Tabel T21
            $sheet = $spreadsheet->getSheetByName('T21');
            $title = $sheet->getCell('D1')->getValue();
            $title = Str::replaceFirst('xxx', ucwords(Str::lower($podes3[0]->R103N)), $title);
            $sheet->setCellValue('D1', $title);

            $startrow = 8;
            $row = 6;
            foreach ($podes3 as $record) {
                $sheet->insertNewRowBefore($startrow);
            }

            $index = 1;
            foreach ($podes3 as $record) {
                $sheet->setCellValue('A' . $row, $index . '.');
                $sheet->setCellValue('B' . $row, ucwords(Str::lower($record->R104N)));
                $sheet->setCellValue('E' . $row, $record->R1207AK2);
                $sheet->setCellValue('F' . $row, $record->R1207BK2);
                $sheet->setCellValue('G' . $row, $record->R1207CK2);
                $sheet->setCellValue('H' . $row, $record->R1207DK2);
                $sheet->setCellValue('I' . $row, $index . '.');
                $sheet->setCellValue('J' . $row, ucwords(Str::lower($record->R104N)));
                $sheet->setCellValue('M' . $row, $record->R1207EK2);
                $sheet->setCellValue('N' . $row, $record->R1207JK2);
                $sheet->setCellValue('O' . $row, $record->R1207FK2);
                $sheet->setCellValue('P' . $row, $index . '.');
                $sheet->setCellValue('Q' . $row, ucwords(Str::lower($record->R104N)));
                $sheet->setCellValue('T' . $row, $record->R1207GK2);
                $sheet->setCellValue('U' . $row, $record->R1207HK2);
                $sheet->setCellValue('V' . $row, $record->R1207IK2);
                $sheet->mergeCells('B' . $row . ':D' . $row);
                $sheet->mergeCells('J' . $row . ':L' . $row);
                $sheet->mergeCells('Q' . $row . ':S' . $row);
                $index++;
                $row++;
            }

            for ($i = 0; $i < 3; $i++) {
                $sheet->removeRow($row);
            }

            $total = $sheet->getCell('A' . $row)->getValue();
            $total = Str::replaceFirst('xxx', ucwords(Str::lower($podes3[0]->R103N)), $total);
            $sheet->setCellValue('A' . $row, $total);
            $sheet->setCellValue('I' . $row, $total);
            $sheet->setCellValue('P' . $row, $total);
            $sheet->unmergeCells('B' . $row . ':D' . $row);
            $sheet->mergeCells('A' . $row . ':D' . $row);
            $sheet->setCellValue('H' . $row, $total);
            $sheet->unmergeCells('J' . $row . ':L' . $row);
            $sheet->mergeCells('I' . $row . ':L' . $row);
            $sheet->unmergeCells('Q' . $row . ':S' . $row);
            $sheet->mergeCells('P' . $row . ':S' . $row);

            $sheet->setCellValue('E' . $row, '=SUM(E6:E' . ($row - 1) . ')');
            $sheet->setCellValue('F' . $row, '=SUM(F6:F' . ($row - 1) . ')');
            $sheet->setCellValue('G' . $row, '=SUM(G6:G' . ($row - 1) . ')');
            $sheet->setCellValue('H' . $row, '=SUM(H6:H' . ($row - 1) . ')');
            $sheet->setCellValue('M' . $row, '=SUM(M6:M' . ($row - 1) . ')');
            $sheet->setCellValue('N' . $row, '=SUM(N6:N' . ($row - 1) . ')');
            $sheet->setCellValue('O' . $row, '=SUM(O6:O' . ($row - 1) . ')');
            $sheet->setCellValue('T' . $row, '=SUM(T6:T' . ($row - 1) . ')');
            $sheet->setCellValue('U' . $row, '=SUM(U6:U' . ($row - 1) . ')');
            $sheet->setCellValue('V' . $row, '=SUM(V6:V' . ($row - 1) . ')');

            $sheet->getStyle('A' . $row . ':V' . $row)
                ->getBorders()->getTop()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);

            //Tabel T21

            //Tabel T22
            $sheet = $spreadsheet->getSheetByName('T22');
            $title = $sheet->getCell('D1')->getValue();
            $title = Str::replaceFirst('xxx', ucwords(Str::lower($podes3[0]->R103N)), $title);
            $sheet->setCellValue('D1', $title);

            $startrow = 7;
            $row = 5;
            foreach ($podes3 as $record) {
                $sheet->insertNewRowBefore($startrow);
            }

            $index = 1;
            foreach ($podes3 as $record) {
                $sheet->setCellValue('A' . $row, $index . '.');
                $sheet->setCellValue('B' . $row, ucwords(Str::lower($record->R104N)));
                $sheet->setCellValue('E' . $row, $record->R1205A1);
                $sheet->setCellValue('F' . $row, $record->R1205A2);
                $sheet->setCellValue('G' . $row, $record->R1205A3);
                $sheet->mergeCells('B' . $row . ':D' . $row);
                $index++;
                $row++;
            }

            for ($i = 0; $i < 3; $i++) {
                $sheet->removeRow($row);
            }

            $total = $sheet->getCell('A' . $row)->getValue();
            $total = Str::replaceFirst('xxx', ucwords(Str::lower($podes3[0]->R103N)), $total);
            $sheet->setCellValue('A' . $row, $total);
            $sheet->unmergeCells('B' . $row . ':D' . $row);
            $sheet->mergeCells('A' . $row . ':D' . $row);

            $sheet->setCellValue('E' . $row, '=SUM(E5:E' . ($row - 1) . ')');
            $sheet->setCellValue('F' . $row, '=SUM(F5:F' . ($row - 1) . ')');
            $sheet->setCellValue('G' . $row, '=SUM(G5:G' . ($row - 1) . ')');
            $sheet->getStyle('A' . $row . ':G' . $row)
                ->getBorders()->getTop()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
            //Tabel T22

            //Tabel T23
            $sheet = $spreadsheet->getSheetByName('T23');
            $title = $sheet->getCell('D1')->getValue();
            $title = Str::replaceFirst('xxx', ucwords(Str::lower($podes3[0]->R103N)), $title);
            $sheet->setCellValue('D1', $title);

            $startrow = 7;
            $row = 5;
            foreach ($podes3 as $record) {
                $sheet->insertNewRowBefore($startrow);
            }

            $index = 1;
            foreach ($podes3 as $record) {
                $sheet->setCellValue('A' . $row, $index . '.');
                $sheet->setCellValue('B' . $row, ucwords(Str::lower($record->R104N)));
                $sheet->setCellValue('E' . $row, $record->R1206A1);
                $sheet->setCellValue('F' . $row, $record->R1206A2);
                $sheet->setCellValue('G' . $row, $record->R1206A3);
                $sheet->setCellValue('H' . $row, $record->R1206A4);
                $sheet->mergeCells('B' . $row . ':D' . $row);
                $index++;
                $row++;
            }

            for ($i = 0; $i < 3; $i++) {
                $sheet->removeRow($row);
            }

            $total = $sheet->getCell('A' . $row)->getValue();
            $total = Str::replaceFirst('xxx', ucwords(Str::lower($podes3[0]->R103N)), $total);
            $sheet->setCellValue('A' . $row, $total);
            $sheet->unmergeCells('B' . $row . ':D' . $row);
            $sheet->mergeCells('A' . $row . ':D' . $row);

            $sheet->setCellValue('E' . $row, '=SUM(E5:E' . ($row - 1) . ')');
            $sheet->setCellValue('F' . $row, '=SUM(F5:F' . ($row - 1) . ')');
            $sheet->setCellValue('G' . $row, '=SUM(G5:G' . ($row - 1) . ')');
            $sheet->setCellValue('H' . $row, '=SUM(H5:H' . ($row - 1) . ')');
            $sheet->getStyle('A' . $row . ':H' . $row)
                ->getBorders()->getTop()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
            //Tabel T23

            //Tabel T25
            $sheet = $spreadsheet->getSheetByName('T25');
            $title = $sheet->getCell('D1')->getValue();
            $title = Str::replaceFirst('xxx', ucwords(Str::lower($podes3[0]->R103N)), $title);
            $sheet->setCellValue('D1', $title);

            $startrow = 7;
            $row = 5;
            foreach ($podes3 as $record) {
                $sheet->insertNewRowBefore($startrow);
            }

            $index = 1;
            foreach ($podes3 as $record) {
                $sheet->setCellValue('A' . $row, $index . '.');
                $sheet->setCellValue('B' . $row, ucwords(Str::lower($record->R104N)));
                $sheet->setCellValue('E' . $row, $record->R1005A);
                $sheet->setCellValue('F' . $row, $record->R1005B);
                $sheet->setCellValue('G' . $row, '=VLOOKUP(' . $record->R1005C . ',K1:L4,2,FALSE)');
                $sheet->mergeCells('B' . $row . ':D' . $row);
                $index++;
                $row++;
            }

            for ($i = 0; $i < 3; $i++) {
                $sheet->removeRow($row);
            }

            $total = $sheet->getCell('A' . $row)->getValue();
            $total = Str::replaceFirst('xxx', ucwords(Str::lower($podes3[0]->R103N)), $total);
            $sheet->setCellValue('A' . $row, $total);
            $sheet->unmergeCells('B' . $row . ':D' . $row);
            $sheet->mergeCells('A' . $row . ':D' . $row);

            $sheet->setCellValue('E' . $row, '=SUM(E5:E' . ($row - 1) . ')');
            $sheet->setCellValue('F' . $row, '=SUM(F5:F' . ($row - 1) . ')');
            //$sheet->setCellValue('G' . $row, '=SUM(G5:G' . ($row - 1) . ')');
            $sheet->getStyle('A' . $row . ':G' . $row)
                ->getBorders()->getTop()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
            //Tabel T25

            //Tabel T26
            $sheet = $spreadsheet->getSheetByName('T26');
            $title = $sheet->getCell('D1')->getValue();
            $title = Str::replaceFirst('xxx', ucwords(Str::lower($podes3[0]->R103N)), $title);
            $sheet->setCellValue('D1', $title);

            $arrayOlahraga = array(
                'Sepak bola' => 'R901AK2',
                'Bola voli' => 'R901BK2',
                'Bulu tangkis' => 'R901CK2',
                'Bola basket' => 'R901DK2',
                'Tenis lapangan' => 'R901EK2',
                'Tenis meja' => 'R901FK2',
                'Futsal' => 'R901GK2',
                'Renang' => 'R901HK2',
                'Bela diri (pencak silat, karate, dll)' => 'R901IK2',
                'Bilyard' => 'R901JK2',
                'Pusat kebugaran (senam, fitness, aerobik, dll)' => 'R901KK2',
                'Lainnya' => 'R901LK2',
            );

            $arrayValue = ['1', '2', '3', '4'];

            $map = array();

            foreach ($arrayOlahraga as $olahraga => $column) {
                foreach ($arrayValue as $value) {
                    $map[$olahraga][$value] = 0;
                }
            }

            foreach ($podes3 as $record) {
                foreach ($arrayOlahraga as $olahraga => $column) {
                    $map[$olahraga][$record[$column]]++;
                }
            }

            $row = 6;
            foreach ($arrayOlahraga as $olahraga => $column) {
                $sheet->setCellValue('E' . $row, $map[$olahraga]['1']);
                $sheet->setCellValue('F' . $row, $map[$olahraga]['2']);
                $sheet->setCellValue('G' . $row, $map[$olahraga]['3']);
                $sheet->setCellValue('H' . $row, $map[$olahraga]['4']);
                $row++;
            }
            //Tabel T26

            //Tabel T27
            $sheet = $spreadsheet->getSheetByName('T27');
            $title = $sheet->getCell('D1')->getValue();
            $title = Str::replaceFirst('xxx', ucwords(Str::lower($podes3[0]->R103N)), $title);
            $sheet->setCellValue('D1', $title);

            $startrow = 7;
            $row = 5;
            foreach ($podes3 as $record) {
                $sheet->insertNewRowBefore($startrow);
            }

            $index = 1;
            foreach ($podes3 as $record) {
                $sheet->setCellValue('A' . $row, $index . '.');
                $sheet->setCellValue('B' . $row, ucwords(Str::lower($record->R104N)));
                $sheet->setCellValue('E' . $row, '=VLOOKUP(' . $record->R1001A . ',I1:J4,2,FALSE)');
                $sheet->setCellValue('F' . $row, '=VLOOKUP(' . $record->R1001C1 . ',L1:M3,2,FALSE)');
                $sheet->mergeCells('B' . $row . ':D' . $row);
                $index++;
                $row++;
            }

            for ($i = 0; $i < 3; $i++) {
                $sheet->removeRow($row);
            }

            //Tabel T27

            //Tabel T28
            $sheet = $spreadsheet->getSheetByName('T28');
            $title = $sheet->getCell('D1')->getValue();
            $title = Str::replaceFirst('xxx', ucwords(Str::lower($podes3[0]->R103N)), $title);
            $sheet->setCellValue('D1', $title);

            $startrow = 7;
            $row = 5;
            foreach ($podes3 as $record) {
                $sheet->insertNewRowBefore($startrow);
            }

            $index = 1;
            foreach ($podes3 as $record) {
                $sheet->setCellValue('A' . $row, $index . '.');
                $sheet->setCellValue('B' . $row, ucwords(Str::lower($record->R104N)));
                $sheet->setCellValue('E' . $row, '=VLOOKUP(' . $record->R1001B1 . ',H1:I5,2,FALSE)');
                $sheet->setCellValue('F' . $row, '=VLOOKUP(' . $record->R1001B2 . ',K1:L4,2,FALSE)');
                $sheet->mergeCells('B' . $row . ':D' . $row);
                $index++;
                $row++;
            }

            for ($i = 0; $i < 3; $i++) {
                $sheet->removeRow($row);
            }

            //Tabel T28

            //Tabel T29
            $sheet = $spreadsheet->getSheetByName('T29');
            $title = $sheet->getCell('D1')->getValue();
            $title = Str::replaceFirst('xxx', ucwords(Str::lower($podes3[0]->R103N)), $title);
            $sheet->setCellValue('D1', $title);

            $startrow = 7;
            $row = 5;
            foreach ($podes3 as $record) {
                $sheet->insertNewRowBefore($startrow);
            }

            $index = 1;
            foreach ($podes3 as $record) {
                $sheet->setCellValue('A' . $row, $index . '.');
                $sheet->setCellValue('B' . $row, ucwords(Str::lower($record->R104N)));
                $sheet->setCellValue('E' . $row, '=VLOOKUP(' . $record->R1007A . ',H1:I4,2,FALSE)');
                $sheet->setCellValue('F' . $row, '=VLOOKUP(' . $record->R1007C . ',H1:I4,2,FALSE)');
                $sheet->mergeCells('B' . $row . ':D' . $row);
                $index++;
                $row++;
            }

            for ($i = 0; $i < 3; $i++) {
                $sheet->removeRow($row);
            }

            //Tabel T29

            // //Tabel T17
            // $sheet = $spreadsheet->getSheetByName('T17');
            // $title = $sheet->getCell('D1')->getValue();
            // $title = Str::replaceFirst('xxx', ucwords(Str::lower($podes1[0]->R103N)), $title);
            // $sheet->setCellValue('D1', $title);

            // $startrow = 7;
            // $row = 5;
            // foreach ($podes1 as $record) {
            //     $sheet->insertNewRowBefore($startrow);
            // }

            // $index = 1;
            // foreach ($podes1 as $record) {
            //     $sheet->setCellValue('A' . $row, $index . '.');
            //     $sheet->setCellValue('B' . $row, ucwords(Str::lower($record->R104N)));
            //     $sheet->setCellValue('E' . $row, $record->R1207_2019);
            //     $sheet->setCellValue('F' . $row, $record->R604);
            //     $sheet->mergeCells('B' . $row . ':D' . $row);
            //     $index++;
            //     $row++;
            // }

            // for ($i = 0; $i < 3; $i++) {
            //     $sheet->removeRow($row);
            // }

            // $total = $sheet->getCell('A' . $row)->getValue();
            // $total = Str::replaceFirst('xxx', ucwords(Str::lower($podes1[0]->R103N)), $total);
            // $sheet->setCellValue('A' . $row, $total);
            // $sheet->unmergeCells('B' . $row . ':D' . $row);
            // $sheet->mergeCells('A' . $row . ':D' . $row);

            // $sheet->setCellValue('E' . $row, '=SUM(E5:E' . ($row - 1) . ')');
            // $sheet->setCellValue('F' . $row, '=SUM(F5:F' . ($row - 1) . ')');
            // $sheet->getStyle('A' . $row . ':F' . $row)
            //     ->getBorders()->getTop()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
            // //Tabel T17

            //Tabel T2
            $sheet = $spreadsheet->getSheetByName('T2');
            $title = $sheet->getCell('D1')->getValue();
            $title = Str::replaceFirst('xxx', ucwords(Str::lower($podes1[0]->R103N)), $title);
            $sheet->setCellValue('D1', $title);

            $arrayContent = array(
                '0' => '-',
                '1' => 'Listrik Pemerintah',
                '2' => 'Listrik Non Pemerintah',
                '3' => 'Non Listrik',
            );
            $year = array(
                '2019' => 'R402B_2019',
                '2020' => 'R502B_2020',
                '2021' => 'R502C',
            );

            $map = array();

            foreach ($arrayContent as $content => $value) {
                foreach ($year as $y => $yv) {
                    $map[$y][$content] = 0;
                }
            }

            foreach ($podes1 as $record) {
                foreach ($year as $y => $yv) {
                    $map[$y][$record[$yv]]++;
                }
            }

            $row = 5;
            $col = 'E';
            foreach ($year as $y => $yv) {
                foreach ($arrayContent as $content => $value) {
                    if ($content != '0') {
                        $sheet->setCellValue($col . $row, $map[$y][$content]);
                        $sheet->setCellValue('B' . $row, $value);
                        $row++;
                    }
                }
                $row = 5;
                $col++;
            }

            $nolight = array();
            foreach ($year as $y => $yv) {
                if ($map[$y]['0'] > 0) {
                    $nolight[] = $y;
                }
            }

            if (count($nolight) > 0) {
                $sheet->setCellValue('C8', 'Ada desa yang tidak memiliki penerangan di jalan utama desa pada tahun ' . implode(', ', $nolight));
            }
            //Tabel T2

            //Tabel T3
            $sheet = $spreadsheet->getSheetByName('T3');
            $title = $sheet->getCell('D1')->getValue();
            $title = Str::replaceFirst('xxx', ucwords(Str::lower($podes1[0]->R103N)), $title);
            $sheet->setCellValue('D1', $title);

            // $arrayContent = array(
            //     '1' => 'Gas Kota',
            //     '2' => 'LPG 3 kg',
            //     '3' => 'LPG lebih dari 3 kg',
            //     '4' => 'Minyak tanah',
            //     '5' => 'Kayu bakar',
            //     '6' => 'Lainnya',
            // );

            $arrayContent = array(
                '5' => 'Gas Kota',
                '4' => 'LPG 3 kg',
                '2' => 'LPG 5,5 kg',
                '3' => 'LPG 12 kg',
                '7' => 'Minyak tanah',
                '10' => 'Kayu bakar',
                '11' => 'Lainnya',
            );

            $year = array(
                // '2019' => 'R403_2019',
                // '2020' => 'R503B_2020',
                '2021' => 'R503B',
            );

            $map = array();

            foreach ($arrayContent as $content => $value) {
                foreach ($year as $y => $yv) {
                    $map[$y][$content] = 0;
                }
            }

            foreach ($podes1 as $record) {
                foreach ($year as $y => $yv) {
                    $map[$y][$record[$yv]]++;
                }
            }

            $row = 6;
            $col = 'E';
            foreach ($year as $y => $yv) {
                foreach ($arrayContent as $content => $value) {
                    $sheet->setCellValue($col . $row, $map[$y][$content]);
                    $sheet->setCellValue('B' . $row, $value);
                    $row++;
                }
                $row = 6;
                $col++;
            }
            //Tabel T3

            //Tabel T4
            $sheet = $spreadsheet->getSheetByName('T4');
            $title = $sheet->getCell('D1')->getValue();
            $title = Str::replaceFirst('xxx', ucwords(Str::lower($podes1[0]->R103N)), $title);
            $sheet->setCellValue('D1', $title);

            $arrayContent = array(
                '1' => 'Air Kemasan Bermerk',
                '2' => 'Air Isi Ulang',
                '3' => 'Ledeng Dengan Meteran',
                '4' => 'Ledeng Tanpa Meteran',
                '5' => 'Sumur Bor atau Pompa',
                '6' => 'Sumur',
                '7' => 'Mata Air',
                '8' => 'Sungai/Danau/Kolam/ Waduk/Situ/Embung/Bendungan',
                '9' => 'Air Hujan',
                '10' => 'Lainnya',
            );

            $year = array(
                '2019' => 'R404_2019',
                '2020' => 'R508_2020',
                '2021' => 'R507A',
            );

            $map = array();

            foreach ($arrayContent as $content => $value) {
                foreach ($year as $y => $yv) {
                    $map[$y][$content] = 0;
                }
            }

            foreach ($podes1 as $record) {
                foreach ($year as $y => $yv) {
                    $map[$y][$record[$yv]]++;
                }
            }

            $row = 6;
            $col = 'E';
            foreach ($year as $y => $yv) {
                foreach ($arrayContent as $content => $value) {
                    $sheet->setCellValue($col . $row, $map[$y][$content]);
                    $sheet->setCellValue('B' . $row, $value);
                    $row++;
                }
                $row = 6;
                $col++;
            }
            //Tabel T4

            //Tabel T5
            $sheet = $spreadsheet->getSheetByName('T5');
            $title = $sheet->getCell('D1')->getValue();
            $title = Str::replaceFirst('xxx', ucwords(Str::lower($podes1[0]->R103N)), $title);
            $sheet->setCellValue('D1', $title);

            $arrayContent = array(
                '1' => 'Sendiri',
                '2' => 'Bersama',
                '3' => 'Umum',
                '4' => 'Bukan Jamban',
            );

            $year = array(
                '2019' => 'R405A_2019',
                '2020' => 'R507A_2020',
                '2021' => 'R505A',
            );

            $map = array();

            foreach ($arrayContent as $content => $value) {
                foreach ($year as $y => $yv) {
                    $map[$y][$content] = 0;
                }
            }

            foreach ($podes1 as $record) {
                foreach ($year as $y => $yv) {
                    $map[$y][$record[$yv]]++;
                }
            }

            $row = 7;
            $col = 'E';
            foreach ($year as $y => $yv) {
                $total = 0;
                foreach ($arrayContent as $content => $value) {
                    $sheet->setCellValue($col . $row, $map[$y][$content]);
                    $sheet->setCellValue('B' . $row, $value);
                    if ($content != '4') $total = $total + $map[$y][$content];
                    $row++;
                }
                $sheet->setCellValue($col . ($row - 5), $total);
                $row = 7;
                $col++;
            }
            //Tabel T5

            //Tabel T24
            $sheet = $spreadsheet->getSheetByName('T24');
            $title = $sheet->getCell('D1')->getValue();
            $title = Str::replaceFirst('xxx', ucwords(Str::lower($podes1[0]->R103N)), $title);
            $sheet->setCellValue('D1', $title);

            $startrow = 7;
            $row = 5;
            foreach ($podes1 as $record) {
                $sheet->insertNewRowBefore($startrow);
            }

            $index = 1;
            foreach ($podes1 as $record) {
                $sheet->setCellValue('A' . $row, $index . '.');
                $sheet->setCellValue('B' . $row, ucwords(Str::lower($record->R104N)));
                $sheet->setCellValue('E' . $row, $record->R511_2020);
                $sheet->setCellValue('F' . $row, $record->R511B);
                $sheet->mergeCells('B' . $row . ':D' . $row);
                $index++;
                $row++;
            }

            for ($i = 0; $i < 3; $i++) {
                $sheet->removeRow($row);
            }

            $total = $sheet->getCell('A' . $row)->getValue();
            $total = Str::replaceFirst('xxx', ucwords(Str::lower($podes1[0]->R103N)), $total);
            $sheet->setCellValue('A' . $row, $total);
            $sheet->unmergeCells('B' . $row . ':D' . $row);
            $sheet->mergeCells('A' . $row . ':D' . $row);

            $sheet->setCellValue('E' . $row, '=SUM(E5:E' . ($row - 1) . ')');
            $sheet->setCellValue('F' . $row, '=SUM(F5:F' . ($row - 1) . ')');
            $sheet->getStyle('A' . $row . ':F' . $row)
                ->getBorders()->getTop()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
            //Tabel T24

            $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
            $writer->save("result/" . $kec . "_2021.xlsx");
        }

        return 'done';
    }
}
