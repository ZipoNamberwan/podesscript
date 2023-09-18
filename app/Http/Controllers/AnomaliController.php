<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;

class AnomaliController extends Controller
{
    function generate()
    {
        $anomali = \PhpOffice\PhpSpreadsheet\IOFactory::load('data/anomali.xlsx');
        $sheet = $anomali->getActiveSheet();
        for ($i = 0; $i < 10; $i++) {
            $lastColumn = 'DP';
            $lastColumn++;
            for ($column = 'A'; $column != $lastColumn; $column++) {
            }
        }

        return 'done';
    }
}
