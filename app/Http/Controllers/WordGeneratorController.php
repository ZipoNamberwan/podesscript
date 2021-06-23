<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use PhpOffice\PhpWord\TemplateProcessor;

class WordGeneratorController extends Controller
{
    public function generate()
    {

        $templateProcessor = new TemplateProcessor('template.docx');
        $templateProcessor->setValue('kecamatan', 'Sumber');

        $values = [
            ['test1' => 1, 'test2' => 'Batman'],
            ['test1' => 2, 'test2' => 'Superman'],
        ];
        $templateProcessor->cloneRowAndSetValues('test1', $values);

        $templateProcessor->saveAs('result/Sample_07_TemplateCloneRow.docx');
    }
}
