<?php

namespace App\Http\Controllers;

use Intervention\Image\Facades\Image;

class InfografisController extends Controller
{
    public function generateImage()
    {
        $color = [
            'primary' => '#424242',
            'secondary' => '#bcbcbc',
            'tertiary' => '#e9e9e9',
            'orange' => '#ff9900',
            'blue' => '#3399bb',
            'black' => '#171717'
        ];

        //Bab 6
        $img = Image::make('template_image/6.png');
        $img->text('40 868 jiwa', 350, 695, function ($font) use ($color) {
            $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
            $font->size(70);
            $font->color($color['primary']);
        });
        $img->text('20 019', 137, 800, function ($font) use ($color) {
            $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
            $font->size(30);
            $font->color($color['orange']);
        });
        $img->text('20 849', 822, 800, function ($font) use ($color) {
            $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
            $font->size(30);
            $font->color($color['blue']);
        });
        $img->text('96', 395, 1080, function ($font) use ($color) {
            $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
            $font->size(50);
            $font->color($color['primary']);
        });
        $img->text('953', 627, 1080, function ($font) use ($color) {
            $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
            $font->size(50);
            $font->color($color['primary']);
        });

        $p1 = explode('\n', 'Penduduk Kecamatan Wonomerto \n berdasarkan data Hasil Sensus \n Penduduk Tahun 2020 adalah ');
        for ($i = 0; $i < count($p1); $i++) {
            $offset = 295 + ($i * 55);
            $img->text($p1[$i], 510, $offset, function ($font) use ($color) {
                $font->file(public_path('assets/font/Proxima Nova Regular.otf'));
                $font->size(22);
                $font->color($color['primary']);
                $font->align('right');
            });
        }

        $p2 = explode('\n', ' Rasio jenis kelamin penduduk \n Kecamatan Wonomerto adalah 96, yang \n berarti setiap 100 perempuan terdapat \n sekitar 96 laki-laki,');
        for ($i = 0; $i < count($p2); $i++) {
            $offset = 300 + ($i * 55);
            $img->text($p2[$i], 530, $offset, function ($font) use ($color) {
                $font->file(public_path('assets/font/Proxima Nova Regular.otf'));
                $font->size(22);
                $font->color($color['primary']);
                $font->align('left');
            });
        }

        $img->text('Kecamatan Wonomerto, 2020', 462, 581, function ($font) use ($color) {
            $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
            $font->size(25);
            $font->color($color['black']);
        });
        $img->save('template_image/6 copy.png');
        //Bab 6

        //Bab 2
        $img = Image::make('template_image/2.png');
        $img->text('Wonomerto', 725, 724, function ($font) use ($color) {
            $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
            $font->size(22);
            $font->color($color['black']);
        });
        $img->text('Wonomerto', 718, 1121, function ($font) use ($color) {
            $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
            $font->size(22);
            $font->color($color['black']);
        });
        $img->text('Wonomerto', 718, 1121, function ($font) use ($color) {
            $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
            $font->size(22);
            $font->color($color['black']);
        });
        $img->text('22', 307, 970, function ($font) use ($color) {
            $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
            $font->size(27);
            $font->color($color['orange']);
        });
        $img->text('10', 520, 970, function ($font) use ($color) {
            $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
            $font->size(27);
            $font->color($color['blue']);
        });
        $img->save('template_image/2 copy.png');
        //Bab 2

        return 'done';
    }
}
