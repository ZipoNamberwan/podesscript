<?php

namespace App\Http\Controllers;

use App\Models\Infografis;
use App\Models\Podes1;
use App\Models\Podes2;
use App\Models\Podes3;
use Illuminate\Support\Facades\Storage;
use Intervention\Image\Facades\Image;
use Illuminate\Support\Str;

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
            'black' => '#171717',
            'green' => '#009900',
            'violet' => 'dasdas'
        ];

        // $kecamatan = ['10', '20', '30', '40', '50', '60', '70', '80', '90', '100', '110', '120', '130', '140', '150', '160', '170', '180', '190', '200', '210', '220', '230', '240'];
        $kecamatan = ['10', '20', '30', '40', '50', '60', '70'];

        foreach ($kecamatan as $kec) {

            Storage::makeDirectory('public/' . $kec);

            $podes1 = Podes1::where(['R103' => $kec])->get();
            $podes2 = Podes2::where(['R103' => $kec])->get();
            $podes3 = Podes3::where(['R103' => $kec])->get();
            $infografis = Infografis::where(['R103' => $kec])->get();

            //Bab 2 PODES
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
            $img->text(
                $podes2->sum('R701DK2') +
                    $podes2->sum('R701DK3') +
                    $podes2->sum('R701EK2') +
                    $podes2->sum('R701EK3'),
                322,
                970,
                function ($font) use ($color) {
                    $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
                    $font->size(27);
                    $font->color($color['orange']);
                    $font->align('center');
                }
            );
            $img->text(
                $podes2->sum('R701FK2') +
                    $podes2->sum('R701FK3') +
                    $podes2->sum('R701GK2') +
                    $podes2->sum('R701GK3'),
                535,
                970,
                function ($font) use ($color) {
                    $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
                    $font->size(27);
                    $font->color($color['blue']);
                    $font->align('center');
                }
            );
            $img->text(
                $podes2->sum('R701HK2') +
                    $podes2->sum('R701HK3') +
                    $podes2->sum('R701IK2') +
                    $podes2->sum('R701IK3') +
                    $podes2->sum('R701JK2') +
                    $podes2->sum('R701JK3'),
                732,
                970,
                function ($font) use ($color) {
                    $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
                    $font->size(27);
                    $font->color($color['green']);
                    $font->align('center');
                }
            );
            $img->text($podes2->sum('R704AK2'), 447, 1332, function ($font) use ($color) {
                $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
                $font->size(27);
                $font->color($color['secondary']);
                $font->align('center');
            });
            $img->text($podes2->sum('R704CK2') + $podes2->sum('R704DK2'), 665, 1332, function ($font) use ($color) {
                $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
                $font->size(27);
                $font->color($color['secondary']);
                $font->align('center');
            });
            $img->save('storage/' . $kec . '/2.png');
            //Bab 2 PODES

            //Bab 7 PODES
            $img = Image::make('template_image/7.png');
            $img->text(number_format($podes1->sum('R501A1'), 0, '.', ' '), 718, 442, function ($font) use ($color) {
                $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
                $font->size(45);
                $font->color($color['tertiary']);
                $font->align('center');
            });
            $img->text(count($podes1->where('R503B', 4)), 310, 1043, function ($font) use ($color) {
                $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
                $font->size(24);
                $font->color($color['primary']);
                $font->align('center');
            });
            $img->text(count($podes1), 405, 1043, function ($font) use ($color) {
                $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
                $font->size(24);
                $font->color($color['primary']);
                $font->align('center');
            });
            $img->text(count($podes1->whereIn('R507A', [3, 4])), 687, 1043, function ($font) use ($color) {
                $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
                $font->size(24);
                $font->color($color['primary']);
                $font->align('center');
            });
            $img->text(count($podes1), 785, 1043, function ($font) use ($color) {
                $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
                $font->size(24);
                $font->color($color['primary']);
                $font->align('center');
            });
            $img->save('storage/' . $kec . '/7.png');
            //Bab 7 PODES

            //Bab 8 PODES
            $img = Image::make('template_image/8.png');
            $img->text($podes3->sum('R1207BK2') + $podes3->sum('R1207CK2') + $podes3->sum('R1207DK2'), 853, 457, function ($font) use ($color) {
                $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
                $font->size(80);
                $font->color($color['primary']);
                $font->align('center');
            });
            $img->text($podes3->sum('R1207EK2'), 905, 642, function ($font) use ($color) {
                $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
                $font->size(80);
                $font->color($color['primary']);
                $font->align('center');
            });
            $img->text($podes3->sum('R1207HK2'), 858, 840, function ($font) use ($color) {
                $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
                $font->size(80);
                $font->color($color['primary']);
                $font->align('center');
            });
            $img->text($podes3->sum('R1207FK2'), 545, 1136, function ($font) use ($color) {
                $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
                $font->size(80);
                $font->color($color['primary']);
                $font->align('center');
            });
            $img->text($podes3->sum('R1205A1') + $podes3->sum('R1205A2') + $podes3->sum('R1205A3'), 337, 1146, function ($font) use ($color) {
                $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
                $font->size(80);
                $font->color($color['primary']);
                $font->align('center');
            });
            $img->text(ucwords(Str::lower($podes3[0]->R103N)), 306, 421, function ($font) use ($color) {
                $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
                $font->size(19);
                $font->color($color['primary']);
            });
            $img->save('storage/' . $kec . '/8.png');
            //Bab 8 PODES

            //Bab 9 PODES
            $img = Image::make('template_image/9.png');

            $percent1 = count($podes3->where('R1001B1', 1)) / count($podes3);
            $percent2 = count($podes3->whereIn('R1001C1', [1, 2])) / count($podes3);
            $img->text(ceil($percent1 * 100) . '%', 313, 920, function ($font) use ($color) {
                $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
                $font->size(50);
                $font->color($color['blue']);
                $font->align('center');
            });
            $img->text(ceil($percent2 * 100) . '%', 738, 920, function ($font) use ($color) {
                $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
                $font->size(50);
                $font->color($color['orange']);
                $font->align('center');
            });
            $img->text(count($podes3->where('R1001B1', 1)) . ' dari ' . count($podes3) . ' Desa', 250, 1030, function ($font) use ($color) {
                $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
                $font->size(25);
                $font->color($color['blue']);
                $font->align('right');
            });
            $img->text(count($podes3->whereIn('R1001C1', [1, 2])) . ' dari ' . count($podes3) . ' Desa', 775, 1030, function ($font) use ($color) {
                $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
                $font->size(25);
                $font->color($color['orange']);
                $font->align('right');
            });
            $p1 = explode('\n', count($podes3->where('R1001B1', 1)) . ' dari ' . count($podes3) . ' Desa di Kecamatan \n Wonomerto sudah menggunakan jalan \n jenis Aspal/Beton.');
            for ($i = 0; $i < count($p1); $i++) {
                $offset = 1185 + ($i * 55);
                $img->text($p1[$i], 510, $offset, function ($font) use ($color) {
                    $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
                    $font->size(20);
                    $font->color($color['secondary']);
                    $font->align('right');
                });
            }
            $p2 = explode('\n', 'Ada ' . count($podes3->whereIn('R1001C1', [1, 2])) . ' dari ' . count($podes3) . ' Desa di Kecamatan \n Wonomerto yang sudah dilalui kendaraan \n umum.');
            for ($i = 0; $i < count($p1); $i++) {
                $offset = 1185 + ($i * 55);
                $img->text($p2[$i], 570, $offset, function ($font) use ($color) {
                    $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
                    $font->size(20);
                    $font->color($color['secondary']);
                    $font->align('left');
                });
            }

            for ($i = 0; $i < ceil($percent1 * 360); $i += 8) {
                $img->circle(6, cos(($i - 90) * pi() / 180) * 76 + 312, sin(($i - 90) * pi() / 180) * 76 + 899, function ($draw) use ($color) {
                    $draw->background($color['blue']);
                });
            }
            for ($i = 0; $i < ceil($percent2 * 360); $i += 8) {
                $img->circle(6, cos(($i - 90) * pi() / 180) * 76 + 737, sin(($i - 90) * pi() / 180) * 76 + 899, function ($draw) use ($color) {
                    $draw->background($color['orange']);
                });
            }
            $img->save('storage/' . $kec . '/9.png');
            //Bab 9 PODES

            //Bab 6 Infografis
            $img = Image::make('template_image/6.png');
            $img->text(number_format(($infografis[0]->Column_6l + $infografis[0]->Column_6p), 0, '.', ' ') . ' jiwa', 350, 695, function ($font) use ($color) {
                $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
                $font->size(70);
                $font->color($color['primary']);
            });
            $img->text(number_format($infografis[0]->Column_6l, 0, '.', ' '), 184, 805, function ($font) use ($color) {
                $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
                $font->size(30);
                $font->color($color['orange']);
                $font->align('center');
            });
            $img->text(number_format($infografis[0]->Column_6p, 0, '.', ' '), 872, 805, function ($font) use ($color) {
                $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
                $font->size(30);
                $font->color($color['blue']);
                $font->align('center');
            });
            $img->text(round($infografis[0]->Column_6l / $infografis[0]->Column_6p * 100), 420, 1080, function ($font) use ($color) {
                $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
                $font->size(50);
                $font->color($color['primary']);
                $font->align('center');
            });
            $img->text(number_format(round(($infografis[0]->Column_6l + $infografis[0]->Column_6p) / $infografis[0]->Column_6luas), 0, '.', ' '), 667, 1080, function ($font) use ($color) {
                $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
                $font->size(50);
                $font->color($color['primary']);
                $font->align('center');
            });

            $p1 = explode('\n', 'Penduduk Kecamatan ' . $infografis[0]->R103N . ' \n berdasarkan data Dinas Dukcapil\n Kabupaten Probolinggo 2021 adalah\n'
                . number_format(($infografis[0]->Column_6l + $infografis[0]->Column_6p), 0, '.', ' ') . ' jiwa');
            for ($i = 0; $i < count($p1); $i++) {
                $offset = 300 + ($i * 55);
                $img->text($p1[$i], 510, $offset, function ($font) use ($color) {
                    $font->file(public_path('assets/font/Proxima Nova Regular.otf'));
                    $font->size(22);
                    $font->color($color['primary']);
                    $font->align('right');
                });
            }

            $p2 = explode('\n', ' Rasio jenis kelamin penduduk \n Kecamatan ' . $infografis[0]->R103N . ' adalah ' .
                round($infografis[0]->Column_6l / $infografis[0]->Column_6p * 100) .
                ', yang \n berarti setiap 100 perempuan terdapat \n sekitar ' .
                round($infografis[0]->Column_6l / $infografis[0]->Column_6p * 100) . ' laki-laki,');
            for ($i = 0; $i < count($p2); $i++) {
                $offset = 300 + ($i * 55);
                $img->text($p2[$i], 530, $offset, function ($font) use ($color) {
                    $font->file(public_path('assets/font/Proxima Nova Regular.otf'));
                    $font->size(22);
                    $font->color($color['primary']);
                    $font->align('left');
                });
            }

            $img->text('Kecamatan ' . $infografis[0]->R103N . ', 2021', 462, 581, function ($font) use ($color) {
                $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
                $font->size(25);
                $font->color($color['black']);
            });
            $img->save('storage/' . $kec . '/6.png');
            //Bab 6 Infografis

            //Bab 3
            $img = Image::make('template_image/3.png');
            $p1 = explode('\n', 'Produksi Padi di ' . $infografis[0]->R103N . ' tahun 2021 adalah ' .
                number_format($infografis[0]->Column_3padi_produksi, 0, '.', ' ')  . ' ton.\nDengan total luas panen mencapai ' .
                number_format($infografis[0]->Column_3padi_luas, 0, '.', ' ') . ' Ha.');
            for ($i = 0; $i < count($p1); $i++) {
                $offset = 905 + ($i * 30);
                $img->text($p1[$i], 304, $offset, function ($font) use ($color) {
                    $font->file(public_path('assets/font/Proxima Nova Regular.otf'));
                    $font->size(20);
                    $font->color($color['primary']);
                });
            }

            $p1 = explode('\n', 'Produksi Jagung di ' . $infografis[0]->R103N . ' tahun 2021 adalah 12 763,18 ton.\nDengan total luas panen mencapai 2 239 Ha.');
            for ($i = 0; $i < count($p1); $i++) {
                $offset = 1090 + ($i * 30);
                $img->text($p1[$i], 304, $offset, function ($font) use ($color) {
                    $font->file(public_path('assets/font/Proxima Nova Regular.otf'));
                    $font->size(20);
                    $font->color($color['primary']);
                });
            }

            $p1 = explode('\n', 'Produksi Jagung di ' . $infografis[0]->R103N . ' tahun 2021 adalah 12 763,18 ton.\nDengan total luas panen mencapai 2 239 Ha.');
            for ($i = 0; $i < count($p1); $i++) {
                $offset = 1277 + ($i * 30);
                $img->text($p1[$i], 304, $offset, function ($font) use ($color) {
                    $font->file(public_path('assets/font/Proxima Nova Regular.otf'));
                    $font->size(20);
                    $font->color($color['primary']);
                });
            }

            $img->save('storage/' . $kec . '/3.png');
            //Bab 3
        }

        //Bab 5
        $img = Image::make('template_image/5.png');
        $img->text('99', 610, 390, function ($font) use ($color) {
            $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
            $font->size(98);
            $font->color($color['blue']);
        });
        $img->text('999', 762, 815, function ($font) use ($color) {
            $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
            $font->size(110);
            $font->color($color['blue']);
        });
        $img->text('99', 127, 874, function ($font) use ($color) {
            $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
            $font->size(130);
            $font->color($color['primary']);
        });
        $img->text('14 Desa, 82 Rukun Warga (RW) dan 288 Rukun Tetangga (RT).', 552, 1135, function ($font) use ($color) {
            $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
            $font->size(24);
            $font->color($color['primary']);
            $font->align('center');
        });
        $img->text('Wonomerto', 719, 1078, function ($font) use ($color) {
            $font->file(public_path('assets/font/Proxima Nova Alt Regular.otf'));
            $font->size(25);
            $font->color($color['primary']);
        });
        $img->save('template_image/5 copy.png');
        //Bab 5

        //Bab 10
        $img = Image::make('template_image/10.png');
        $img->text('Rp 99 999 999 999', 504, 1040, function ($font) use ($color) {
            $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
            $font->size(47);
            $font->color($color['primary']);
            $font->align('center');
        });
        $img->text('Sokaan, Kedungcaluk, dan Opo - Opo', 504, 1190, function ($font) use ($color) {
            $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
            $font->size(28);
            $font->color($color['primary']);
            $font->align('center');
        });
        $img->text('Jumlah total dana desa dari 17 Desa di Kecamatan Krejengan', 504, 925, function ($font) use ($color) {
            $font->file(public_path('assets/font/Proxima Nova Alt Regular.otf'));
            $font->size(24);
            $font->color($color['primary']);
            $font->align('center');
        });
        $img->save('template_image/10 copy.png');
        //Bab 10

        //Bab 4
        $img = Image::make('template_image/4.png');
        $p1 = explode('\n', 'Jumlah Wisatawan di Agro Strawberry Jetak ds;dkas;dlka,\nKecamatan Sukapura Tahun 2020');
        for ($i = 0; $i < count($p1); $i++) {
            $offset = 750 + ($i * 40);
            $img->text($p1[$i], 550, $offset, function ($font) use ($color, $i) {
                if ($i == 0)
                    $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
                else
                    $font->file(public_path('assets/font/Proxima Nova Regular.otf'));

                $font->size(30 - ($i * 5));
                $font->color($color['primary']);
                $font->align('center');
            });
        }

        $img->text('999 999', 560, 915, function ($font) use ($color) {
            $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
            $font->size(90);
            $font->color($color['primary']);
            $font->align('center');
        });

        $img->text('orang', 560, 975, function ($font) use ($color) {
            $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
            $font->size(25);
            $font->color($color['primary']);
            $font->align('center');
        });
        $img->text('999 999', 318, 1076, function ($font) use ($color) {
            $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
            $font->size(20);
            $font->color($color['primary']);
            $font->align('center');
        });
        $img->text('djas;das d;kasd;l kasdasdk ', 706, 1080, function ($font) use ($color) {
            $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
            $font->size(20);
            $font->color($color['primary']);
            $font->align('left');
        });
        $img->text('999 999', 418, 1115, function ($font) use ($color) {
            $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
            $font->size(20);
            $font->color($color['primary']);
            $font->align('center');
        });
        $img->text('999 999', 377, 1152, function ($font) use ($color) {
            $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
            $font->size(20);
            $font->color($color['primary']);
            $font->align('right');
        });

        $img->save('template_image/4 copy.png');
        //Bab 4

        return 'done';
    }
}
