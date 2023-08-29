<?php

namespace App\Http\Controllers;

use App\Models\Infografis;
use App\Models\Podes1;
use App\Models\Podes2;
use App\Models\Podes3;
use Exception;
use Illuminate\Support\Facades\Storage;
use Intervention\Image\Facades\Image;
use Illuminate\Support\Str;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

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

        $kecamatan = ['10', '20', '30', '40', '50', '60', '70', '80', '90', '100', '110', '120', '130', '140', '150', '160', '170', '180', '190', '200', '210', '220', '230', '240'];
        // $kecamatan = ['10', '20', '30'];

        foreach ($kecamatan as $kec) {

            $podes1 = Podes1::where(['R103' => $kec])->get();
            $podes2 = Podes2::where(['R103' => $kec])->get();
            $podes3 = Podes3::where(['R103' => $kec])->get();
            $infografis = Infografis::where(['R103' => $kec])->get();

            Storage::makeDirectory('public/' . $kec . '_' . $podes1[0]->R103N);

            //Bab 4 PODES BARU
            $img = Image::make('template_image/4.png');
            $img->text(ucwords(Str::lower($podes2[0]->R103N)), 725, 724, function ($font) use ($color) {
                $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
                $font->size(22);
                $font->color($color['black']);
            });
            $img->text(ucwords(Str::lower($podes2[0]->R103N)), 718, 1121, function ($font) use ($color) {
                $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
                $font->size(22);
                $font->color($color['black']);
            });
            $img->text(ucwords(Str::lower($podes2[0]->R103N)), 718, 1121, function ($font) use ($color) {
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
            $img->save('storage/' . $kec . '_' . $podes1[0]->R103N . '/4.png');
            //Bab 4 PODES BARU

            //Bab 7 PODES
            // $img = Image::make('template_image/7.png');
            // $img->text(($infografis[0]->wisata_name != null ? 7 : 6) . '.', 354, 187, function ($font) use ($color) {
            //     $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
            //     $font->size(48);
            //     $font->color($color['tertiary']);
            //     $font->align('center');
            // });
            // $img->text(number_format($podes1->sum('R501A1'), 0, '.', ' '), 718, 442, function ($font) use ($color) {
            //     $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
            //     $font->size(45);
            //     $font->color($color['tertiary']);
            //     $font->align('center');
            // });
            // $img->text(count($podes1->where('R503B', 4)), 310, 1043, function ($font) use ($color) {
            //     $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
            //     $font->size(24);
            //     $font->color($color['primary']);
            //     $font->align('center');
            // });
            // $img->text(count($podes1), 405, 1043, function ($font) use ($color) {
            //     $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
            //     $font->size(24);
            //     $font->color($color['primary']);
            //     $font->align('center');
            // });
            // $img->text(count($podes1->whereIn('R507A', [3, 4])), 687, 1043, function ($font) use ($color) {
            //     $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
            //     $font->size(24);
            //     $font->color($color['primary']);
            //     $font->align('center');
            // });
            // $img->text(count($podes1), 785, 1043, function ($font) use ($color) {
            //     $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
            //     $font->size(24);
            //     $font->color($color['primary']);
            //     $font->align('center');
            // });
            // $img->save('storage/' . $kec . '_' . $podes1[0]->R103N . '/' . ($infografis[0]->wisata_name != null ? 7 : 6) . '.png');
            //Bab 7 PODES

            //Bab 7 PODES BARU
            $img = Image::make('template_image/7.png');

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
            $img->save('storage/' . $kec . '_' . $podes1[0]->R103N . '/7.png');
            //Bab 7 PODES BARU

            //Bab 6 PODES BARU
            $img = Image::make('template_image/6.png');

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
            $p1 = explode('\n', count($podes3->where('R1001B1', 1)) . ' dari ' . count($podes3) . ' Desa di Kecamatan \n ' . ucwords(Str::lower($podes3[0]->R103N)) . ' sudah menggunakan jalan \n jenis Aspal/Beton.');
            for ($i = 0; $i < count($p1); $i++) {
                $offset = 1185 + ($i * 55);
                $img->text($p1[$i], 510, $offset, function ($font) use ($color) {
                    $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
                    $font->size(20);
                    $font->color($color['secondary']);
                    $font->align('right');
                });
            }
            $p2 = explode('\n', 'Ada ' . count($podes3->whereIn('R1001C1', [1, 2])) . ' dari ' . count($podes3) . ' Desa di Kecamatan \n ' . ucwords(Str::lower($podes3[0]->R103N)) . ' yang sudah dilalui kendaraan \n umum.');
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
            $img->save('storage/' . $kec . '_' . $podes1[0]->R103N . '/6.png');
            //Bab 6 PODES BARU

            //Bab 3 Infografis BARU
            $img = Image::make('template_image/3.png');
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
            $img->save('storage/' . $kec . '_' . $podes1[0]->R103N . '/3.png');
            //Bab 6 Infografis BARU

            //Bab 3
            // $img = Image::make('template_image/3.png');
            // $p1 = explode('\n', 'Produksi Padi di ' . $infografis[0]->R103N . ' tahun 2021 adalah ' .
            //     number_format($infografis[0]->Column_3padi_produksi, 0, '.', ' ')  . ' ton.\nDengan total luas panen mencapai ' .
            //     number_format($infografis[0]->Column_3padi_luas, 0, '.', ' ') . ' Ha.');
            // for ($i = 0; $i < count($p1); $i++) {
            //     $offset = 905 + ($i * 30);
            //     $img->text($p1[$i], 304, $offset, function ($font) use ($color) {
            //         $font->file(public_path('assets/font/Proxima Nova Regular.otf'));
            //         $font->size(20);
            //         $font->color($color['primary']);
            //     });
            // }

            // $p1 = explode('\n', 'Komoditas sayuran yang paling banyak diproduksi di Kecamatan \n' . $infografis[0]->R103N . ' tahun 2021 adalah ' .
            //     $infografis[0]->horti .
            //     ' dengan produksi \nsebesar ' . number_format($infografis[0]->produksi_horti, 0, '.', ' ') .
            //     ' kuintal dan total luas panen mencapai ' . number_format($infografis[0]->luas_horti, 0, '.', ' ') .
            //     ' Ha.');
            // for ($i = 0; $i < count($p1); $i++) {
            //     $offset = 1090 + ($i * 30);
            //     $img->text($p1[$i], 304, $offset, function ($font) use ($color) {
            //         $font->file(public_path('assets/font/Proxima Nova Regular.otf'));
            //         $font->size(20);
            //         $font->color($color['primary']);
            //     });
            // }

            // $p1 = explode('\n', 'Jumlah Sapi Potong di Kecamatan ' . $infografis[0]->R103N .
            //     ' tahun 2021 adalah \n' .
            //     number_format($infografis[0]->ternak, 0, '.', ' ') . ' ekor');
            // for ($i = 0; $i < count($p1); $i++) {
            //     $offset = 1277 + ($i * 30);
            //     $img->text($p1[$i], 304, $offset, function ($font) use ($color) {
            //         $font->file(public_path('assets/font/Proxima Nova Regular.otf'));
            //         $font->size(20);
            //         $font->color($color['primary']);
            //     });
            // }

            // $img->save('storage/' . $kec . '_' . $podes1[0]->R103N . '/3.png');
            //Bab 3

            //Bab 4
            // if ($infografis[0]->wisata_name != null) {
            //     $img = Image::make('template_image/4.png');
            //     $p1 = explode('\n', 'Jumlah Wisatawan ' . ($infografis[0]->wisata_name) . '\nKecamatan ' . ($infografis[0]->R103N) . ' Tahun 2021');
            //     for ($i = 0; $i < count($p1); $i++) {
            //         $offset = 750 + ($i * 40);
            //         $img->text($p1[$i], 550, $offset, function ($font) use ($color, $i) {
            //             if ($i == 0)
            //                 $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
            //             else
            //                 $font->file(public_path('assets/font/Proxima Nova Regular.otf'));

            //             $font->size(30 - ($i * 5));
            //             $font->color($color['primary']);
            //             $font->align('center');
            //         });
            //     }

            //     $img->text((number_format($infografis[0]->wisata_total, 0, '.', ' ')), 560, 915, function ($font) use ($color) {
            //         $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
            //         $font->size(90);
            //         $font->color($color['primary']);
            //         $font->align('center');
            //     });

            //     $img->text('orang', 560, 975, function ($font) use ($color) {
            //         $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
            //         $font->size(25);
            //         $font->color($color['primary']);
            //         $font->align('center');
            //     });
            //     $img->text(number_format($infografis[0]->wisata_total, 0, '.', ' '), 318, 1076, function ($font) use ($color) {
            //         $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
            //         $font->size(20);
            //         $font->color($color['primary']);
            //         $font->align('center');
            //     });
            //     $img->text($infografis[0]->wisata_name, 706, 1077, function ($font) use ($color) {
            //         $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
            //         $font->size(20);
            //         $font->color($color['primary']);
            //         $font->align('left');
            //     });

            //     $img->save('storage/' . $kec . '_' . $podes1[0]->R103N . '/4.png');
            // }

            //Bab 4

            //Bab 2 Infografis BARU
            $img = Image::make('template_image/2.png');

            $img->text($infografis[0]->total_rw, 630, 375, function ($font) use ($color) {
                $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
                $font->size(110);
                $font->color($color['blue']);
            });
            $img->text($infografis[0]->total_rt, 762, 815, function ($font) use ($color) {
                $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
                $font->size(100);
                $font->color($color['blue']);
            });
            $img->text($infografis[0]->total_village, 127, 874, function ($font) use ($color) {
                $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
                $font->size(130);
                $font->color($color['primary']);
            });
            $img->text($infografis[0]->total_village . ' Desa, ' . $infografis[0]->total_rw . ' Rukun Warga (RW) dan ' . $infografis[0]->total_rt . ' Rukun Tetangga (RT).', 552, 1135, function ($font) use ($color) {
                $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
                $font->size(24);
                $font->color($color['primary']);
                $font->align('center');
            });
            $img->text($infografis[0]->R103N . ' memiliki, ', 719, 1078, function ($font) use ($color) {
                $font->file(public_path('assets/font/Proxima Nova Alt Regular.otf'));
                $font->size(25);
                $font->color($color['primary']);
            });
            $img->text($infografis[0]->R103N, 721, 268, function ($font) use ($color) {
                $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
                $font->size(23);
                $font->color($color['primary']);
            });
            $img->save('storage/' . $kec . '_' . $podes1[0]->R103N . '/2.png');
            //Bab 2 Infografis BARU

            //Bab 10
            // $img = Image::make('template_image/10.png');
            // $img->text('BAB ' . ($infografis[0]->wisata_name != null ? 10 : 9) . '. ' . ($infografis[0]->is_price_avail ? 'KEUANGAN DAERAH DAN HARGA' : 'KEUANGAN DAERAH'), 518, 123, function ($font) use ($color, $infografis) {
            //     $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
            //     $font->size(($infografis[0]->is_price_avail ?  32 : 40));
            //     $font->color($color['tertiary']);
            //     $font->align('center');
            // });
            // $img->text('Rp ' . number_format($infografis[0]->total_dd, 0, ',', ' '), 504, 1040, function ($font) use ($color) {
            //     $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
            //     $font->size(47);
            //     $font->color($color['primary']);
            //     $font->align('center');
            // });
            // $img->text($infografis[0]->biggest_dd, 504, 1190, function ($font) use ($color) {
            //     $font->file(public_path('assets/font/Proxima Nova Alt Bold.otf'));
            //     $font->size(28);
            //     $font->color($color['primary']);
            //     $font->align('center');
            // });
            // $img->text('Jumlah total dana desa dari ' . $infografis[0]->total_village . ' Desa di Kecamatan ' . $infografis[0]->R103N, 504, 925, function ($font) use ($color) {
            //     $font->file(public_path('assets/font/Proxima Nova Alt Regular.otf'));
            //     $font->size(24);
            //     $font->color($color['primary']);
            //     $font->align('center');
            // });
            // $img->save('storage/' . $kec . '_' . $podes1[0]->R103N . '/' . ($infografis[0]->wisata_name != null ? 10 : 9) . '.png');
            //Bab 10
        }

        return 'done';
    }

    public function generateData()
    {
        //data peternakan
        $peternakan = \PhpOffice\PhpSpreadsheet\IOFactory::load('data/peternakan.xlsx');
        $peternakanresult = array();

        $sheetnames = $peternakan->getSheetNames();
        for ($i = 0; $i < $peternakan->getSheetCount(); $i++) {
            $sheet = $peternakan->getSheet($i);
            $peternakanresult[$sheetnames[$i]] = $sheet->getCell('E7')->getOldCalculatedValue();
        }
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $i = 1;
        foreach ($peternakanresult as $key => $value) {
            $sheet->getCell('A' . $i)
                ->setValue($key);
            $sheet->getCell('B' . $i)
                ->setValue($value);
            $i++;
        }

        $writer = new Xlsx($spreadsheet);
        $writer->save('data/peternakan result.xlsx');
        //data peternakan

        //data hortikultura
        $hortikultura = \PhpOffice\PhpSpreadsheet\IOFactory::load('data/hortikultura.xlsx');
        $hortikulturaresult = array();

        $startrow = 9;
        $sheetnames = $hortikultura->getSheetNames();

        for ($i = 0; $i < $hortikultura->getSheetCount(); $i++) {
            $sheet = $hortikultura->getSheet($i);
            $value = $sheet->rangeToArray('G' . $startrow . ':G34');
            $valueclean = array();
            for ($j = 0; $j < count($value); $j++) {
                $valueclean[] = (int) str_replace('-', 0, str_replace(' ', '', $value[$j][0]));
            }
            $hortikulturaresult[str_replace('KEC. ', '', $sheetnames[$i])] = [
                'jenis' => $sheet->getCell('B' . (array_keys($valueclean, max($valueclean))[0] + $startrow))->getValue(),
                'produksi' => max($valueclean),
                'luas' => str_replace(' ', '', $sheet->getCell('E' . (array_keys($valueclean, max($valueclean))[0] + $startrow))->getValue())
            ];
        }
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $i = 1;
        foreach ($hortikulturaresult as $key => $value) {
            $sheet->getCell('A' . $i)
                ->setValue($key);
            $sheet->getCell('B' . $i)
                ->setValue($value['jenis']);
            $sheet->getCell('C' . $i)
                ->setValue($value['produksi']);
            $sheet->getCell('D' . $i)
                ->setValue($value['luas']);
            $i++;
        }

        $writer = new Xlsx($spreadsheet);
        $writer->save('data/hortikultura result.xlsx');

        //data hortikultura

        //data pariwisata

        $pariwisata = \PhpOffice\PhpSpreadsheet\IOFactory::load('data/pariwisata.xlsx');
        $pariwisataresult = array();

        $startrow = 9;
        $sheetnames = $pariwisata->getSheetNames();

        for ($i = 0; $i < $pariwisata->getSheetCount(); $i++) {
            $sheet = $pariwisata->getSheet($i);

            $row = 20;
            $wisataarray = collect();
            do {
                $strpos = strpos(str_replace('Jumlah Wisatawan Domestik dan Asing di ', '', $sheet->getCell('C' . ($row - 18))->getValue()), ' di ');

                $name = substr(str_replace('Jumlah Wisatawan Domestik dan Asing di ', '', $sheet->getCell('C' . ($row - 18))->getValue()), 0, $strpos);
                $total = (int)$sheet->getCell('E' . $row)->getOldCalculatedValue();
                $dom = (int)$sheet->getCell('C' . $row)->getOldCalculatedValue();
                $for = (int)$sheet->getCell('D' . $row)->getOldCalculatedValue();

                $row += 21;

                if ($total == null) {
                    break;
                }
                $wisataarray[] = [
                    'name' => $name,
                    'total' => $total,
                    'dom' => $dom,
                    'for' => $for
                ];
            } while (true);

            $wisataarray = $wisataarray->sortBy('total');

            $pariwisataresult[str_replace(' ', '', str_replace('KEC.', '', $sheetnames[$i]))] = $wisataarray->last();
        }

        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $i = 1;
        foreach ($pariwisataresult as $key => $value) {
            $sheet->getCell('A' . $i)
                ->setValue($key);
            $sheet->getCell('B' . $i)
                ->setValue($value['name']);
            $sheet->getCell('C' . $i)
                ->setValue($value['total']);
            $sheet->getCell('D' . $i)
                ->setValue($value['dom']);
            $sheet->getCell('E' . $i)
                ->setValue($value['for']);
            $i++;
        }

        $writer = new Xlsx($spreadsheet);
        $writer->save('data/pariwisata result.xlsx');
        //data pariwisata

        //data dana desa
        $danadesa = \PhpOffice\PhpSpreadsheet\IOFactory::load('data/danadesa.xlsx');
        $danadesaresult = array();

        $startrow = 6;
        $sheetnames = $danadesa->getSheetNames();

        for ($i = 0; $i < $danadesa->getSheetCount(); $i++) {
            $sheet = $danadesa->getSheet($i);
            $villagenum = count(Podes1::where(['R103N' => $sheetnames[$i]])->get());

            $danadesavalue = $sheet->rangeToArray('B' . $startrow . ':C' . ($startrow + $villagenum - 1));

            $endrow = $startrow + $villagenum;
            if ($sheetnames[$i] == 'Kraksaan') {
                $endrow = $startrow + 13;
            }

            $totaldanadesa = $sheet->getCell('C' . $endrow)->getOldCalculatedValue();

            $danadesavalueclean = array();
            for ($j = 0; $j < count($danadesavalue); $j++) {
                $danadesavalueclean[$danadesavalue[$j][0]] = (float) str_replace(',', 0, str_replace(' ', '', $danadesavalue[$j][1]));
            }
            arsort($danadesavalueclean);
            $biggest = array_slice($danadesavalueclean, 0, 3);

            $last  = array_slice(array_keys($biggest), -1);
            $first = join(', ', array_slice(array_keys($biggest), 0, -1));
            $both  = array_filter(array_merge(array($first), $last), 'strlen');
            $biggeststring = join(' dan ', $both);

            $danadesaresult[str_replace('KEC. ', '', $sheetnames[$i])] = [
                'total' => $totaldanadesa,
                'biggest_desa' => $biggeststring,
            ];
        }

        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $i = 1;
        foreach ($danadesaresult as $key => $value) {
            $sheet->getCell('A' . $i)
                ->setValue($key);
            $sheet->getCell('B' . $i)
                ->setValue($value['total']);
            $sheet->getCell('C' . $i)
                ->setValue($value['biggest_desa']);
            $i++;
        }

        $writer = new Xlsx($spreadsheet);
        $writer->save('data/dana desa result.xlsx');

        //data dana desa


        return 'done';
    }
}
