<?php

use Illuminate\Support\Facades\Route;

/*
|--------------------------------------------------------------------------
| Web Routes
|--------------------------------------------------------------------------
|
| Here is where you can register web routes for your application. These
| routes are loaded by the RouteServiceProvider within a group which
| contains the "web" middleware group. Now create something great!
|
*/

Route::get('/', function () {
    return view('welcome');
});

Route::get('/import', [App\Http\Controllers\MainController::class, 'import']);
// Route::get('/generate', [App\Http\Controllers\WordGeneratorController::class, 'generate']);
Route::get('/generate', [App\Http\Controllers\InfografisController::class, 'generateImage']);
Route::get('/generate-data', [App\Http\Controllers\InfografisController::class, 'generateData']);
Route::get('/simdasi', [App\Http\Controllers\SimdasiController::class, 'transform']);
