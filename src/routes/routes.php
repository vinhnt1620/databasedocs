<?php

use Illuminate\Support\Facades\Artisan;
use Vinhnt\Databasedocs\Exports\TablesExport;
use Maatwebsite\Excel\Facades\Excel;
use Illuminate\Support\Facades\File;

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

Artisan::command('generate:databasedocs', function () {
    $path = storage_path('app/databasedocs/');
    if(!File::isDirectory($path)){
        File::makeDirectory($path);
    }   

    $this->call('generate:erd', ['filename' => $path.'erd.png']);

    Excel::store(new TablesExport(), 'tables.xlsx', 'local');
})->purpose('Generate database document');

