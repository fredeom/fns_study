<?php


use App\Http\Controllers\ExcelController;
use App\Http\Controllers\FileUploadController;
use Illuminate\Support\Facades\Route;

Route::get('/', function () {
    return redirect('/excel');
});

Route::prefix('excel')->group(function () {
    Route::get('/', [ExcelController::class, 'index'])->name('excel.index');
    Route::get('/export-word', [ExcelController::class, 'exportWord'])->name('excel.exportWord');
    Route::get('/export-pdf', [ExcelController::class, 'exportPdf'])->name('excel.exportPdf');
    Route::get('/template', [ExcelController::class, 'downloadTemplate'])->name('excel.template');
    Route::post('/import', [ExcelController::class, 'import'])->name('excel.import');
    Route::post('/preview', [ExcelController::class, 'preview'])->name('excel.preview');
});
