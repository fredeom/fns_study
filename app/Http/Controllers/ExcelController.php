<?php

namespace App\Http\Controllers;

use Exception;
use Illuminate\Http\Request;
use Illuminate\Http\JsonResponse;
use App\Models\User; // Пример модели
use Illuminate\Support\Facades\Validator;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use Symfony\Component\HttpFoundation\StreamedResponse;
use Illuminate\Support\Facades\Storage;
use \PhpOffice\PhpWord\PhpWord;
use \Dompdf\Dompdf;
use \Dompdf\Options;

class ExcelController extends Controller
{
    public function index()
    {
        $disk = Storage::disk('public');
        $files = $disk->files('uploads');

        return view('excel.index', compact('files'));
    }

    public function exportWord(Request $request): StreamedResponse {

        $spreadsheet = IOFactory::load(storage_path('app/public/' . $request->query('file')));
        $sheet = $spreadsheet->getActiveSheet();

        $sheet->setCellValue('A1', 'Здесь');
        $sheet->setCellValue('B1', 'Был');
        $sheet->setCellValue('C1', 'Вася');

        $data = $sheet->toArray();

        $htmlContent = '<h1 style="color: #1F497D;">HTML Отчет</h1>
        <p style="color: #2F5496; font-style: italic; margin-bottom: 2rem;">Этот отчет был сгенерирован автоматически</p><p>&nbsp;</p><table>';
        $flag = false;
        foreach ($data as $key => $row) {
            if ($key == 0) {
                $htmlContent .= '<thead><tr style="background-color: #D9D9D9;">';
                foreach ($row as $cell) {
                    $htmlContent .= '<th>' . $cell . '</th>';
                }
                $htmlContent .= '</tr></thead><tbody>';
            } else {
                $htmlContent .= '<tr>';
                foreach ($row as $cell) {
                    $htmlContent .= '<td>' . $cell . '</td>';
                }
                $htmlContent .= '</tr>';
                $flag = true;
            }
        }
        $htmlContent .= ($flag ? '</tbody>' : '') . '</table>';

        $phpWord = new PhpWord();
        $section = $phpWord->addSection();
        $section->addImage(
            storage_path('app/public/images/logo.png'),
            ['width' => 100]
        );

        \PhpOffice\PhpWord\Shared\Html::addHtml($section, $htmlContent);

        return response()->streamDownload(function() use ($phpWord) {
            $writer = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
            $writer->save('php://output');
        }, 'export_' . date('Y-m-d') . '.docx');
    }


    private function createHtmlWithImage($data) {
        $logoPath = storage_path('app/public/images/logo.png');
        $logoBase64 = base64_encode(file_get_contents($logoPath));

        $html = '
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <style>
                body { font-family: DejaVu Sans, sans-serif; }
                .header { text-align: center; margin-bottom: 20px; }
                .logo { height: 80px; }
                .table { width: 100%; border-collapse: collapse; }
                .table th, .table td { border: 1px solid #000; padding: 8px; }
                .table th { background-color: #f2f2f2; }
            </style>
        </head>
        <body>
            <div class="header">
                <img src="data:image/png;base64,' . $logoBase64 . '" class="logo" alt="Logo">
                <h1>ОТЧЕТ ПО ДАННЫМ</h1>
                <p>Дата формирования: ' . date('d.m.Y H:i') . '</p>
            </div>
        ';

        if (!empty($data)) {
            $html .= '<table class="table"><thead><tr>';
            foreach ($data[0] as $header) {
                $html .= '<th>' . htmlspecialchars($header) . '</th>';
            }
            $html .= '</tr></thead><tbody>';

            for ($i = 1; $i < count($data); $i++) {
                $html .= '<tr>';
                foreach ($data[$i] as $cell) {
                    $html .= '<td>' . htmlspecialchars($cell) . '</td>';
                }
                $html .= '</tr>';
            }

            $html .= '</tbody></table>';
        }

        $html .= '</body></html>';

        return $html;
    }

    public function exportPdf(Request $request) {
        $spreadsheet = IOFactory::load(storage_path('app/public/' . $request->query('file')));
        $sheet = $spreadsheet->getActiveSheet();
        $data = $sheet->toArray();

        $html = $this->createHtmlWithImage($data);

        $options = new Options();
        $options->set('isRemoteEnabled', true);
        $options->set('defaultFont', 'DejaVu Sans'); // Поддержка русского

        $dompdf = new Dompdf($options);
        $dompdf->loadHtml($html, 'UTF-8');
        $dompdf->setPaper('A4', 'portrait');
        $dompdf->render();

        $filename = 'export_' . date('Y-m-d') . '.pdf';
        $tempPath = storage_path('app/temp/' . $filename);

        file_put_contents($tempPath, $dompdf->output());

        return response()->download($tempPath, $filename, [
            'Content-Type' => 'application/pdf'
        ])->deleteFileAfterSend(true);
    }

    public function export(): StreamedResponse
    {
        // Создаем новый spreadsheet
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        // Заголовки
        $sheet->setTitle('Пользователи');

        // Устанавливаем заголовки таблицы
        $headers = ['ID', 'Имя', 'Email', 'Дата регистрации', 'Статус'];
        $sheet->fromArray($headers, null, 'A1');

        // Получаем данные из базы
        $users = User::select('id', 'name', 'email', 'created_at', 'active')->get();

        // Заполняем данными
        $data = [];
        foreach ($users as $index => $user) {
            $data[] = [
                $user->id,
                $user->name,
                $user->email,
                $user->created_at->format('d.m.Y H:i'),
                $user->active ? 'Активен' : 'Неактивен'
            ];
        }

        $sheet->fromArray($data, null, 'A2');

        // Стили для заголовков
        $headerStyle = [
            'font' => [
                'bold' => true,
                'color' => ['rgb' => 'FFFFFF']
            ],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['rgb' => '3498DB']
            ],
            'borders' => [
                'allBorders' => [
                    'borderStyle' => Border::BORDER_THIN
                ]
            ],
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER
            ]
        ];

        $sheet->getStyle('A1:E1')->applyFromArray($headerStyle);

        // Авто-ширина колонок
        foreach (range('A', 'E') as $column) {
            $sheet->getColumnDimension($column)->setAutoSize(true);
        }

        // Стили для данных
        $dataStyle = [
            'borders' => [
                'allBorders' => [
                    'borderStyle' => Border::BORDER_THIN
                ]
            ]
        ];

        $lastRow = count($data) + 1;
        $sheet->getStyle("A2:E{$lastRow}")->applyFromArray($dataStyle);

        // Создаем writer и возвращаем файл
        $writer = new Xlsx($spreadsheet);

        return response()->streamDownload(function() use ($writer) {
            $writer->save('php://output');
        }, 'users_' . date('Y-m-d') . '.xlsx');
    }

    /**
     * Импорт данных из Excel
     */
    public function import(Request $request): JsonResponse
    {
        $validator = Validator::make($request->all(), [
           'excel_file' => 'required|file|mimes:xlsx,xls|max:5120'
        ]);

        if ($validator->fails()) {
            return response()->json([
                'message' => 'Ошибка:' . json_encode($validator->errors()->get('excel_file')),
                'success' => false,
            ], 500);
        }

        $file = $request->file('excel_file');
        $filename = time() . '.' . $file->getClientOriginalName();
        $file->storeAs('uploads', $filename, 'public');

        $disk = Storage::disk('public');
        $files = $disk->files('uploads');

        return response()->json([
            'success' => true,
            'files' => $files,
        ]);
    }

    /**
     * Создание шаблона для импорта
     */
    public function downloadTemplate(): StreamedResponse
    {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        // Заголовки
        $headers = ['ID (авто)', 'Имя', 'Email', 'Дата регистрации', 'Статус (активен/неактивен)'];
        $sheet->fromArray($headers, null, 'A1');

        // Пример данных
        $examples = [
            ['', 'Иван Иванов', 'ivan@example.com', date('d.m.Y H:i'), 'активен'],
            ['', 'Петр Петров', 'petr@example.com', date('d.m.Y H:i'), 'неактивен']
        ];

        $sheet->fromArray($examples, null, 'A2');

        // Стили
        $headerStyle = [
            'font' => ['bold' => true],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['rgb' => 'E8F4FD']
            ]
        ];

        $sheet->getStyle('A1:E1')->applyFromArray($headerStyle);

        // Авто-ширина
        foreach (range('A', 'E') as $column) {
            $sheet->getColumnDimension($column)->setAutoSize(true);
        }

        $writer = new Xlsx($spreadsheet);

        return response()->streamDownload(function() use ($writer) {
            $writer->save('php://output');
        }, 'import_template_' . date('Y-m-d') . '.xlsx');
    }

    /**
     * Просмотр данных из Excel файла
     */
    public function preview(Request $request): JsonResponse
    {
        $validator = Validator::make($request->all(), [
           'excel_file' => 'required|file|mimes:xlsx,xls|max:5120'
        ]);

        if ($validator->fails()) {
            return response()->json([
                'message' => 'Ошибка:' . json_encode($validator->errors()->get('excel_file')),
                'success' => false,
            ], 500);
        }

        try {
            $file = $request->file('excel_file');
            $spreadsheet = IOFactory::load($file->getPathname());
            $sheet = $spreadsheet->getActiveSheet();

            $data = $sheet->toArray();
            $previewData = array_slice($data, 0, 10); // Первые 10 строк для предпросмотра

            return response()->json([
                'success' => true,
                'preview' => $previewData,
                'total_rows' => count($data)
            ]);
        } catch (\Exception $e) {
            return response()->json([
                'message' => 'Ошибка при чтении файла: ' . $e->getMessage() . $e->getLine(),
                'success' => false,
            ], 500);
        }
    }
}
