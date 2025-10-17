<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Работа с Excel</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
</head>
<body>
    <div class="container mt-4">
        <div class="row">
            <div class="col-12">
                <h1 class="mb-4">
                    <i class="fas fa-file-excel text-success"></i>
                    Работа с Excel файлами
                </h1>
            </div>
        </div>

        <div class="row">
            <!-- Экспорт -->
            <div class="col-md-6 mb-4">
                <div class="card">
                    <div class="card-header bg-success text-white">
                        <h5 class="card-title mb-0">
                            <i class="fas fa-download"></i> Экспорт данных
                        </h5>
                    </div>
                    <div class="card-body">
                        <p>Экспортировать данные пользователей:</p>
                        <select id="selectFiles">
                            @foreach ($files as $index => $file)
                                <option value="{{ $file }}">{{ ($index + 1) . ' ' . implode(".", array_slice(explode(".", $file),1)) }}</option>
                            @endforeach
                        </select>
                        <a id="exportWord" href="#" class="btn btn-success">
                            <i class="fas fa-file-export"></i> Экспорт в Word
                        </a>
                        <a id="exportPdf" href="#" class="btn btn-success">
                            <i class="fas fa-file-export"></i> Экспорт в Pdf
                        </a>
                    </div>
                </div>
            </div>

            <!-- Импорт -->
            <div class="col-md-6 mb-4">
                <div class="card">
                    <div class="card-header bg-primary text-white">
                        <h5 class="card-title mb-0">
                            <i class="fas fa-upload"></i> Загрузка данных
                        </h5>
                    </div>
                    <div class="card-body">
                        <p>Загрузить данные из Excel файла</p>
                        <a href="{{ route('excel.template') }}" class="btn btn-outline-primary btn-sm mb-3">
                            <i class="fas fa-file-download"></i> Скачать шаблон
                        </a>

                        <form id="importForm" enctype="multipart/form-data">
                            @csrf
                            <div class="mb-3">
                                <label for="excel_file" class="form-label">Выберите Excel файл:</label>
                                <input type="file" class="form-control" id="excel_file" name="excel_file"
                                       accept=".xlsx,.xls,.csv" required>
                            </div>
                            <button type="submit" class="btn btn-primary" id="importBtn">
                                <i class="fas fa-upload"></i> Загрузить
                            </button>
                        </form>
                    </div>
                </div>
            </div>
        </div>

        <!-- Предпросмотр -->
        <div class="row d-none" id="previewSection">
            <div class="col-12">
                <div class="card">
                    <div class="card-header bg-info text-white">
                        <h5 class="card-title mb-0">
                            <i class="fas fa-eye"></i> Предпросмотр данных
                        </h5>
                    </div>
                    <div class="card-body">
                        <div id="previewTable"></div>
                    </div>
                </div>
            </div>
        </div>

    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
    <script>
        // Предпросмотр файла при выборе
        document.getElementById('excel_file').addEventListener('change', function(e) {
            const file = e.target.files[0];
            if (!file) return;

            const formData = new FormData();
            formData.append('excel_file', file);
            formData.append('_token', '{{ csrf_token() }}');

            // Показываем загрузку
            document.getElementById('previewSection').classList.add('d-none');
            document.getElementById('previewTable').innerHTML = '<div class="text-center"><div class="spinner-border"></div></div>';

            fetch('{{ route("excel.preview") }}', {
                method: 'POST',
                body: formData
            })
            .then(response => response.json())
            .then(data => {
                if (data.success) {
                    showPreview(data.preview);
                    document.getElementById('previewSection').classList.remove('d-none');
                } else {
                    alert(data.message);
                }
            })
            .catch(error => {
                console.error('Error:', error);
                alert('Ошибка при предпросмотре файла');
            });
        });

        // Функция показа предпросмотра
        function showPreview(data) {
            if (data.length === 0) {
                document.getElementById('previewTable').innerHTML = '<p>Файл пуст</p>';
                return;
            }

            let html = '<div class="table-responsive"><table class="table table-bordered table-sm">';

            // Заголовки
            html += '<thead class="table-light"><tr>';
            data[0].forEach((cell, index) => {
                html += `<th>Колонка ${index + 1}</th>`;
            });
            html += '</tr></thead>';

            // Данные
            html += '<tbody>';
            data.forEach(row => {
                html += '<tr>';
                row.forEach(cell => {
                    html += `<td>${cell || ''}</td>`;
                });
                html += '</tr>';
            });
            html += '</tbody></table></div>';

            document.getElementById('previewTable').innerHTML = html;
        }

        const selectFiles = document.getElementById('selectFiles');
        selectFiles.addEventListener('click', (e) => {
            if (e.target.value != '') {
                const exportWord = "{{ route('excel.exportWord') }}";
                document.getElementById('exportWord').href = exportWord + "?file=" + e.target.value;
                const exportPdf = "{{ route('excel.exportPdf') }}";
                document.getElementById('exportPdf').href = exportPdf + "?file=" + e.target.value;
            }
        });

        function showFiles(files) {
            selectFiles.innerHTML = '';
            files.forEach((file, index) => {
                selectFiles.appendChild(new Option((index+1) + ' ' + file.split('.').slice(1).join('.'), file));
            })
            selectFiles.click();
        }

        // Обработка импорта
        document.getElementById('importForm').addEventListener('submit', function(e) {
            e.preventDefault();

            const formData = new FormData(this);
            const importBtn = document.getElementById('importBtn');

            // Блокируем кнопку
            importBtn.disabled = true;
            importBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Загрузка...';

            fetch('{{ route("excel.import") }}', {
                method: 'POST',
                body: formData
            })
            .then(response => response.json())
            .then(data => {
                if (!data?.success) {
                    alert('Ошибка загрузки')
                } else {
                    showFiles(data.files);
                }
                importBtn.disabled = false;
                importBtn.innerHTML = '<i class="fas fa-upload"></i> Загрузить';
            })
            .catch(error => {
                console.error('Error:', error);
                alert('Ошибка при загрузке файла');
                importBtn.disabled = false;
                importBtn.innerHTML = '<i class="fas fa-upload"></i> Загрузить';
            });
        });
    </script>
</body>
</html>
