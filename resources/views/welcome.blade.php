<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>File Upload</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body>
    <div class="container mt-5">
        <div class="row justify-content-center">
            <div class="col-md-6">
                <div class="card">
                    <div class="card-header">
                        <h4>Загрузка xls/xlsx файлов</h4>
                    </div>
                    <div class="card-body">
                        @if(session('success'))
                            <div class="alert alert-success">
                                {{ session('success') }}
                            </div>
                        @endif

                        @if($errors->any())
                            <div class="alert alert-danger">
                                <ul class="mb-0">
                                    @foreach($errors->all() as $error)
                                        <li>{{ $error }}</li>
                                    @endforeach
                                </ul>
                            </div>
                        @endif

                        <form action="{{ route('upload.file') }}" method="POST" enctype="multipart/form-data">
                            @csrf

                            <div class="mb-3">
                                <label for="file" class="form-label">Выберите файл</label>
                                <input type="file" class="form-control @error('file') is-invalid @enderror"
                                       id="file" name="file" required>
                                @error('file')
                                    <div class="invalid-feedback">{{ $message }}</div>
                                @enderror
                            </div>

                            <button type="submit" class="btn btn-primary">Загрузить файл</button>
                        </form>
                    </div>
                </div>

                @if(isset($files) && $files->count() > 0)
                <div class="card mt-4">
                    <div class="card-header">
                        <h5>Uploaded Files</h5>
                    </div>
                    <div class="card-body">
                        <div class="list-group">
                            @foreach($files as $file)
                            <div class="list-group-item">
                                <div class="d-flex justify-content-between align-items-center">
                                    <div>
                                        <h6 class="mb-1">{{ $file->original_name }}</h6>
                                        <small class="text-muted">
                                            Uploaded: {{ $file->created_at->format('M d, Y H:i') }}
                                        </small>
                                    </div>
                                    <div>
                                        <a href="{{ Storage::url($file->path) }}"
                                           target="_blank"
                                           class="btn btn-sm btn-outline-primary">View</a>
                                        <form action="{{ route('file.delete', $file->id) }}"
                                              method="POST" class="d-inline">
                                            @csrf
                                            @method('DELETE')
                                            <button type="submit" class="btn btn-sm btn-outline-danger"
                                                    onclick="return confirm('Are you sure?')">Delete</button>
                                        </form>
                                    </div>
                                </div>
                            </div>
                            @endforeach
                        </div>
                    </div>
                </div>
                @endif
            </div>
        </div>
    </div>
</body>
</html>
