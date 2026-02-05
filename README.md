# ConvertData

Консольное приложение на .NET 10 для конвертации входных Excel/табличных файлов в JSON.

## Что делает

- Берёт входные файлы из папки `ConvertData/EXCEL/` (относительно каталога проекта).
- Для каждого входного файла формирует JSON с тем же базовым именем.
- Сохраняет результаты в `ConvertData/JSON_OUT/`.

Выходных `.json` файлов получается **столько же**, сколько входных файлов.

## Формат входа

Поддерживаются файлы с расширением `.xls`.

Особенность:
- `H2_1.xls` в этом репозитории фактически содержит **TSV-текст** (таб-разделитель), поэтому читается как текстовый файл.
- Остальные `.xls`:
  - если файл является `.xlsx` (zip-сигнатура) — читается через EPPlus;
  - если файл является бинарным `.xls` (OLE) — выполняется конвертация во временный `.xlsx` через установленный Microsoft Excel (COM), затем чтение через EPPlus.

## Запуск

Из каталога репозитория:

```powershell
dotnet restore .\ConvertData\ConvertData.csproj

dotnet run --project .\ConvertData\ConvertData.csproj
```

### Запуск с аргументами

Можно передать пути к конкретным файлам (в этом случае папка `EXCEL` не сканируется):

```powershell
dotnet run --project .\ConvertData\ConvertData.csproj -- "C:\path\to\file1.xls" "C:\path\to\file2.xls"
```

## Куда складывать файлы

- Входные файлы: `ConvertData/EXCEL/`
- Результат: `ConvertData/JSON_OUT/`

`JSON_OUT` создаётся автоматически.

> Папка `JSON_OUT` добавлена в `.gitignore`, чтобы не коммитить сгенерированные файлы.

## Требования

- .NET SDK 10
- Пакет `EPPlus` **7.x** (в проекте зафиксирована версия `7.5.1`).
- Для чтения бинарных `.xls` на Windows может потребоваться установленный Microsoft Excel (используется COM automation для конвертации в `.xlsx`).

## Структура проекта (упрощённая "чистая архитектура")

Код разнесён по слоям, логика при этом сохранена:

- `ConvertData/Program.cs` — точка входа и composition root.
- `ConvertData/Application/` — use-case и контракты:
  - `ConvertApp` — оркестрация конвертации.
  - `IRowReader`, `IRowReaderFactory` — чтение входных файлов.
  - `IRowWriter` — запись результата.
  - `ILicenseConfigurator` — настройка лицензии EPPlus.
  - `IPathResolver` — поиск каталога проекта.
- `ConvertData/Domain/Row.cs` — доменная модель строки.
- `ConvertData/Infrastructure/` — реализации:
  - `TsvRowReader` — парсинг TSV (`H2_1.xls`).
  - `EpplusRowReader` — чтение Excel через EPPlus + конвертация бинарного `.xls` через Excel.
  - `JsonRowWriter` — генерация JSON.
  - `EpplusLicenseConfigurator` — `ExcelPackage.LicenseContext = NonCommercial`.
  - `PathResolver` — поиск проекта (`*.csproj`) и фильтрация входных файлов.

## Что было сделано Copilot-ом в этом рабочем каталоге

- Исправлены проблемы лицензии EPPlus, путём понижения с EPPlus 8.x до EPPlus 7.5.1.
- Реализовано пакетное чтение: обработка всех файлов из `EXCEL` и запись соответствующих JSON в `JSON_OUT`.
- Выполнен рефакторинг в классы по слоям (Application/Domain/Infrastructure) без изменения логики.
- Добавлены XML-doc комментарии на русском.
- Добавлены файлы `README.md` и `.gitignore`.

## Примечание про `Console.ReadKey()`

В `Program.cs` стоит `Console.ReadKey()` — приложение будет ждать нажатия клавиши перед выходом (удобно при запуске двойным кликом).
Если запускаешь в CI/скриптах, можно удалить эту строку.
