# Архитектура ConvertData

## 🏗️ Общий обзор

ConvertData построен на принципах **Clean Architecture** с разделением на слои:

```
┌─────────────────────────────────────────┐
│         Presentation Layer              │
│         (Program.cs)                    │
└─────────────────────────────────────────┘
                  ↓
┌─────────────────────────────────────────┐
│         Application Layer               │
│  (Use Cases & Business Logic)           │
│  - ConvertApp                           │
│  - JsonRecordEnricher                   │
│  - ProfileExcelToJsonExporter           │
│  - ConnectionCodeDeduplicator           │
└─────────────────────────────────────────┘
                  ↓
┌─────────────────────────────────────────┐
│         Domain Layer                    │
│  (Core Business Models)                 │
│  - Row                                  │
│  - CoordinatesBolts                     │
│  - ProfileGeometry                      │
└─────────────────────────────────────────┘
                  ↑
┌─────────────────────────────────────────┐
│      Infrastructure Layer               │
│  (External Dependencies)                │
│  - EpplusRowReader (EPPlus)             │
│  - JsonRowWriter (System.Text.Json)     │
│  - ExcelXlsConverter (COM Interop)      │
└─────────────────────────────────────────┘
```

## 📦 Слои приложения

### 1. Presentation Layer (Точка входа)

**Файлы:**
- `Program.cs` — Main entry point

**Ответственность:**
- Создание и запуск `ConvertApp`
- Ожидание ввода пользователя

### 2. Application Layer (Бизнес-логика)

#### 2.1 Главный оркестратор

**ConvertApp.cs**
```csharp
internal sealed class ConvertApp
{
    public void Run(string[] args)
    {
        // Этап 1: CreateJson
        // Этап 2: ApplyProfiles
        // Этап 3: MergeAll
        // Этап 3.5: Enrich
        // Этап 4-9: Export & Deduplicate
    }
}
```

Оркестрирует весь процесс конвертации, вызывая нужные компоненты в правильном порядке.

#### 2.2 Use Cases (Сценарии использования)

**ProfileExcelToJsonExporter** — Экспорт справочника профилей
```csharp
public void Export(string excelDir, string outputJsonPath)
{
    // Читает ProfileI.xlsx, ProfileC.xlsx, ProfileL.xlsx
    // Объединяет в Profile.json с категориями
}
```

**JsonRecordEnricher** — Обогащение неполных записей
```csharp
public int Enrich(JsonArray arr)
{
    // Группирует по CONNECTION_CODE
    // Находит наиболее полную запись (template)
    // Копирует Geometry, Bolts, Welds, TableBrand
}
```

**ConnectionCodeDeduplicator** — Устранение дубликатов
```csharp
public int CreateDeduplicatedJson(...)
{
    // Находит дубликаты CONNECTION_CODE
    // Переименовывает: M-3 → M-4, M-5...
}
```

**ProfileLookupLoader** — Загрузка справочника
```csharp
public Dictionary<string, ProfileGeometry> Load(string dir)
{
    // Читает ProfileBeam.xls
    // Создаёт словарь Profile → Geometry
}
```

**JsonProfilePatcher** — Применение профилей
```csharp
public void ApplyProfilesToJson(string jsonDir, Dictionary<...> lookup)
{
    // Для каждого JSON файла
    // Обогащает Beam_H, Beam_B, Beam_s, Beam_t и т.д.
}
```

#### 2.3 Интерфейсы

```csharp
interface IRowReader
{
    List<Row> Read(string path);
}

interface IRowWriter
{
    void Write(List<Row> rows, string outputPath);
}

interface IRowReaderFactory
{
    IRowReader Create(string path);
}

interface IPathResolver
{
    string? GetProjectDir(string startDir);
}

interface ILicenseConfigurator
{
    void Configure();
}
```

### 3. Domain Layer (Доменные модели)

#### 3.1 Основная модель данных

**Row.cs** — Центральная сущность, представляющая одно соединение
```csharp
internal sealed class Row
{
    // Идентификация
    public string Name { get; set; }
    public string CONNECTION_CODE { get; set; }
    public int variable { get; set; }
    public string TableBrand { get; set; }
    
    // Геометрия балки
    public string ProfileBeam { get; set; }
    public double Beam_H, Beam_B, Beam_s, Beam_t { get; set; }
    public double Beam_A, Beam_P { get; set; }
    public double Beam_Iz, Beam_Iy, Beam_Ix { get; set; }
    // ... 14 свойств балки
    
    // Геометрия колонны
    public string ProfileColumn { get; set; }
    public double Column_H, Column_B, Column_s, Column_t { get; set; }
    // ... 14 свойств колонны
    
    // Геометрия пластин и фланцев
    public double Plate_H, Plate_B, Plate_t { get; set; }
    public double Flange_Lb, Flange_H, Flange_B, Flange_t { get; set; }
    
    // Рёбра жёсткости
    public double Stiff_tbp, Stiff_tg, Stiff_tf { get; set; }
    public double Stiff_Lh, Stiff_Hh, Stiff_tr1, Stiff_tr2, Stiff_twp { get; set; }
    
    // Болты
    public List<CoordinatesBolts> CoordinatesBolts { get; set; }
    public int F, Bolts_Nb, N_Rows { get; set; }
    public double OptionBolts { get; set; }
    public int e1, d1, d2 { get; set; }
    public double p1, p2, p3, p4, p5, p6, p7, p8, p9, p10 { get; set; }
    
    // Сварные швы
    public int kf1, kf2, kf3, kf4, kf5, kf6, kf7, kf8, kf9, kf10 { get; set; }
    
    // Жёсткости
    public int Sj, Sjo { get; set; }
    
    // Внутренние силы
    public int Nt, Nc, N, Qy, Qz, Qx, My, T { get; set; }
    public double Mneg, Mz, Mx, Mw { get; set; }
    
    // Коэффициенты
    public double Alpha, Beta, Gamma, Delta, Epsilon, Lambda { get; set; }
}
```

**CoordinatesBolts.cs** — Координаты болта (X, Y, Z)
```csharp
internal class CoordinatesBolts
{
    public int X { get; set; }
    public int Y { get; set; }
    public int Z { get; set; }
}
```

**ProfileGeometry.cs** — Геометрия профиля из справочника
```csharp
internal sealed record ProfileGeometry
{
    public double H, B, t_w, t_f { get; init; }
    public double r1, r2, A, P { get; init; }
    public double Iz, Iy, Ix { get; init; }
    public double Wz, Wy, Wx { get; init; }
    public double Sz, Sy { get; init; }
    public double iz, iy, iu { get; init; }
    public double xo, yo { get; init; }
}
```

### 4. Infrastructure Layer (Инфраструктура)

#### 4.1 Чтение данных

**EpplusRowReader.cs** — Читает Excel файлы
```csharp
public List<Row> Read(string path)
{
    // 1. Определяет формат (.xls или .xlsx)
    // 2. Конвертирует .xls → .xlsx если нужно
    // 3. Читает основной лист
    // 4. Объединяет данные из дополнительных листов
    // 5. Возвращает List<Row>
}
```

Этапы чтения:
1. **Определение формата** (`ExcelFileSignature.Detect`)
2. **Конвертация .xls** (если нужно, через `ExcelXlsConverter`)
3. **Поиск заголовков** (`FindHeaderRow`)
4. **Разрешение колонок** (`ExcelHeaderResolver.Resolve`)
5. **Чтение данных** (`RowMapper.MapMainRow/MapProfileRow`)
6. **Объединение листов** (`MergeAdditionalSheets`)

**ExcelHeaderResolver.cs** — Разрешает заголовки Excel
```csharp
public static ExcelColumnMap Resolve(List<string> header)
{
    // Ищет колонки по известным вариантам имён
    // Поддерживает русские и английские названия
    // Применяет fallback стратегии для греческих букв
}
```

**RowMapper.cs** — Маппинг из строк Excel в Row
```csharp
public static Row MapMainRow(
    string name, string code, string profile, ...)
{
    // Создаёт Row
    // Парсит все поля через NumericParser
}
```

#### 4.2 Запись данных

**JsonRowWriter.cs** — Записывает Row в JSON
```csharp
public void Write(List<Row> rows, string outputPath)
{
    // Строит JSON вручную через StringBuilder
    // Форматирует числа в InvariantCulture
    // Экранирует спецсимволы в строках
}
```

Структура вывода:
```
[
  {
    Name,
    CONNECTION_CODE,
    variable,
    TableBrand,
    Stiffness { Sj, Sjo },
    Geometry { Beam, Column, Plate, Flange, Stiff },
    Bolts { Option, DiameterBolt, CountBolt, BoltRow, CoordinatesBolts },
    Welds { kf1-kf10 },
    InternalForces { N, Nt, Nc, My, Mz, Mx, Mw, Mneg, T, Qy, Qz, Qx },
    Coefficients { Alpha, Beta, Gamma, Delta, Epsilon, Lambda }
  }
]
```

**JsonMerger.cs** — Объединяет JSON файлы
```csharp
public JsonArray MergeAll(string jsonDir)
{
    // Читает все *.json
    // Объединяет массивы
    // Возвращает единый JsonArray
}
```

#### 4.3 Парсинг

**NumericParser.cs** — Парсинг чисел
```csharp
public static double ParseDouble(string? s)
{
    // Пытается русский формат (запятая)
    // Пытается инвариантный формат (точка)
    // Возвращает 0.0 при ошибке
}

public static int ParseInt(string? s)
{
    // Пытается парсить как int
    // Если не удалось, парсит как double и округляет
}
```

**HeaderUtils.cs** — Утилиты для заголовков
```csharp
public static string NormalizeHeader(string h)
{
    // Удаляет невидимые символы (U+00A0, U+FEFF, U+200B...)
    // Обрезает пробелы
}

public static int IndexOfHeader(List<string> header, string name)
{
    // Ищет заголовок (OrdinalIgnoreCase)
}

public static int IndexOfHeaderAny(List<string> header, IEnumerable<string> names)
{
    // Ищет любой из вариантов
}
```

**ExcelFileSignature.cs** — Определение формата Excel
```csharp
public static ExcelFileFormat Detect(string path)
{
    // Читает первые байты файла
    // "PK" (0x50 0x4B) → ZipXlsx
    // 0xD0 0xCF 0x11 0xE0 → CompoundFileBinary (xls)
}
```

#### 4.4 Interop

**ExcelXlsConverter.cs** — Конвертация через COM
```csharp
public static void ConvertXlsToXlsxViaExcel(string xlsPath, string xlsxPath)
{
    // Создаёт Excel.Application через COM
    // Открывает .xls
    // Сохраняет как .xlsx (XlFileFormat.xlOpenXMLWorkbook = 51)
    // Закрывает Excel
}
```

## 🔄 Паттерны проектирования

### Factory Pattern
**RowReaderFactory** создаёт нужную реализацию `IRowReader` в зависимости от расширения файла:
```csharp
public IRowReader Create(string path)
{
    if (ext == ".xls" || ext == ".xlsx")
        return new EpplusRowReader();
    
    throw new NotSupportedException();
}
```

### Strategy Pattern
Разные стратегии парсинга чисел:
1. Русская культура (запятая)
2. Инвариантная культура (точка)
3. Fallback через double парсинг для int

### Template Method Pattern
`ConvertApp.Run()` определяет скелет алгоритма, вызывая методы в строгом порядке.

### Dependency Injection (через поля)
```csharp
private readonly IRowWriter _writer = new JsonRowWriter();
private readonly IRowReaderFactory _readerFactory = new RowReaderFactory();
```

### Builder Pattern (косвенно)
`StringBuilder` в `JsonRowWriter` для построения JSON.

## 📏 Принципы SOLID

### Single Responsibility Principle (SRP)
- `NumericParser` — только парсинг чисел
- `HeaderUtils` — только работа с заголовками
- `JsonRowWriter` — только запись JSON
- `EpplusRowReader` — только чтение Excel

### Open/Closed Principle (OCP)
- `IRowReader` открыт для расширения (новые реализации), закрыт для модификации
- Новые форматы можно добавить через новые реализации интерфейса

### Liskov Substitution Principle (LSP)
- Любая реализация `IRowReader` может заменить другую без изменения логики

### Interface Segregation Principle (ISP)
- Маленькие интерфейсы: `IRowReader`, `IRowWriter`, `IPathResolver`
- Клиенты зависят только от нужных методов

### Dependency Inversion Principle (DIP)
- `ConvertApp` зависит от абстракций (`IRowWriter`, `IRowReaderFactory`), а не от конкретных реализаций

## 🗂️ Структура данных

### Многолистовая архитектура Excel

```
ExcelWorkbook.xlsx
├── Sheet: Main (обязательный)
│   ├── Name, CONNECTION_CODE, ProfileBeam
│   ├── variable, Sj, Sjo
│   ├── Усилия: Nt, Nc, N, Qy, Qz, My, Mz, Mx, Mw, Mneg, T
│   └── Коэффициенты: α, β, γ, δ, ε, λ
├── Sheet: geometry (опционально)
│   ├── CONNECTION_CODE (ключ связи)
│   ├── H, B, tp (пластина/фланец)
│   └── tbp, tg, tf, Lh, Hh, tr1, tr2, twp (рёбра)
├── Sheet: bolts (опционально)
│   ├── CONNECTION_CODE
│   ├── Option, F, Nb
│   ├── e1, p1-p10 (координаты Y)
│   ├── d1, d2 (координаты X)
│   └── Марка опорного столика
└── Sheet: weld (опционально)
    ├── CONNECTION_CODE
    └── kf1-kf10 (катеты швов)
```

### Логика объединения листов

```csharp
// 1. Создать словарь main rows по CONNECTION_CODE
var codeLookup = rows.ToDictionary(r => r.CONNECTION_CODE);

// 2. Для каждого дополнительного листа
foreach (var ws in ["geometry", "bolts", "weld"])
{
    // 3. Прочитать заголовки
    var headers = ReadHeaders(ws);
    
    // 4. Для каждой строки данных
    foreach (var dataRow in ws.DataRows)
    {
        // 5. Найти целевую Row
        Row? target = null;
        
        // Стратегия 1: По CONNECTION_CODE
        if (keyCol >= 0)
        {
            var code = GetCell(dataRow, keyCol);
            codeLookup.TryGetValue(code, out target);
        }
        
        // Стратегия 2: По индексу строки
        if (target == null)
        {
            int idx = dataRow.Index - headerRow - 1;
            target = rows[idx];
        }
        
        // 6. Применить значения
        foreach (var (col, setter) in columnMappings)
        {
            var value = GetCell(dataRow, col);
            setter(target, value);
        }
    }
}
```

## 🔐 Безопасность и надёжность

### Обработка ошибок
```csharp
try
{
    // Чтение файла
}
catch (Exception ex)
{
    Console.WriteLine($"Error: {ex.Message}");
    // Продолжаем работу с другими файлами
}
```

### Валидация данных
```csharp
if (string.IsNullOrWhiteSpace(code))
    continue; // Пропускаем строки без кода

if (map.IdxCode < 0)
    throw new InvalidDataException("Cannot find CODE column");
```

### Безопасный парсинг
```csharp
public static double ParseDouble(string? s)
{
    // Никогда не выбрасывает исключения
    // Возвращает 0.0 при ошибке
}
```

### Очистка ресурсов
```csharp
try
{
    ConvertXlsToXlsx(xls, tmpXlsx);
    return ReadXlsx(tmpXlsx);
}
finally
{
    // Удаляем временный файл
    if (File.Exists(tmpXlsx))
        File.Delete(tmpXlsx);
}
```

## 🎯 Точки расширения

### Новые форматы ввода
Реализуйте `IRowReader`:
```csharp
internal class CsvRowReader : IRowReader
{
    public List<Row> Read(string path)
    {
        // Логика чтения CSV
    }
}
```

Зарегистрируйте в `RowReaderFactory`:
```csharp
if (ext == ".csv")
    return new CsvRowReader();
```

### Новые форматы вывода
Реализуйте `IRowWriter`:
```csharp
internal class XmlRowWriter : IRowWriter
{
    public void Write(List<Row> rows, string outputPath)
    {
        // Логика записи XML
    }
}
```

### Новые типы профилей
Добавьте в `ProfileExcelToJsonExporter`:
```csharp
private static readonly Dictionary<string, string> FileCategoryMap = new()
{
    ["ProfileI.xlsx"] = "Двутавр",
    ["ProfileC.xlsx"] = "Швеллер",
    ["ProfileL.xlsx"] = "Уголок",
    ["ProfileT.xlsx"] = "Тавр", // Новый тип
};
```

---

**См. также:**
- [Поток данных](DataFlow.md)
- [UML диаграммы](UML.md)
- [API документация](API.md)
