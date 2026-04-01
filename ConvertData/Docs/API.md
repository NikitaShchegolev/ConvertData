# API Документация ConvertData

## 📚 Оглавление

- [Интерфейсы](#интерфейсы)
- [Application Layer](#application-layer)
- [Domain Models](#domain-models)
- [Infrastructure Layer](#infrastructure-layer)
- [Parsing Utilities](#parsing-utilities)
- [Примеры использования](#примеры-использования)

---

## Интерфейсы

### IRowReader

Интерфейс для чтения данных из различных источников и преобразования в `List<Row>`.

```csharp
namespace ConvertData.Application;

/// <summary>
/// Читает данные из файла и возвращает список объектов Row.
/// </summary>
public interface IRowReader
{
    /// <summary>
    /// Читает данные из указанного файла.
    /// </summary>
    /// <param name="path">Путь к файлу для чтения.</param>
    /// <returns>Список объектов Row.</returns>
    /// <exception cref="InvalidDataException">Если формат данных некорректен.</exception>
    /// <exception cref="IOException">Если файл недоступен.</exception>
    List<Row> Read(string path);
}
```

**Реализации:**
- `EpplusRowReader` — чтение Excel файлов через EPPlus

---

### IRowWriter

Интерфейс для записи `List<Row>` в файл.

```csharp
namespace ConvertData.Application;

/// <summary>
/// Записывает список объектов Row в файл.
/// </summary>
public interface IRowWriter
{
    /// <summary>
    /// Записывает данные в указанный файл.
    /// </summary>
    /// <param name="rows">Список объектов Row для записи.</param>
    /// <param name="outputPath">Путь к выходному файлу.</param>
    void Write(List<Row> rows, string outputPath);
}
```

**Реализации:**
- `JsonRowWriter` — запись в JSON формат

---

### IRowReaderFactory

Фабрика для создания подходящего `IRowReader` на основе типа файла.

```csharp
namespace ConvertData.Application;

/// <summary>
/// Фабрика для создания читателей файлов.
/// </summary>
public interface IRowReaderFactory
{
    /// <summary>
    /// Создаёт подходящий IRowReader для указанного файла.
    /// </summary>
    /// <param name="path">Путь к файлу.</param>
    /// <returns>Экземпляр IRowReader.</returns>
    /// <exception cref="NotSupportedException">Если формат файла не поддерживается.</exception>
    IRowReader Create(string path);
}
```

**Реализации:**
- `RowReaderFactory` — поддерживает .xls и .xlsx

---

### IPathResolver

Интерфейс для разрешения путей к файлам проекта.

```csharp
namespace ConvertData.Application;

/// <summary>
/// Разрешает пути к директориям проекта.
/// </summary>
public interface IPathResolver
{
    /// <summary>
    /// Находит корневую директорию проекта, начиная с указанной.
    /// </summary>
    /// <param name="startDir">Стартовая директория (обычно bin/).</param>
    /// <returns>Путь к директории проекта или null.</returns>
    string? GetProjectDir(string startDir);
}
```

**Реализации:**
- `PathResolver` — ищет директорию с .csproj файлом

---

### ILicenseConfigurator

Интерфейс для конфигурации лицензий внешних библиотек.

```csharp
namespace ConvertData.Application;

/// <summary>
/// Настраивает лицензии для внешних библиотек.
/// </summary>
public interface ILicenseConfigurator
{
    /// <summary>
    /// Применяет необходимую конфигурацию лицензий.
    /// </summary>
    void Configure();
}
```

**Реализации:**
- `EpplusLicenseConfigurator` — устанавливает `LicenseContext = NonCommercial`

---

## Application Layer

### ConvertApp

Главный оркестратор приложения, координирующий все этапы конвертации.

```csharp
namespace ConvertData.Application;

/// <summary>
/// Главное приложение для конвертации данных из Excel в JSON.
/// </summary>
public sealed class ConvertApp
{
    /// <summary>
    /// Запускает процесс конвертации.
    /// </summary>
    /// <param name="args">Аргументы командной строки.</param>
    public void Run(string[] args);
}
```

#### Методы

##### Run(args)

```csharp
public void Run(string[] args)
```

**Параметры:**
- `args` — аргументы командной строки
  - `args[0] = "1"` → только CreateJson
  - `args[0] = "2"` → только ApplyProfiles
  - `args[0] = путь` → конкретный файл
  - `--profile-column=ColumnName` → переопределить колонку профиля

**Последовательность выполнения:**

1. **Определение режима** (`RunModeParser.GetMode`)
2. **Конфигурация лицензий** (`ILicenseConfigurator.Configure`)
3. **Этап 1-2:** Создание JSON и применение профилей (если режим All или CreateJson)
4. **Этап 3-9:** Объединение, обогащение, дедупликация, экспорт (если режим All)

**Пример:**
```csharp
var app = new ConvertApp();
app.Run(new[] { "1", "path/to/file.xlsx" });
```

---

### JsonRecordEnricher

Обогащает неполные записи данными из наиболее полных записей с тем же `CONNECTION_CODE`.

```csharp
namespace ConvertData.Application;

/// <summary>
/// Обогащает неполные записи JSON, копируя данные из наиболее полных записей.
/// </summary>
public static class JsonRecordEnricher
{
    /// <summary>
    /// Обогащает массив JSON объектов.
    /// </summary>
    /// <param name="arr">Массив JSON объектов для обогащения.</param>
    /// <returns>Количество обогащённых записей.</returns>
    public static int Enrich(JsonArray arr);
}
```

#### Пример использования

```csharp
using System.Text.Json.Nodes;

var json = JsonNode.Parse(File.ReadAllText("all.json"))!.AsArray();
int enriched = JsonRecordEnricher.Enrich(json);
Console.WriteLine($"Обогащено записей: {enriched}");
File.WriteAllText("all_enriched.json", json.ToJsonString(...));
```

**Алгоритм:**

1. Группирует объекты по `CONNECTION_CODE`
2. Для каждой группы находит наиболее полную запись (template) по score
3. Копирует `Geometry`, `Bolts`, `Welds`, `TableBrand` из template в неполные записи
4. `InternalForces` и `Coefficients` остаются уникальными

**Оценка полноты (score):**
- +1 за каждое ненулевое значение в `Geometry.Column`, `Geometry.Plate`, `Geometry.Flange`, `Geometry.Stiff`
- +10 если `Bolts.DiameterBolt.F != 0`
- +1 за каждую координату болта
- +1 за каждый катет сварного шва

---

### ConnectionCodeDeduplicator

Устраняет дубликаты `CONNECTION_CODE` путём переименования.

```csharp
namespace ConvertData.Application;

/// <summary>
/// Устраняет дубликаты CONNECTION_CODE путём переименования.
/// </summary>
public static class ConnectionCodeDeduplicator
{
    /// <summary>
    /// Создаёт новый JSON файл без дубликатов CONNECTION_CODE.
    /// </summary>
    /// <param name="inputPath">Путь к входному JSON файлу.</param>
    /// <param name="outputPath">Путь к выходному JSON файлу (без дубликатов).</param>
    /// <param name="reportPath">Путь к файлу отчёта о заменах.</param>
    /// <returns>Количество заменённых кодов.</returns>
    public static int CreateDeduplicatedJson(
        string inputPath, 
        string outputPath, 
        string reportPath);
}
```

#### Пример использования

```csharp
int replaced = ConnectionCodeDeduplicator.CreateDeduplicatedJson(
    inputPath: "all.json",
    outputPath: "all_NotDuplicate.json",
    reportPath: "replacements.txt"
);
Console.WriteLine($"Заменено кодов: {replaced}");
```

**Алгоритм переименования:**

```
Входные данные:
M-3, M-3, M-3, M-5, M-5

Максимальный номер в группе "M": 5

Обработка:
1. M-3 (первое вхождение) → оставить M-3
2. M-3 (дубликат) → переименовать в M-6
3. M-3 (дубликат) → переименовать в M-7
4. M-5 (первое вхождение) → оставить M-5
5. M-5 (дубликат) → переименовать в M-8

Отчёт:
M-3 => M-6
M-3 => M-7
M-5 => M-8
```

---

### ProfileLookupLoader

Загружает справочник профилей балок из Excel.

```csharp
namespace ConvertData.Application;

/// <summary>
/// Загружает справочник геометрических характеристик профилей.
/// </summary>
public sealed class ProfileLookupLoader
{
    /// <summary>
    /// Загружает справочник профилей из Excel файла.
    /// </summary>
    /// <param name="excelProfileDir">Путь к директории с ProfileBeam.xls.</param>
    /// <returns>Словарь: название профиля → геометрия.</returns>
    public Dictionary<string, ProfileGeometry> Load(string excelProfileDir);
}
```

#### Пример использования

```csharp
var loader = new ProfileLookupLoader();
var lookup = loader.Load("EXCEL_Profile/");

if (lookup.TryGetValue("20Б1", out var geom))
{
    Console.WriteLine($"H = {geom.H}, B = {geom.B}");
}
```

**Структура ProfileGeometry:**
```csharp
public sealed record ProfileGeometry
{
    public double H { get; init; }      // Высота
    public double B { get; init; }      // Ширина полки
    public double t_w { get; init; }    // Толщина стенки
    public double t_f { get; init; }    // Толщина полки
    public double A { get; init; }      // Площадь сечения
    public double P { get; init; }      // Масса 1 м
    public double Iz { get; init; }     // Момент инерции относительно оси z
    public double Iy { get; init; }     // Момент инерции относительно оси y
    public double Wz { get; init; }     // Момент сопротивления z
    public double Wy { get; init; }     // Момент сопротивления y
    public double iz { get; init; }     // Радиус инерции z
    public double iy { get; init; }     // Радиус инерции y
    // ... и другие
}
```

---

### JsonProfilePatcher

Применяет справочник профилей к JSON файлам.

```csharp
namespace ConvertData.Application;

/// <summary>
/// Применяет геометрические характеристики из справочника профилей к JSON файлам.
/// </summary>
public sealed class JsonProfilePatcher
{
    /// <summary>
    /// Применяет профили ко всем JSON файлам в директории.
    /// </summary>
    /// <param name="jsonDir">Директория с JSON файлами.</param>
    /// <param name="lookup">Справочник профилей.</param>
    public void ApplyProfilesToJson(
        string jsonDir, 
        Dictionary<string, ProfileGeometry> lookup);
}
```

#### Пример использования

```csharp
var loader = new ProfileLookupLoader();
var lookup = loader.Load("EXCEL_Profile/");

var patcher = new JsonProfilePatcher();
patcher.ApplyProfilesToJson("JSON_OUT/", lookup);
```

**Применяемые свойства:**
- `Beam_H` ← `ProfileGeometry.H`
- `Beam_B` ← `ProfileGeometry.B`
- `Beam_s` ← `ProfileGeometry.t_w`
- `Beam_t` ← `ProfileGeometry.t_f`
- `Beam_A`, `Beam_P`, `Beam_Iz`, `Beam_Iy`, `Beam_Wz`, `Beam_Wy`, `Beam_Sz`, `Beam_Sy`, `Beam_iz`, `Beam_iy` ← соответствующие свойства

---

### ProfileExcelToJsonExporter

Экспортирует справочник профилей из Excel в JSON с категориями.

```csharp
namespace ConvertData.Application;

/// <summary>
/// Экспортирует справочник профилей из Excel в JSON формат.
/// </summary>
public sealed class ProfileExcelToJsonExporter
{
    /// <summary>
    /// Экспортирует профили из Excel файлов в JSON.
    /// </summary>
    /// <param name="excelDir">Директория с Excel файлами профилей.</param>
    /// <param name="outputJsonPath">Путь к выходному JSON файлу.</param>
    public void Export(string excelDir, string outputJsonPath);
}
```

#### Формат выходного JSON

```json
{
  "Двутавр": [
    { "Profile": "20Б1", "H": 200, "B": 100, ... },
    { "Profile": "30Б1", "H": 300, "B": 135, ... }
  ],
  "Швеллер": [
    { "Profile": "20У", "H": 200, "B": 76, ... }
  ],
  "Уголок": [
    { "Profile": "L100x100x8", "H": 100, "B": 100, ... }
  ]
}
```

---

### RunModeParser

Парсит аргументы командной строки для определения режима работы.

```csharp
namespace ConvertData.Application;

/// <summary>
/// Парсер аргументов командной строки.
/// </summary>
public static class RunModeParser
{
    /// <summary>
    /// Определяет режим выполнения приложения.
    /// </summary>
    /// <param name="args">Аргументы командной строки.</param>
    /// <returns>Режим выполнения.</returns>
    public static RunMode GetMode(string[] args);
    
    /// <summary>
    /// Извлекает аргументы файлов для режима CreateJson.
    /// </summary>
    /// <param name="args">Аргументы командной строки.</param>
    /// <returns>Массив путей к файлам.</returns>
    public static string[] GetInputArgsForCreateJson(string[] args);
    
    /// <summary>
    /// Извлекает значение параметра --profile-column.
    /// </summary>
    /// <param name="args">Аргументы командной строки.</param>
    /// <returns>Имя колонки или null.</returns>
    public static string? GetProfileColumn(string[] args);
}
```

#### Примеры

```csharp
// Режим All (по умолчанию)
var mode = RunModeParser.GetMode(new string[0]);
// mode = RunMode.All

// Режим CreateJson
var mode = RunModeParser.GetMode(new[] { "1" });
// mode = RunMode.CreateJson

// Режим ApplyProfiles
var mode = RunModeParser.GetMode(new[] { "2" });
// mode = RunMode.ApplyProfiles

// Переопределение колонки профиля
var profileCol = RunModeParser.GetProfileColumn(
    new[] { "--profile-column=CustomProfile" });
// profileCol = "CustomProfile"
```

---

## Domain Models

### Row

Центральная доменная модель, представляющая одно узловое соединение.

```csharp
namespace ConvertData.Domain;

/// <summary>
/// Модель данных для одного узлового соединения балки.
/// </summary>
public sealed class Row
{
    // === Идентификация ===
    
    /// <summary>Имя группы соединений.</summary>
    public string Name { get; set; } = "";
    
    /// <summary>Уникальный код соединения.</summary>
    public string CONNECTION_CODE { get; set; } = "";
    
    /// <summary>Вариант расчёта.</summary>
    public int variable { get; set; }
    
    /// <summary>Марка опорного столика.</summary>
    public string TableBrand { get; set; } = "";
    
    // === Геометрия балки (18 свойств) ===
    
    /// <summary>Название профиля балки (например, "20Б1").</summary>
    public string ProfileBeam { get; set; } = "";
    
    /// <summary>Высота балки, мм.</summary>
    public double Beam_H { get; set; }
    
    /// <summary>Ширина полки балки, мм.</summary>
    public double Beam_B { get; set; }
    
    /// <summary>Толщина стенки балки, мм.</summary>
    public double Beam_s { get; set; }
    
    /// <summary>Толщина полки балки, мм.</summary>
    public double Beam_t { get; set; }
    
    /// <summary>Площадь сечения балки, см².</summary>
    public double Beam_A { get; set; }
    
    /// <summary>Масса 1 м балки, кг/м.</summary>
    public double Beam_P { get; set; }
    
    /// <summary>Момент инерции относительно оси z, см⁴.</summary>
    public double Beam_Iz { get; set; }
    
    /// <summary>Момент инерции относительно оси y, см⁴.</summary>
    public double Beam_Iy { get; set; }
    
    /// <summary>Момент инерции при кручении, см⁴.</summary>
    public double Beam_Ix { get; set; }
    
    /// <summary>Момент сопротивления относительно оси z, см³.</summary>
    public double Beam_Wz { get; set; }
    
    /// <summary>Момент сопротивления относительно оси y, см³.</summary>
    public double Beam_Wy { get; set; }
    
    /// <summary>Момент сопротивления при кручении, см³.</summary>
    public double Beam_Wx { get; set; }
    
    /// <summary>Статический момент полусечения относительно оси z, см³.</summary>
    public double Beam_Sz { get; set; }
    
    /// <summary>Статический момент полусечения относительно оси y, см³.</summary>
    public double Beam_Sy { get; set; }
    
    /// <summary>Радиус инерции относительно оси z, см.</summary>
    public double Beam_iz { get; set; }
    
    /// <summary>Радиус инерции относительно оси y, см.</summary>
    public double Beam_iy { get; set; }
    
    /// <summary>Координата центра изгиба по x, см.</summary>
    public double Beam_xo { get; set; }
    
    /// <summary>Координата центра изгиба по y, см.</summary>
    public double Beam_yo { get; set; }
    
    // === Геометрия колонны (18 свойств) ===
    // Аналогичная структура
    
    public string ProfileColumn { get; set; } = "";
    public double Column_H { get; set; }
    // ... (14 свойств аналогично Beam)
    
    // === Пластина (3 свойства) ===
    
    public double Plate_H { get; set; }
    public double Plate_B { get; set; }
    public double Plate_t { get; set; }
    
    // === Фланец (4 свойства) ===
    
    public double Flange_Lb { get; set; }
    public double Flange_H { get; set; }
    public double Flange_B { get; set; }
    public double Flange_t { get; set; }
    
    // === Рёбра жёсткости (8 свойств) ===
    
    public double Stiff_tbp { get; set; }
    public double Stiff_tg { get; set; }
    public double Stiff_tf { get; set; }
    public double Stiff_Lh { get; set; }
    public double Stiff_Hh { get; set; }
    public double Stiff_tr1 { get; set; }
    public double Stiff_tr2 { get; set; }
    public double Stiff_twp { get; set; }
    
    // === Болты ===
    
    /// <summary>Список координат болтов.</summary>
    public List<CoordinatesBolts> CoordinatesBolts { get; set; } = new();
    
    /// <summary>Диаметр болта, мм.</summary>
    public int F { get; set; }
    
    /// <summary>Количество болтов.</summary>
    public int Bolts_Nb { get; set; }
    
    /// <summary>Количество рядов болтов.</summary>
    public int N_Rows { get; set; }
    
    /// <summary>Версия болтового соединения.</summary>
    public double OptionBolts { get; set; }
    
    /// <summary>Координата Y первого ряда болтов, мм.</summary>
    public int e1 { get; set; }
    
    /// <summary>Координата X первого ряда болтов (из CoordinatesBolts[0].X), мм.</summary>
    public int d1 { get; set; }
    
    /// <summary>Координата X второго ряда болтов (из CoordinatesBolts[1].X), мм.</summary>
    public int d2 { get; set; }
    
    /// <summary>Расстояния между рядами болтов по Y, мм.</summary>
    public double p1 { get; set; }
    public double p2 { get; set; }
    public double p3 { get; set; }
    public double p4 { get; set; }
    public double p5 { get; set; }
    public double p6 { get; set; }
    public double p7 { get; set; }
    public double p8 { get; set; }
    public double p9 { get; set; }
    public double p10 { get; set; }
    
    // === Сварные швы (10 катетов) ===
    
    public int kf1 { get; set; }
    public int kf2 { get; set; }
    public int kf3 { get; set; }
    public int kf4 { get; set; }
    public int kf5 { get; set; }
    public int kf6 { get; set; }
    public int kf7 { get; set; }
    public int kf8 { get; set; }
    public int kf9 { get; set; }
    public int kf10 { get; set; }
    
    // === Жёсткости ===
    
    public int Sj { get; set; }
    public int Sjo { get; set; }
    
    // === Внутренние усилия (12 компонент) ===
    
    /// <summary>Усилие растяжения, кН.</summary>
    public int Nt { get; set; }
    
    /// <summary>Усилие сжатия, кН.</summary>
    public int Nc { get; set; }
    
    /// <summary>Комбинированное усилие, кН.</summary>
    public int N { get; set; }
    
    /// <summary>Поперечная сила по оси Y, кН.</summary>
    public int Qy { get; set; }
    
    /// <summary>Поперечная сила по оси Z, кН.</summary>
    public int Qz { get; set; }
    
    /// <summary>Поперечная сила по оси X, кН.</summary>
    public int Qx { get; set; }
    
    /// <summary>Изгибающий момент относительно оси Y, кН·м.</summary>
    public int My { get; set; }
    
    /// <summary>Крутящий момент, кН·м.</summary>
    public int T { get; set; }
    
    /// <summary>Обратный изгибающий момент, кН·м.</summary>
    public double Mneg { get; set; }
    
    /// <summary>Изгибающий момент относительно оси Z, кН·м.</summary>
    public double Mz { get; set; }
    
    /// <summary>Изгибающий момент относительно оси X, кН·м.</summary>
    public double Mx { get; set; }
    
    /// <summary>Крутящий момент Mw, кН·м.</summary>
    public double Mw { get; set; }
    
    // === Коэффициенты (6 штук) ===
    
    public double Alpha { get; set; }
    public double Beta { get; set; }
    public double Gamma { get; set; }
    public double Delta { get; set; }
    public double Epsilon { get; set; }
    public double Lambda { get; set; }
}
```

---

### CoordinatesBolts

Координаты одного болта в трёхмерном пространстве.

```csharp
namespace ConvertData.Domain;

/// <summary>
/// Координаты болта в узловом соединении.
/// В блоке Bolts подблоки Y (e1, p1–p10) и X (d1) — это координаты точек 
/// расположения болтов на пластине, не "межболтовые расстояния".
/// </summary>
public class CoordinatesBolts
{
    /// <summary>Координата X (поперечное расстояние), мм.</summary>
    public int X { get; set; }
    
    /// <summary>Координата Y (продольное расстояние), мм.</summary>
    public int Y { get; set; }
    
    /// <summary>Координата Z (высотное расстояние), мм.</summary>
    public int Z { get; set; }
    
    /// <summary>
    /// Инициализирует новый экземпляр CoordinatesBolts.
    /// </summary>
    public CoordinatesBolts(int x, int y, int z)
    {
        X = x;
        Y = y;
        Z = z;
    }
}
```

---

## Infrastructure Layer

### EpplusRowReader

Реализация `IRowReader` для чтения Excel файлов через EPPlus.

```csharp
namespace ConvertData.Infrastructure;

/// <summary>
/// Читает строки из Excel-файлов (.xls/.xlsx) используя библиотеку EPPlus.
/// Поддерживает автоматическую конвертацию .xls в .xlsx через COM Interop.
/// </summary>
public sealed class EpplusRowReader : IRowReader
{
    /// <summary>
    /// Читает данные из Excel-файла и возвращает список объектов Row.
    /// </summary>
    /// <param name="path">Путь к Excel-файлу (.xls или .xlsx).</param>
    /// <returns>Список прочитанных строк.</returns>
    /// <exception cref="InvalidDataException">Если формат данных некорректен.</exception>
    public List<Row> Read(string path);
}
```

#### Поддерживаемые листы

1. **Main** (обязательный) — основные данные
2. **geometry** (опционально) — геометрия пластин, фланцев, рёбер
3. **bolts** (опционально) — параметры болтов
4. **weld** (опционально) — параметры сварных швов

#### Обработка .xls файлов

```csharp
// Автоматическая конвертация через COM Interop
if (format == ExcelFileFormat.CompoundFileBinary)
{
    var tmpXlsx = Path.Combine(Path.GetTempPath(), 
        Path.GetFileNameWithoutExtension(path) + "_converted_" + Guid.NewGuid() + ".xlsx");
    
    ExcelXlsConverter.ConvertXlsToXlsxViaExcel(path, tmpXlsx);
    return ReadXlsxWithEpplus(tmpXlsx);
}
```

---

### JsonRowWriter

Реализация `IRowWriter` для записи в JSON формат.

```csharp
namespace ConvertData.Infrastructure;

/// <summary>
/// Записывает список объектов Row в JSON-файл с форматированием.
/// </summary>
public sealed class JsonRowWriter : IRowWriter
{
    /// <summary>
    /// Записывает список объектов Row в JSON-файл.
    /// </summary>
    /// <param name="rows">Список объектов Row.</param>
    /// <param name="outputPath">Путь к выходному JSON-файлу.</param>
    public void Write(List<Row> rows, string outputPath);
}
```

#### Особенности

- **Формат**: Отступы по 2/4 пробела, читаемый JSON
- **Кодировка**: UTF-8 без BOM
- **Числа**: InvariantCulture (точка в десятичных)
- **Экранирование**: Полное экранирование спецсимволов JSON

---

### RowMapper

Отображает данные из Excel в объекты `Row`.

```csharp
namespace ConvertData.Infrastructure;

/// <summary>
/// Преобразует данные из Excel в объекты Row.
/// </summary>
public static class RowMapper
{
    /// <summary>
    /// Создаёт Row из основной таблицы Excel.
    /// </summary>
    public static Row MapMainRow(
        string name, 
        string code, 
        string profile, 
        string? profileColumn,
        string? h, string? b, string? s, string? t,
        string? nt, string? qy, string? qz, string? tcell,
        string? nc, string? n, string? my, string? variableStr,
        string? sj, string? sjo, string? mneg,
        string? mz, string? mx, string? mw,
        string? alpha, string? beta, string? gamma,
        string? delta, string? epsilon, string? lambda);
    
    /// <summary>
    /// Создаёт Row из таблицы профилей.
    /// </summary>
    public static Row MapProfileRow(
        string profile, 
        string? h, string? b, string? s, string? t);
}
```

---

## Parsing Utilities

### NumericParser

Утилиты для парсинга чисел с поддержкой русского и инвариантного форматов.

```csharp
namespace ConvertData.Infrastructure.Parsing;

/// <summary>
/// Парсер числовых значений, поддерживающий как русский (с запятой), 
/// так и инвариантный (с точкой) форматы.
/// </summary>
public static class NumericParser
{
    /// <summary>
    /// Парсит строку в значение типа double.
    /// </summary>
    /// <param name="s">Строка для парсинга.</param>
    /// <returns>Числовое значение или 0.0 при ошибке.</returns>
    public static double ParseDouble(string? s);
    
    /// <summary>
    /// Парсит строку в значение типа int.
    /// Если прямой парсинг не удался, пытается через double с округлением.
    /// </summary>
    /// <param name="s">Строка для парсинга.</param>
    /// <returns>Целое число или 0 при ошибке.</returns>
    public static int ParseInt(string? s);
}
```

#### Примеры

```csharp
NumericParser.ParseDouble("3,14");    // 3.14 (русский формат)
NumericParser.ParseDouble("3.14");    // 3.14 (инвариантный формат)
NumericParser.ParseDouble("1 234,56"); // 1234.56 (с разделителями тысяч)

NumericParser.ParseInt("42");         // 42
NumericParser.ParseInt("42.7");       // 43 (округление)
NumericParser.ParseInt("42,7");       // 43 (округление, русский формат)
NumericParser.ParseInt("abc");        // 0 (ошибка парсинга)
```

---

### HeaderUtils

Утилиты для работы с заголовками Excel.

```csharp
namespace ConvertData.Infrastructure.Parsing;

/// <summary>
/// Утилиты для нормализации и поиска заголовков в Excel.
/// </summary>
public static class HeaderUtils
{
    /// <summary>
    /// Нормализует заголовок, удаляя невидимые символы.
    /// </summary>
    /// <param name="h">Исходный заголовок.</param>
    /// <returns>Нормализованный заголовок.</returns>
    public static string NormalizeHeader(string h);
    
    /// <summary>
    /// Находит индекс заголовка в списке (без учёта регистра).
    /// </summary>
    /// <param name="header">Список заголовков.</param>
    /// <param name="name">Искомое имя.</param>
    /// <returns>Индекс или -1, если не найдено.</returns>
    public static int IndexOfHeader(List<string> header, string name);
    
    /// <summary>
    /// Находит индекс любого из указанных заголовков.
    /// </summary>
    /// <param name="header">Список заголовков.</param>
    /// <param name="names">Варианты имён для поиска.</param>
    /// <returns>Индекс или -1, если ни один не найден.</returns>
    public static int IndexOfHeaderAny(
        List<string> header, 
        IEnumerable<string> names);
}
```

#### Примеры

```csharp
var headers = new List<string> { "Name", "CODE", "Profile" };

HeaderUtils.IndexOfHeader(headers, "code");        // 1 (игнорирует регистр)
HeaderUtils.IndexOfHeader(headers, "Missing");     // -1

HeaderUtils.IndexOfHeaderAny(headers, new[] { "Код", "CODE", "Connection_Code" });  
// 1 (нашёл "CODE")
```

---

### ExcelFileSignature

Определяет формат Excel файла по сигнатуре.

```csharp
namespace ConvertData.Infrastructure.Parsing;

/// <summary>
/// Определяет формат Excel файла по первым байтам.
/// </summary>
public static class ExcelFileSignature
{
    /// <summary>
    /// Определяет формат Excel файла.
    /// </summary>
    /// <param name="path">Путь к файлу.</param>
    /// <returns>Формат файла.</returns>
    public static ExcelFileFormat Detect(string path);
}

public enum ExcelFileFormat
{
    Unknown,
    ZipXlsx,              // .xlsx (ZIP-архив)
    CompoundFileBinary    // .xls (двоичный формат)
}
```

#### Примеры

```csharp
var format = ExcelFileSignature.Detect("file.xlsx");
// format = ExcelFileFormat.ZipXlsx

var format = ExcelFileSignature.Detect("file.xls");
// format = ExcelFileFormat.CompoundFileBinary
```

---

## Примеры использования

### Полный цикл конвертации

```csharp
using ConvertData.Application;

var app = new ConvertApp();
app.Run(Array.Empty<string>());  // Режим All
```

### Конвертация конкретного файла

```csharp
using ConvertData.Application;
using ConvertData.Infrastructure;
using ConvertData.Domain;

var reader = new EpplusRowReader();
List<Row> rows = reader.Read("path/to/file.xlsx");

var writer = new JsonRowWriter();
writer.Write(rows, "output.json");
```

### Применение справочника профилей

```csharp
var loader = new ProfileLookupLoader();
var lookup = loader.Load("EXCEL_Profile/");

var patcher = new JsonProfilePatcher();
patcher.ApplyProfilesToJson("JSON_OUT/", lookup);
```

### Обогащение неполных записей

```csharp
using System.Text.Json.Nodes;

var json = JsonNode.Parse(File.ReadAllText("all.json"))!.AsArray();
int enriched = JsonRecordEnricher.Enrich(json);

var options = new JsonSerializerOptions { WriteIndented = true };
File.WriteAllText("all_enriched.json", 
    json.ToJsonString(options), 
    new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
```

### Устранение дубликатов

```csharp
int replaced = ConnectionCodeDeduplicator.CreateDeduplicatedJson(
    inputPath: "all.json",
    outputPath: "all_NotDuplicate.json",
    reportPath: "replacements.txt"
);

Console.WriteLine($"Заменено кодов: {replaced}");
```

---

## Обработка ошибок

### Исключения

- **`InvalidDataException`** — некорректный формат данных в Excel
- **`IOException`** — проблемы доступа к файлам
- **`NotSupportedException`** — неподдерживаемый формат файла
- **`COMException`** — ошибка при работе с Excel COM Interop

### Стратегия обработки

```csharp
try
{
    var rows = reader.Read(path);
}
catch (InvalidDataException ex)
{
    Console.WriteLine($"Ошибка формата данных: {ex.Message}");
}
catch (IOException ex)
{
    Console.WriteLine($"Ошибка доступа к файлу: {ex.Message}");
}
```

---

**См. также:**
- [README](README.md)
- [Архитектура](Architecture.md)
- [Поток данных](DataFlow.md)
- [UML диаграммы](UML.md)
