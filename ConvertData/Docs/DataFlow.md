# Поток данных ConvertData

## 🌊 Общая схема потока данных

```
┌──────────────┐
│ Excel Files  │
│ (.xls/.xlsx) │
└──────┬───────┘
       │
       │ [1] Read & Parse
       ▼
┌──────────────┐      ┌─────────────────┐
│  Row Reader  │─────▶│ List<Row>       │
│  (EPPlus)    │      │ (In Memory)     │
└──────────────┘      └────────┬────────┘
                               │
                               │ [2] Apply Profiles
                               ▼
                      ┌─────────────────┐
                      │ ProfileLookup   │
                      │ Dictionary      │
                      └────────┬────────┘
                               │
                               │ [3] Write Individual JSON
                               ▼
                      ┌─────────────────┐
                      │ JSON_OUT/*.json │
                      └────────┬────────┘
                               │
                               │ [4] Merge All
                               ▼
                      ┌─────────────────┐
                      │ JsonArray       │
                      │ (All Records)   │
                      └────────┬────────┘
                               │
                               │ [5] Enrich
                               ▼
                      ┌─────────────────┐
                      │ Enriched        │
                      │ JsonArray       │
                      └────────┬────────┘
                               │
                               │ [6] Save all.json
                               │ [7] Deduplicate
                               ▼
                      ┌─────────────────┐
                      │all_NotDuplicate │
                      │    .json        │
                      └────────┬────────┘
                               │
                               │ [8] Export Lists
                               ▼
                      ┌─────────────────┐
                      │ CONNECTION_CODE │
                      │ _new.json       │
                      │ Profile.json    │
                      │ NameConnections │
                      │ .json           │
                      └─────────────────┘
```

## 📂 Этап 1: Чтение Excel → List<Row>

### Входные данные
```
EXCEL/
├── M_P.xls          (балка к стенке колонны, опорный столик)
├── M_BK.xlsx        (балка к полке колонны)
└── ...
```

### Процесс

#### 1.1 Определение формата файла

```csharp
ExcelFileSignature.Detect(path)
┌─────────────────────────────────────┐
│ Читает первые 4 байта файла        │
│                                     │
│ 0x50 0x4B (PK) → ZipXlsx (.xlsx)   │
│ 0xD0 0xCF 0x11 0xE0 → Binary (.xls) │
└─────────────────────────────────────┘
```

#### 1.2 Конвертация .xls → .xlsx (если нужно)

```csharp
if (format == ExcelFileFormat.CompoundFileBinary)
{
    ExcelXlsConverter.ConvertXlsToXlsxViaExcel(path, tmpXlsx)
    
    ┌──────────────────────────────────┐
    │ COM Interop                      │
    │ Excel.Application.Workbooks.Open │
    │ workbook.SaveAs(tmpXlsx, 51)     │
    └──────────────────────────────────┘
}
```

Временный файл:
```
C:\Users\...\Temp\M_P_converted_abc123.xlsx
```

#### 1.3 Чтение основного листа

```csharp
ReadXlsxWithEpplus(xlsxPath)
│
├─ 1. Открыть ExcelPackage
├─ 2. Взять первый worksheet
├─ 3. Найти строку заголовков (FindHeaderRow)
│    │
│    └─ Ищет строку, содержащую:
│         - CONNECTION_CODE + Name + ProfileBeam (основная таблица)
│         - ProfileBeam + Beam_H + Beam_B + Beam_s + Beam_t (таблица профилей)
│
├─ 4. Разрешить колонки (ExcelHeaderResolver.Resolve)
│    │
│    └─ Создать ExcelColumnMap:
│         IdxName = 0
│         IdxCode = 1
│         IdxProfile = 2
│         IdxH = 3
│         IdxB = 4
│         ...
│
├─ 5. Прочитать данные строк
│    │
│    └─ Для каждой строки после заголовка:
│         if (IsMainTable)
│             row = RowMapper.MapMainRow(name, code, profile, ...)
│         else if (IsProfileTable)
│             row = RowMapper.MapProfileRow(profile, h, b, s, t)
│
└─ 6. Объединить дополнительные листы (MergeAdditionalSheets)
```

#### 1.4 Объединение данных из листов geometry, bolts, weld

```
Worksheet: geometry
┌─────────────┬──────┬──────┬──────┬──────┬──────┬─────┐
│ CODE        │  H   │  B   │  tp  │  Lb  │ tbp  │ ... │
├─────────────┼──────┼──────┼──────┼──────┼──────┼─────┤
│ M-3         │ 300  │ 150  │  10  │  40  │  8   │ ... │
│ M-4         │ 350  │ 180  │  12  │  50  │  10  │ ... │
└─────────────┴──────┴──────┴──────┴──────┴──────┴─────┘

Действие:
1. Найти Row с CONNECTION_CODE = "M-3" в основном списке
2. Установить:
   row.Plate_H = 300
   row.Plate_B = 150
   row.Plate_t = 10
   row.Flange_H = 300
   row.Flange_B = 150
   row.Flange_t = 10
   row.Flange_Lb = 40
   row.Stiff_tbp = 8
   ...
```

```
Worksheet: bolts
┌─────────────┬────────┬─────┬─────┬─────┬─────┬─────┬──────────────────────────┐
│ CODE        │ Option │  F  │ Nb  │ e1  │ d1  │ d2  │ Марка опорного столика   │
├─────────────┼────────┼─────┼─────┼─────┼─────┼─────┼──────────────────────────┤
│ M-3         │   1    │ 20  │  4  │ 50  │ 60  │ 120 │ Т1                       │
└─────────────┴────────┴─────┴─────┴─────┴─────┴─────┴──────────────────────────┘

Действие:
row.OptionBolts = 1
row.F = 20
row.Bolts_Nb = 4
row.e1 = 50
row.CoordinatesBolts[0].X = 60
row.CoordinatesBolts[1].X = 120
row.N_Rows = 2
row.TableBrand = "Т1"  ← ВАЖНО!
```

```
Worksheet: weld
┌─────────────┬─────┬─────┬─────┬─────┐
│ CODE        │ kf1 │ kf2 │ kf3 │ ... │
├─────────────┼─────┼─────┼─────┼─────┤
│ M-3         │  6  │  6  │  8  │ ... │
└─────────────┴─────┴─────┴─────┴─────┘

Действие:
row.kf1 = 6
row.kf2 = 6
row.kf3 = 8
...
```

### Выходные данные этапа 1

```csharp
List<Row> rows = [
    Row { 
        Name = "Балка-Колонна M",
        CONNECTION_CODE = "M-3",
        ProfileBeam = "20Б1",
        variable = 1,
        TableBrand = "Т1",
        F = 20,
        e1 = 50,
        d1 = 60, d2 = 120,
        Plate_H = 300, Plate_B = 150,
        kf1 = 6, kf2 = 6,
        ...
    },
    ...
]
```

## 📂 Этап 2: Применение справочника профилей

### Входные данные

```
EXCEL_Profile/
└── ProfileBeam.xls
    ├── Sheet: ProfileI (двутавры)
    │   ┌─────────┬──────┬──────┬──────┬──────┬──────┬─────┐
    │   │ Profile │  H   │  B   │  s   │  t   │  A   │ ... │
    │   ├─────────┼──────┼──────┼──────┼──────┼──────┼─────┤
    │   │ 20Б1    │ 200  │ 100  │ 5.6  │ 8.5  │ 30.6 │ ... │
    │   │ 30Б1    │ 300  │ 135  │ 6.5  │ 9.5  │ 42.7 │ ... │
    │   └─────────┴──────┴──────┴──────┴──────┴──────┴─────┘
    ├── Sheet: ProfileC (швеллеры)
    └── Sheet: ProfileL (уголки)
```

### Процесс

#### 2.1 Загрузка справочника

```csharp
ProfileLookupLoader.Load(excelProfileDir)

Dictionary<string, ProfileGeometry> lookup = {
    ["20Б1"] = ProfileGeometry {
        H = 200, B = 100, t_w = 5.6, t_f = 8.5,
        A = 30.6, P = 24.0,
        Iz = 1840, Iy = 155, Ix = 1.68,
        Wz = 184, Wy = 24.2, Wx = 2.35,
        Sz = 104, Sy = 15.6,
        iz = 7.75, iy = 2.25, iu = 0.74,
        xo = 0, yo = 0
    },
    ["30Б1"] = ProfileGeometry { ... },
    ...
}
```

#### 2.2 Применение к Row

```csharp
JsonProfilePatcher.ApplyProfilesToJson(jsonOutDir, lookup)

Для каждого файла в JSON_OUT/:
    Прочитать JsonArray
    Для каждого объекта:
        profile = obj["Geometry"]["Beam"]["ProfileBeam"]
        
        if (lookup.TryGetValue(profile, out var geom))
        {
            obj["Geometry"]["Beam"]["Beam_H"] = geom.H
            obj["Geometry"]["Beam"]["Beam_B"] = geom.B
            obj["Geometry"]["Beam"]["Beam_s"] = geom.t_w
            obj["Geometry"]["Beam"]["Beam_t"] = geom.t_f
            obj["Geometry"]["Beam"]["Beam_A"] = geom.A
            obj["Geometry"]["Beam"]["Beam_P"] = geom.P
            obj["Geometry"]["Beam"]["Beam_Iz"] = geom.Iz
            obj["Geometry"]["Beam"]["Beam_Iy"] = geom.Iy
            obj["Geometry"]["Beam"]["Beam_Iz"] = geom.Iz
            obj["Geometry"]["Beam"]["Beam_Wz"] = geom.Wz
            obj["Geometry"]["Beam"]["Beam_Wy"] = geom.Wy
            obj["Geometry"]["Beam"]["Beam_Sz"] = geom.Sz
            obj["Geometry"]["Beam"]["Beam_Sy"] = geom.Sy
            obj["Geometry"]["Beam"]["Beam_iz"] = geom.iz
            obj["Geometry"]["Beam"]["Beam_iy"] = geom.iy
        }
    
    Сохранить файл обратно
```

### Выходные данные этапа 2

```
JSON_OUT/
├── M_P.json    (обновлён с геометрией из справочника)
├── M_BK.json   (обновлён)
└── ...
```

Каждый Row теперь имеет полные геометрические характеристики:
```json
{
  "Geometry": {
    "Beam": {
      "ProfileBeam": "20Б1",
      "Beam_H": 200,
      "Beam_B": 100,
      "Beam_s": 5.6,
      "Beam_t": 8.5,
      "Beam_A": 30.6,
      "Beam_P": 24.0,
      "Beam_Iz": 1840,
      "Beam_Iy": 155,
      "Beam_Wz": 184,
      "Beam_Wy": 24.2,
      "Beam_Sz": 104,
      "Beam_Sy": 15.6,
      "Beam_iz": 7.75,
      "Beam_iy": 2.25
    }
  }
}
```

## 📂 Этап 3: Объединение JSON

### Входные данные

```
JSON_OUT/
├── M_P.json     [50 записей]
├── M_BK.json    [30 записей]
├── M_SK.json    [20 записей]
└── ...          [10 файлов]
```

### Процесс

```csharp
JsonMerger.MergeAll(jsonOutDir)

merged = new JsonArray()

foreach (file in Directory.GetFiles(jsonOutDir, "*.json"))
{
    arr = JsonNode.Parse(File.ReadAllText(file))
    
    foreach (item in arr)
    {
        merged.Add(item)
    }
}

return merged  // 100+ записей
```

### Выходные данные этапа 3

```
JSON_All/
└── all.json    [100+ записей]
```

## 📂 Этап 3.5: Обогащение неполных записей

### Входные данные

```json
// all.json содержит:
[
  {
    "CONNECTION_CODE": "M-3",
    "Geometry": { "Beam": {...}, "Plate": {...}, "Flange": {...} },
    "Bolts": { "F": 20, "CoordinatesBolts": {...} },
    "Welds": { "kf1": 6, "kf2": 6 },
    "TableBrand": "Т1",
    "InternalForces": { "Nt": 100, "My": 50 },
    "Coefficients": { "Alpha": 1.0, "Beta": 1.0 }
  },
  {
    "CONNECTION_CODE": "M-3",  // Тот же код!
    "Geometry": {},            // Пусто!
    "Bolts": {},               // Пусто!
    "Welds": {},               // Пусто!
    "TableBrand": "",          // Пусто!
    "InternalForces": { "Nt": 200, "My": 80 },  // Другие значения
    "Coefficients": { "Alpha": 1.1, "Beta": 1.2 }
  }
]
```

### Процесс

```csharp
JsonRecordEnricher.Enrich(merged)

// 1. Группировка по CONNECTION_CODE
groups = {
    "M-3": [index 0, index 1],
    "M-4": [index 5],
    ...
}

// 2. Для каждой группы с 2+ элементами
foreach (var (code, indices) in groups)
{
    // 3. Найти наиболее полную запись (template)
    templateIdx = 0
    templateScore = ScoreCompleteness(merged[0])
    
    for (i = 1; i < indices.Count; i++)
    {
        score = ScoreCompleteness(merged[indices[i]])
        if (score > templateScore)
        {
            templateIdx = indices[i]
            templateScore = score
        }
    }
    
    template = merged[templateIdx]
    
    // 4. Обогатить остальные записи
    foreach (idx in indices)
    {
        if (idx == templateIdx) continue
        
        target = merged[idx]
        
        // Копировать блоки
        DeepCopy(template["Geometry"], target["Geometry"])
        DeepCopy(template["Bolts"], target["Bolts"])
        DeepCopy(template["Welds"], target["Welds"])
        
        // Копировать TableBrand
        if (IsEmpty(target["TableBrand"]))
            target["TableBrand"] = template["TableBrand"]
    }
}
```

### Оценка полноты записи

```csharp
ScoreCompleteness(obj)

score = 0

// Geometry
score += CountNonZeroValues(obj["Geometry"]["Column"])
score += CountNonZeroValues(obj["Geometry"]["Plate"])
score += CountNonZeroValues(obj["Geometry"]["Flange"])
score += CountNonZeroValues(obj["Geometry"]["Stiff"])

// Bolts
if (obj["Bolts"]["DiameterBolt"]["F"] != 0)
    score += 10
score += CountNonZeroValues(obj["Bolts"]["CoordinatesBolts"])

// Welds
score += CountNonZeroValues(obj["Welds"])

return score
```

### Выходные данные этапа 3.5

```json
// Обогащённый all.json:
[
  {
    "CONNECTION_CODE": "M-3",
    "Geometry": { ... },       // Полностью заполнено
    "Bolts": { ... },          // Полностью заполнено
    "Welds": { ... },          // Полностью заполнено
    "TableBrand": "Т1",        // Заполнено
    "InternalForces": { "Nt": 100, "My": 50 },
    "Coefficients": { "Alpha": 1.0, "Beta": 1.0 }
  },
  {
    "CONNECTION_CODE": "M-3",
    "Geometry": { ... },       // СКОПИРОВАНО ИЗ ШАБЛОНА!
    "Bolts": { ... },          // СКОПИРОВАНО ИЗ ШАБЛОНА!
    "Welds": { ... },          // СКОПИРОВАНО ИЗ ШАБЛОНА!
    "TableBrand": "Т1",        // СКОПИРОВАНО ИЗ ШАБЛОНА!
    "InternalForces": { "Nt": 200, "My": 80 },  // ОСТАЛОСЬ СВОЁ
    "Coefficients": { "Alpha": 1.1, "Beta": 1.2 }
  }
]
```

## 📂 Этап 7: Устранение дубликатов CONNECTION_CODE

### Входные данные

```json
// all.json содержит дубликаты:
[
  { "CONNECTION_CODE": "M-3", ... },
  { "CONNECTION_CODE": "M-3", ... },  // Дубликат
  { "CONNECTION_CODE": "M-3", ... },  // Дубликат
  { "CONNECTION_CODE": "M-4", ... },
  { "CONNECTION_CODE": "M-5", ... },
  { "CONNECTION_CODE": "M-5", ... },  // Дубликат
]
```

### Процесс

```csharp
ConnectionCodeDeduplicator.CreateDeduplicatedJson(...)

// 1. Подсчёт количества каждого кода
countsByCode = {
    "M-3": 3,
    "M-4": 1,
    "M-5": 2
}

// 2. Найти максимальный номер для каждого префикса
maxByPrefix = {
    "M": 5
}

// 3. Для каждой записи
seen = new HashSet<string>()
changed = 0

foreach (obj in clonedArr)
{
    code = obj["CONNECTION_CODE"]
    
    // Первое вхождение — оставляем
    if (seen.Add(code))
        continue
    
    // Дубликат — переименовываем
    prefix = "M"  // Из "M-3"
    max = maxByPrefix["M"]  // 5
    
    do
    {
        max++
        newCode = "M-" + max  // "M-6", "M-7", ...
    }
    while (seen.Contains(newCode) || countsByCode.ContainsKey(newCode))
    
    obj["CONNECTION_CODE"] = newCode
    seen.Add(newCode)
    maxByPrefix["M"] = max
    changed++
}

// Сохранить all_NotDuplicate.json
```

### Выходные данные этапа 7

```json
// all_NotDuplicate.json:
[
  { "CONNECTION_CODE": "M-3", ... },  // Оригинал
  { "CONNECTION_CODE": "M-6", ... },  // Было M-3
  { "CONNECTION_CODE": "M-7", ... },  // Было M-3
  { "CONNECTION_CODE": "M-4", ... },
  { "CONNECTION_CODE": "M-5", ... },  // Оригинал
  { "CONNECTION_CODE": "M-8", ... },  // Было M-5
]
```

Отчёт в `CONNECTION_CODE_replacements.txt`:
```
M-3 => M-6
M-3 => M-7
M-5 => M-8
```

## 📂 Этапы 4-9: Экспорт списков и справочников

### Этап 4: Создание списков

```
profile.txt                CONNECTION_CODE.txt
───────────                ───────────────────
20Б1                       M-3
30Б1                       M-4
20К1                       M-5
30К2                       ...
...
```

### Этап 5: Создание JSON справочников

```json
// ProfileBeam.json
[
  {
    "CONNECTION_GUID": "550e8400-e29b-41d4-a716-446655440000",
    "Profile": "20Б1"
  },
  {
    "CONNECTION_GUID": "6ba7b810-9dad-11d1-80b4-00c04fd430c8",
    "Profile": "30Б1"
  }
]

// CONNECTION_CODE.json
[
  {
    "CONNECTION_GUID": "7c9e6679-7425-40de-944b-e07fc1f90ae7",
    "CONNECTION_CODE": "M-3"
  },
  {
    "CONNECTION_GUID": "8d8e7780-8536-51ef-a947-f18ed2f01bf8",
    "CONNECTION_CODE": "M-4"
  }
]
```

### Этап 9: Экспорт имён соединений

```json
// NameConnections.json
[
  {
    "NAME_GUID": "9f9f8891-9647-62f0-b958-029fe3f02cg9",
    "Name": "Балка-Колонна M"
  },
  {
    "NAME_GUID": "a0a09902-a758-73f1-c069-13af14f03dha",
    "Name": "Балка-Балка B"
  }
]
```

## 🔄 Полный цикл данных

```
Excel Files (.xls/.xlsx)
    │
    │ [1] Read & Parse
    ▼
List<Row> (In Memory)
    │
    │ [2] Apply Profiles
    ▼
Enriched List<Row>
    │
    │ [3] Write Individual JSON
    ▼
JSON_OUT/*.json
    │
    │ [4] Merge All
    ▼
JsonArray (all.json)
    │
    │ [5] Enrich Incomplete
    ▼
Enriched JsonArray
    │
    │ [6] Save
    ▼
JSON_All/all.json
    │
    │ [7] Deduplicate
    ▼
JSON_All/all_NotDuplicate.json
    │
    │ [8-9] Export
    ▼
EXCEL_Profile_OUT/
├── CONNECTION_CODE_new.json
├── Profile.json
└── NameConnections.json
```

## 📊 Статистика потока

### Типичная статистика обработки

```
Входных Excel файлов: 10
Записей прочитано: 120
Применено профилей: 120
Индивидуальных JSON: 10
Объединено записей: 120
Обогащено записей: 35
Найдено дубликатов: 8
Заменено кодов: 8
Уникальных профилей: 45
Уникальных кодов: 112
Уникальных имён: 8
```

---

**См. также:**
- [Архитектура проекта](Architecture.md)
- [UML диаграммы](UML.md)
