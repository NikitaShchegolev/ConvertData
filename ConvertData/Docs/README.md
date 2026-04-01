# ConvertData - Конвертер данных узловых соединений

## 📋 Описание проекта

**ConvertData** — консольное приложение для конвертации данных о балочных узловых соединениях из файлов Excel в структурированный JSON формат. Предназначено для использования в расчётах строительных конструкций.

## 🎯 Основные возможности

- ✅ Чтение данных из Excel файлов (.xls, .xlsx)
- ✅ Автоматическая конвертация .xls → .xlsx через COM Interop
- ✅ Поддержка многолистовых Excel книг (main, geometry, bolts, weld)
- ✅ Применение справочника профилей балок и колонн
- ✅ Обогащение неполных записей данными из эталонных записей
- ✅ Автоматическое устранение дубликатов CONNECTION_CODE
- ✅ Экспорт в JSON с форматированием и UTF-8 без BOM
- ✅ Поддержка русских и английских заголовков
- ✅ Парсинг числовых значений в русском и инвариантном форматах

## 📁 Структура проекта

```
ConvertData/
├── Application/          # Бизнес-логика приложения
│   ├── ConvertApp.cs                    # Главный оркестратор процесса конвертации
│   ├── JsonRecordEnricher.cs            # Обогащение неполных записей
│   ├── ConnectionCodeDeduplicator.cs    # Устранение дубликатов кодов
│   ├── ProfileExcelToJsonExporter.cs    # Экспорт справочника профилей
│   ├── ProfileLookupLoader.cs           # Загрузка справочника профилей
│   ├── JsonProfilePatcher.cs            # Применение профилей к данным
│   └── ...
├── Domain/               # Доменные модели
│   ├── Row.cs                           # Модель записи соединения
│   ├── CoordinatesBolts.cs              # Координаты болтов
│   └── ProfileGeometry.cs               # Геометрия профиля
├── Infrastructure/       # Инфраструктурный слой
│   ├── EpplusRowReader.cs               # Чтение Excel через EPPlus
│   ├── JsonRowWriter.cs                 # Запись в JSON
│   ├── ExcelHeaderResolver.cs           # Разрешение заголовков Excel
│   ├── RowMapper.cs                     # Маппинг данных
│   ├── JsonMerger.cs                    # Объединение JSON файлов
│   ├── Parsing/                         # Парсинг данных
│   │   ├── NumericParser.cs             # Парсинг чисел
│   │   ├── HeaderUtils.cs               # Утилиты заголовков
│   │   └── ExcelFileSignature.cs        # Определение формата файла
│   └── Interop/
│       └── ExcelXlsConverter.cs         # Конвертация .xls через COM
├── Entitys/              # Сущности для экспорта
│   ├── ConnectionCodeItem.cs
│   ├── ProfileItem.cs
│   └── NameItem.cs
├── EXCEL/                # Входные Excel файлы
├── EXCEL_Profile/        # Справочники профилей
├── JSON_OUT/             # Индивидуальные JSON файлы
├── JSON_All/             # Объединённые результаты
├── EXCEL_Profile_OUT/    # Экспортированные справочники
└── Docs/                 # Документация проекта

```

## 🚀 Использование

### Запуск приложения

```bash
# Полный цикл конвертации (все этапы)
ConvertData.exe

# Только создание JSON из Excel
ConvertData.exe 1

# Только применение профилей к существующим JSON
ConvertData.exe 2

# Указание конкретных файлов
ConvertData.exe path\to\file1.xls path\to\file2.xlsx

# Переопределение колонки профиля
ConvertData.exe --profile-column=CustomProfileColumn
```

### Режимы работы

| Режим | Команда | Описание |
|-------|---------|----------|
| **All** | `ConvertData.exe` | Выполняет все этапы конвертации |
| **CreateJson** | `ConvertData.exe 1` | Создаёт JSON из Excel, пропускает применение профилей |
| **ApplyProfiles** | `ConvertData.exe 2` | Применяет справочник профилей к существующим JSON |

## 📊 Этапы конвертации

### Этап 1: Создание JSON из Excel
- Чтение файлов из папки `EXCEL/`
- Парсинг заголовков и данных
- Объединение данных из листов (main, geometry, bolts, weld)
- Запись индивидуальных JSON в `JSON_OUT/`

### Этап 2: Применение справочника профилей
- Загрузка справочника из `EXCEL_Profile/ProfileBeam.xls`
- Обогащение записей геометрическими характеристиками (H, B, s, t, A, P, Iz, Iy, Wz, Wy, iz, iy)
- Экспорт справочника в `EXCEL_Profile_OUT/Profile.json`

### Этап 3: Объединение JSON
- Слияние всех файлов из `JSON_OUT/` в `all.json`
- Обогащение неполных записей данными от записей с тем же CONNECTION_CODE

### Этап 4-5: Экспорт списков
- Создание `profile.txt` и `CONNECTION_CODE.txt`
- Создание `ProfileBeam.json` и `CONNECTION_CODE.json`

### Этап 6: Проверка дубликатов
- Поиск дубликатов CONNECTION_CODE
- Запись отчёта в `CONNECTION_CODE_duplicates.txt`

### Этап 7: Устранение дубликатов
- Переименование дубликатов (M-3 → M-3, M-4, M-5...)
- Создание `all_NotDuplicate.json`
- Запись отчёта замен в `CONNECTION_CODE_replacements.txt`

### Этап 8: Финальная проверка
- Экспорт `CONNECTION_CODE_new.json` и `CONNECTION_CODE_new.txt`
- Проверка отсутствия дубликатов

### Этап 9: Экспорт имён
- Создание `NameConnections.json` с уникальными именами соединений

## 📄 Формат данных

### Входные Excel файлы

#### Основной лист (Main)
Обязательные колонки:
- `Name` — имя группы соединений
- `CONNECTION_CODE` / `Code` / `Код` — уникальный код соединения
- `ProfileBeam` / `Профиль` — профиль балки
- `variable` — вариант расчёта

Опциональные колонки геометрии:
- `Beam_H`, `Beam_B`, `Beam_s`, `Beam_t` — размеры балки
- `ProfileColumn` — профиль колонны

Усилия и моменты:
- `Nt`, `Nc`, `N` — растяжение, сжатие, комбинированное усилие
- `Qy`, `Qz`, `Qx` — поперечные силы
- `My`, `Mz`, `Mx`, `Mw` — изгибающие и крутящий моменты
- `Mneg` — обратный момент
- `T` — кручение

Жёсткости и коэффициенты:
- `Sj`, `Sjo` — жёсткости
- `α`, `β`, `γ`, `δ`, `ε`, `λ` (или Alpha, Beta, Gamma, Delta, Epsilon, Lambda)

#### Лист Geometry (опционально)
Геометрические параметры пластин, фланцев, рёбер жёсткости:
- `CONNECTION_CODE` — ключ связи
- `H`, `B`, `tp` — размеры пластины/фланца
- `Lb` — расстояние для фланца
- `tbp`, `tg`, `tf`, `Lh`, `Hh`, `tr1`, `tr2`, `twp` — параметры рёбер жёсткости

#### Лист Bolts (опционально)
Параметры болтовых соединений:
- `CONNECTION_CODE` — ключ связи
- `Option` — версия болтов
- `F` — диаметр болта
- `Nb` — количество болтов
- `e1` — координата Y первого болта
- `d1`, `d2` — координаты X рядов болтов
- `p1`-`p10` — расстояния между рядами
- `Марка опорного столика` — марка опорного столика

#### Лист Weld (опционально)
Параметры сварных швов:
- `CONNECTION_CODE` — ключ связи
- `kf1`-`kf10` — минимальные катеты сварных швов

### Выходной JSON формат

```json
{
  "Name": "Балка-Колонна Тип 1",
  "CONNECTION_CODE": "M-3",
  "variable": 1,
  "TableBrand": "Т1",
  
  "Stiffness": {
    "Sj": 0,
    "Sjo": 0
  },
  
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
      ...
    },
    "Column": { ... },
    "Plate": { ... },
    "Flange": { ... },
    "Stiff": { ... }
  },
  
  "Bolts": {
    "Option": { "version": 1 },
    "DiameterBolt": { "F": 20 },
    "CountBolt": { "Bolts_Nb": 4 },
    "BoltRow": { "N_Rows": 2 },
    "CoordinatesBolts": {
      "Y": {
        "Bolt1_e1": 50,
        "Bolt2_p1": 70,
        ...
      },
      "X": {
        "d1": 60,
        "d2": 120
      },
      "Z": {
        "BoltCoordinateZ": 0
      }
    }
  },
  
  "Welds": {
    "kf1": 6,
    "kf2": 6,
    ...
  },
  
  "InternalForces": {
    "N": 0,
    "Nt": 100,
    "Nc": 0,
    "My": 50,
    "Mz": 0,
    ...
  },
  
  "Coefficients": {
    "Alpha": 1.0,
    "Beta": 1.0,
    "Gamma": 1.0,
    "Delta": 1.0,
    "Epsilon": 1.0,
    "Lambda": 1.0
  }
}
```

## 🔧 Технологии

- **.NET 10.0** — целевая платформа
- **C# 14.0** — язык программирования (с collection expressions)
- **EPPlus 7.x** — чтение/запись Excel (.xlsx)
- **COM Interop (Microsoft.Office.Interop.Excel)** — конвертация .xls → .xlsx
- **System.Text.Json** — работа с JSON

## 📝 Особенности реализации

### Парсинг чисел
Поддержка русского (запятая) и инвариантного (точка) форматов:
```csharp
"3,14" → 3.14
"3.14" → 3.14
```

### Обогащение записей
Если несколько записей имеют одинаковый `CONNECTION_CODE`, но разную полноту данных, система автоматически копирует Geometry, Bolts, Welds из наиболее полной записи в неполные. InternalForces и Coefficients остаются индивидуальными.

### Устранение дубликатов
Автоматическое переименование:
```
M-3  (оригинал)
M-3  (дубликат) → M-4
M-3  (дубликат) → M-5
```

### Греческие символы
Поддержка как Unicode символов (α, β, γ), так и латинских названий (Alpha, Beta, Gamma).

## 🐛 Известные ограничения

- Требует наличия Microsoft Excel для конвертации .xls файлов
- Обрабатывает только первый лист Excel книги как основной
- Максимум 11 болтов по оси Y (e1, p1-p10)
- Максимум 2 ряда болтов по оси X (d1, d2)

## 📞 Контакты

- GitHub: [NikitaShchegolev/ConvertData](https://github.com/NikitaShchegolev/ConvertData)
- Автор: Nikita Shchegolev

## 📜 Лицензия

Проект использует EPPlus под лицензией NonCommercial.

---

Для более подробной информации см.:
- [Архитектура проекта](Architecture.md)
- [Поток данных](DataFlow.md)
- [UML диаграммы](UML.md)
- [API документация](API.md)
