# Визуальные диаграммы ConvertData (Mermaid)

> **Примечание:** Эти диаграммы в формате Mermaid можно просматривать напрямую на GitHub. Для локального просмотра используйте [Mermaid Live Editor](https://mermaid.live) или VS Code расширение [Mermaid Preview](https://marketplace.visualstudio.com/items?itemName=bierner.markdown-mermaid).

## 📊 Общий поток данных

```mermaid
flowchart TD
    Start([Запуск ConvertData.exe]) --> ParseArgs[Парсинг аргументов]
    ParseArgs --> CheckMode{Режим работы?}
    
    CheckMode -->|All| Stage1[Этап 1-2: CreateJson + ApplyProfiles]
    CheckMode -->|CreateJson| Stage1
    CheckMode -->|ApplyProfiles| Stage2Only[Этап 2: ApplyProfiles]
    
    Stage1 --> ReadExcel[Чтение Excel файлов]
    ReadExcel --> DetectFormat{Формат?}
    DetectFormat -->|.xls| ConvertXLS[Конвертация .xls → .xlsx<br/>через COM Interop]
    DetectFormat -->|.xlsx| ReadXLSX[Чтение через EPPlus]
    ConvertXLS --> ReadXLSX
    
    ReadXLSX --> ParseHeaders[Разрешение заголовков]
    ParseHeaders --> MergeSheets[Объединение листов<br/>main, geometry, bolts, weld]
    MergeSheets --> CreateRows[Создание List&lt;Row&gt;]
    
    CreateRows --> WriteJSON[Запись в JSON_OUT/*.json]
    WriteJSON --> Stage2{Режим All?}
    
    Stage2 -->|Да| LoadProfiles[Загрузка справочника<br/>ProfileBeam.xls]
    Stage2Only --> LoadProfiles
    LoadProfiles --> ApplyProfiles[Применение профилей<br/>к JSON файлам]
    
    Stage2 -->|Нет| EndCreate([Завершено: CreateJson])
    
    ApplyProfiles --> Stage3{Режим All?}
    Stage3 -->|Нет| EndApply([Завершено: ApplyProfiles])
    Stage3 -->|Да| MergeAll[Объединение в all.json]
    
    MergeAll --> Enrich[Обогащение неполных записей]
    Enrich --> SaveAll[Сохранение all.json]
    SaveAll --> Export[Экспорт списков]
    Export --> CheckDupes[Проверка дубликатов]
    CheckDupes --> Dedup[Устранение дубликатов]
    Dedup --> SaveFinal[Сохранение all_NotDuplicate.json]
    SaveFinal --> ExportFinal[Финальный экспорт справочников]
    ExportFinal --> End([Завершено])
    
    style Start fill:#90EE90
    style End fill:#FFB6C1
    style EndCreate fill:#FFB6C1
    style EndApply fill:#FFB6C1
```

## 🏗️ Архитектура слоёв

```mermaid
flowchart LR
    subgraph Presentation
        Program[Program.cs<br/>Main Entry Point]
    end
    
    subgraph Application["Application Layer"]
        ConvertApp[ConvertApp<br/>Orchestrator]
        Enricher[JsonRecordEnricher]
        Dedup[ConnectionCodeDeduplicator]
        ProfileExporter[ProfileExcelToJsonExporter]
        ProfileLoader[ProfileLookupLoader]
        ProfilePatcher[JsonProfilePatcher]
    end
    
    subgraph Domain["Domain Layer"]
        Row[Row<br/>Main Entity]
        Bolts[CoordinatesBolts]
        ProfileGeom[ProfileGeometry]
    end
    
    subgraph Infrastructure["Infrastructure Layer"]
        Reader[EpplusRowReader]
        Writer[JsonRowWriter]
        Mapper[RowMapper]
        Parser[NumericParser]
        HeaderResolver[ExcelHeaderResolver]
    end
    
    subgraph External["External Dependencies"]
        EPPlus[EPPlus 7.x]
        JSON[System.Text.Json]
        COM[COM Interop Excel]
    end
    
    Program --> ConvertApp
    ConvertApp --> Enricher
    ConvertApp --> Dedup
    ConvertApp --> ProfileExporter
    ConvertApp --> ProfileLoader
    ConvertApp --> ProfilePatcher
    ConvertApp --> Reader
    ConvertApp --> Writer
    
    Reader --> Mapper
    Reader --> HeaderResolver
    Reader --> Parser
    Reader --> Row
    Reader --> Bolts
    
    Mapper --> Row
    Mapper --> Parser
    
    Writer --> Row
    
    ProfileLoader --> ProfileGeom
    
    Reader --> EPPlus
    Reader --> COM
    Enricher --> JSON
    Dedup --> JSON
    ProfilePatcher --> JSON
    
    style Program fill:#E1F5FE
    style ConvertApp fill:#B3E5FC
    style Row fill:#FFF9C4
    style Reader fill:#C8E6C9
    style EPPlus fill:#FFCCBC
```

## 🔄 Последовательность обогащения записей

```mermaid
sequenceDiagram
    participant Main as ConvertApp
    participant Enricher as JsonRecordEnricher
    participant JSON as JsonArray
    
    Main->>Enricher: Enrich(jsonArray)
    activate Enricher
    
    Enricher->>JSON: Группировка по CONNECTION_CODE
    Note over JSON: M-3: [0, 1, 2]<br/>M-4: [3]<br/>M-5: [4, 5]
    
    loop Для каждой группы с 2+ элементами
        Enricher->>JSON: Найти наиболее полную запись (template)
        Note over JSON: ScoreCompleteness:<br/>idx 0 = 150<br/>idx 1 = 50<br/>idx 2 = 100<br/>→ template = idx 0
        
        loop Для каждой неполной записи
            Enricher->>JSON: DeepCopy Geometry
            Enricher->>JSON: DeepCopy Bolts
            Enricher->>JSON: DeepCopy Welds
            Enricher->>JSON: Копировать TableBrand
            Note over JSON: InternalForces и Coefficients<br/>остаются уникальными
        end
    end
    
    Enricher-->>Main: Количество обогащённых записей
    deactivate Enricher
```

## 📦 Структура данных Row

```mermaid
classDiagram
    class Row {
        +string Name
        +string CONNECTION_CODE
        +int variable
        +string TableBrand
        
        +string ProfileBeam
        +double Beam_H, Beam_B, Beam_s, Beam_t
        +double Beam_A, Beam_P
        +double Beam_Iz, Beam_Iy, Beam_Ix
        +double Beam_Wz, Beam_Wy, Beam_Wx
        +double Beam_Sz, Beam_Sy
        +double Beam_iz, Beam_iy
        +double Beam_xo, Beam_yo
        
        +string ProfileColumn
        +double Column_H, Column_B, ...(14 свойств)
        
        +double Plate_H, Plate_B, Plate_t
        +double Flange_Lb, Flange_H, Flange_B, Flange_t
        +double Stiff_tbp, Stiff_tg, ...(8 свойств)
        
        +List~CoordinatesBolts~ CoordinatesBolts
        +int F, Bolts_Nb, N_Rows
        +double OptionBolts
        +int e1, d1, d2
        +double p1, p2, ..., p10
        
        +int kf1, kf2, ..., kf10
        
        +int Sj, Sjo
        
        +int Nt, Nc, N, Qy, Qz, Qx, My, T
        +double Mneg, Mz, Mx, Mw
        
        +double Alpha, Beta, Gamma, Delta, Epsilon, Lambda
    }
    
    class CoordinatesBolts {
        +int X
        +int Y
        +int Z
        +CoordinatesBolts(x, y, z)
    }
    
    Row "1" o-- "*" CoordinatesBolts : contains
```

## 🗂️ Схема Excel → JSON маппинга

```mermaid
graph TB
    subgraph Excel["Excel Workbook"]
        MainSheet[Main Sheet<br/>───────────<br/>Name, CODE, Profile<br/>variable, Sj, Sjo<br/>Nt, My, Qy, ...]
        GeomSheet[geometry Sheet<br/>───────────<br/>CODE, H, B, tp<br/>Lb, tbp, tg, ...]
        BoltsSheet[bolts Sheet<br/>───────────<br/>CODE, Option, F, Nb<br/>e1, d1, d2, p1-p10<br/>Марка опорного столика]
        WeldSheet[weld Sheet<br/>───────────<br/>CODE, kf1-kf10]
    end
    
    subgraph Mapping["RowMapper + EpplusRowReader"]
        MapMain[MapMainRow]
        MergeGeom[MergeSheet geometry]
        MergeBolts[MergeSheet bolts]
        MergeWeld[MergeSheet weld]
    end
    
    subgraph Domain["Row Object"]
        RowObj[Row<br/>───────────<br/>✓ Name, CONNECTION_CODE<br/>✓ ProfileBeam, variable<br/>✓ Beam geometry<br/>✓ Plate/Flange geometry<br/>✓ Bolts coordinates<br/>✓ Welds kf1-kf10<br/>✓ Internal Forces<br/>✓ Coefficients]
    end
    
    MainSheet -->|основные поля| MapMain
    MapMain --> RowObj
    
    GeomSheet -->|по CODE| MergeGeom
    MergeGeom -->|Plate_H, Flange_Lb, ...| RowObj
    
    BoltsSheet -->|по CODE| MergeBolts
    MergeBolts -->|F, e1, d1, d2, TableBrand| RowObj
    
    WeldSheet -->|по CODE| MergeWeld
    MergeWeld -->|kf1-kf10| RowObj
    
    style MainSheet fill:#E3F2FD
    style GeomSheet fill:#FFF3E0
    style BoltsSheet fill:#F3E5F5
    style WeldSheet fill:#E8F5E9
    style RowObj fill:#FFF9C4
```

## 🔍 Процесс устранения дубликатов

```mermaid
stateDiagram-v2
    [*] --> ReadJSON: Прочитать all.json
    
    ReadJSON --> CountCodes: Подсчитать количество<br/>каждого CODE
    
    CountCodes --> FindMax: Найти максимальные<br/>номера по префиксам
    
    FindMax --> ProcessRecords: Обработать записи
    
    state ProcessRecords {
        [*] --> CheckCode
        CheckCode --> FirstOccurrence: Первое вхождение?
        FirstOccurrence --> KeepCode: Да
        FirstOccurrence --> Rename: Нет (дубликат)
        
        Rename --> GenerateNew: max++<br/>newCode = prefix + max
        GenerateNew --> CheckExists: CODE уже существует?
        CheckExists --> GenerateNew: Да
        CheckExists --> ApplyNew: Нет
        
        ApplyNew --> LogReplacement: Записать в лог
        
        KeepCode --> [*]
        LogReplacement --> [*]
    }
    
    ProcessRecords --> SaveResult: Сохранить<br/>all_NotDuplicate.json
    
    SaveResult --> SaveReport: Сохранить<br/>replacements.txt
    
    SaveReport --> [*]
    
    note right of ProcessRecords
        Пример:
        M-3, M-3, M-3
        → M-3, M-6, M-7
        
        (макс. было M-5)
    end note
```

## 🌐 Граф зависимостей компонентов

```mermaid
graph LR
    Program --> ConvertApp
    
    ConvertApp --> IRowWriter
    ConvertApp --> IRowReaderFactory
    ConvertApp --> IPathResolver
    ConvertApp --> ILicenseConfigurator
    ConvertApp --> JsonRecordEnricher
    ConvertApp --> ConnectionCodeDeduplicator
    ConvertApp --> ProfileExcelToJsonExporter
    ConvertApp --> ProfileLookupLoader
    ConvertApp --> JsonProfilePatcher
    ConvertApp --> RunModeParser
    ConvertApp --> JsonMerger
    
    IRowWriter -.->|implements| JsonRowWriter
    IRowReaderFactory -.->|implements| RowReaderFactory
    IPathResolver -.->|implements| PathResolver
    ILicenseConfigurator -.->|implements| EpplusLicenseConfigurator
    
    RowReaderFactory -->|creates| EpplusRowReader
    
    EpplusRowReader --> ExcelHeaderResolver
    EpplusRowReader --> RowMapper
    EpplusRowReader --> NumericParser
    EpplusRowReader --> HeaderUtils
    EpplusRowReader --> ExcelFileSignature
    EpplusRowReader --> ExcelXlsConverter
    EpplusRowReader -->|creates| Row
    EpplusRowReader -->|creates| CoordinatesBolts
    
    RowMapper --> NumericParser
    RowMapper -->|creates| Row
    
    JsonRowWriter -->|reads| Row
    
    ProfileLookupLoader -->|creates| ProfileGeometry
    
    Row -->|contains| CoordinatesBolts
    
    style Program fill:#90EE90
    style ConvertApp fill:#87CEEB
    style Row fill:#FFD700
    style EpplusRowReader fill:#98FB98
    style JsonRowWriter fill:#DDA0DD
```

## 📈 Статистика обработки данных

```mermaid
pie title Типичное распределение записей
    "Полные записи" : 60
    "Частично заполненные" : 25
    "Только InternalForces" : 15
```

```mermaid
pie title Результат обогащения
    "Были полными" : 60
    "Обогащены" : 25
    "Только силы" : 15
```

```mermaid
journey
    title Путь записи через систему
    section Чтение
      Открыть Excel: 5: Reader
      Найти заголовки: 4: HeaderResolver
      Прочитать данные: 5: Mapper
      Объединить листы: 4: Reader
    section Обогащение
      Применить профили: 5: ProfilePatcher
      Объединить в all.json: 3: JsonMerger
      Обогатить неполные: 4: Enricher
    section Финализация
      Устранить дубликаты: 5: Deduplicator
      Экспорт справочников: 3: Exporter
```

## 🔐 Обработка ошибок

```mermaid
flowchart TD
    Start([Чтение файла]) --> TryRead{Попытка чтения}
    
    TryRead -->|Успех| CheckFormat{Формат корректен?}
    TryRead -->|IOException| HandleIO[Логировать:<br/>'Файл недоступен']
    
    CheckFormat -->|Да| ProcessData[Обработка данных]
    CheckFormat -->|Нет| ThrowInvalid[InvalidDataException:<br/>'Некорректный формат']
    
    ProcessData --> TryParse{Парсинг успешен?}
    TryParse -->|Да| Success([Успешно])
    TryParse -->|Нет| UseDefault[Использовать<br/>значения по умолчанию]
    
    UseDefault --> LogWarning[Логировать предупреждение]
    LogWarning --> Success
    
    HandleIO --> Skip([Пропустить файл])
    ThrowInvalid --> Skip
    
    style Success fill:#90EE90
    style Skip fill:#FFB6C1
    style HandleIO fill:#FFE4B5
    style ThrowInvalid fill:#FFE4B5
```

---

## 📝 Как использовать диаграммы

### GitHub
Диаграммы Mermaid автоматически рендерятся в README.md и других Markdown файлах на GitHub.

### VS Code
1. Установите расширение: **Markdown Preview Mermaid Support**
2. Откройте файл в режиме предпросмотра (Ctrl+Shift+V)

### Mermaid Live Editor
1. Откройте https://mermaid.live
2. Скопируйте код диаграммы
3. Вставьте в редактор
4. Экспортируйте как PNG/SVG

### CLI
```bash
npm install -g @mermaid-js/mermaid-cli
mmdc -i diagram.mmd -o diagram.png
```

---

**См. также:**
- [README](README.md)
- [Архитектура](Architecture.md)
- [Поток данных](DataFlow.md)
- [UML диаграммы](UML.md) — PlantUML версии
- [API документация](API.md)
