namespace ConvertData.Application;

/// <summary>
/// Режим выполнения приложения, определяющий какие этапы конвертации будут выполнены.
/// </summary>
internal enum RunMode
{
    /// <summary>
    /// Выполнить все этапы конвертации: создание JSON, применение профилей и объединение.
    /// </summary>
    All,
    /// <summary>
    /// Только создать JSON-файлы из Excel без применения профилей.
    /// </summary>
    CreateJson,
    /// <summary>
    /// Только применить справочник профилей к существующим JSON-файлам.
    /// </summary>
    ApplyProfiles
}

/// <summary>
/// Логические блоки этапов для группового управления выполнением.
/// </summary>
[Flags]
internal enum Block
{
    /// <summary>
    /// Ни один блок не выполняется.
    /// </summary>
    None = 0,
    
    /// <summary>
    /// Создание JSON из Excel (без профилей) - этап 1.
    /// </summary>
    CreateJson = 1,
    
    /// <summary>
    /// Применение справочника профилей - этап 2.
    /// </summary>
    ApplyProfiles = 2,
    
    /// <summary>
    /// Объединение всех JSON в один файл и обогащение записей - этапы 3-4.
    /// </summary>
    MergeAndEnrich = 4,
    
    /// <summary>
    /// Экспорт профилей и кодов (profile.txt, CONNECTION_CODE.txt, ProfileBeam.json, CONNECTION_CODE.json) - этапы 5-6.
    /// </summary>
    ExportProfiles = 8,
    
    /// <summary>
    /// Дедупликация: проверка дубликатов, создание all_NotDuplicate.json, экспорт новых кодов - этапы 7-9.
    /// </summary>
    Deduplication = 16,
    
    /// <summary>
    /// Копирование all_NotDuplicate.json в ConvertData.Data\JSON\ - этап 10.
    /// </summary>
    CopyToData = 32,
    
    /// <summary>
    /// Экспорт анкеров из Anchor.xlsx в JSON - этап 11.
    /// </summary>
    AnchorExport = 64,
    
    /// <summary>
    /// Экспорт анкеров из MarkSteel.xlsx в JSON - этап 12.
    /// </summary>
    SteelExport = 128,

    /// <summary>
    /// Экспорт болтов из TableBoltsSP43.xlsx в JSON - этап 13.
    /// </summary>
    BoltsExport = 129,
    
    /// <summary>
    /// Все блоки конвертации (CreateJson + ApplyProfiles).
    /// </summary>
    Conversion = CreateJson | ApplyProfiles,
    
    /// <summary>
    /// Все блоки обработки (MergeAndEnrich + ExportProfiles + Deduplication + CopyToData).
    /// </summary>
    Processing = MergeAndEnrich | ExportProfiles | Deduplication | CopyToData,
    
    /// <summary>
    /// Все блоки экспорта анкеров (AnchorExport + SteelExport).
    /// </summary>
    Anchors = AnchorExport | SteelExport,
    
    Bolts = BoltsExport,
    
    /// <summary>
    /// Все блоки (полный цикл обработки).
    /// </summary>
    All = Conversion | Processing | Anchors
}
