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
