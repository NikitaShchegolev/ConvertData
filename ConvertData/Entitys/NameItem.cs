using System;

namespace ConvertData.Entitys;

/// <summary>
/// Элемент списка имён соединений с уникальным идентификатором.
/// Используется для экспорта в JSON-файл NameConnections.json.
/// </summary>
internal sealed class NameItem
{
    /// <summary>
    /// Уникальный идентификатор имени (GUID).
    /// </summary>
    public Guid NAME_GUID { get; set; }
    /// <summary>
    /// Имя группы узлового соединения.
    /// </summary>
    public string Name { get; set; } = "";
}
