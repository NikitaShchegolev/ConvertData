using System;

namespace ConvertData.Entitys;

/// <summary>
/// Элемент списка кодов соединений с уникальным идентификатором.
/// Используется для экспорта в JSON-файл CONNECTION_CODE.json.
/// </summary>
internal sealed class ConnectionCodeItem
{
    /// <summary>
    /// Уникальный идентификатор соединения (GUID).
    /// </summary>
    public Guid CONNECTION_GUID { get; set; }
    /// <summary>
    /// Код соединения (CONNECTION_CODE).
    /// </summary>
    public string CONNECTION_CODE { get; set; } = "";
}
