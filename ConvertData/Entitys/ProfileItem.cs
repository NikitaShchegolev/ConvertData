using System;

namespace ConvertData.Entitys;

/// <summary>
/// Элемент списка профилей балок с уникальным идентификатором.
/// Используется для экспорта в JSON-файл ProfileBeam.json.
/// </summary>
internal sealed class ProfileItem
{
    /// <summary>
    /// Уникальный идентификатор профиля (GUID).
    /// </summary>
    public Guid CONNECTION_GUID { get; set; }
    /// <summary>
    /// Название профиля балки (например, "20Б1", "30К1").
    /// </summary>
    public string Profile { get; set; } = "";
}
