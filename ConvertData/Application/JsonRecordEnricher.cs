using System;
using System.Collections.Generic;
using System.Text.Json.Nodes;

namespace ConvertData.Application;

/// <summary>
/// Обогащает неполные записи в JSON-массиве данными от записей с тем же CONNECTION_CODE.
/// Копирует Geometry (Column, Plate, Flange, Stiff), Bolts, Welds из наиболее полной записи-шаблона.
/// InternalForces и Coefficients остаются индивидуальными для каждой записи.
/// </summary>
internal sealed class JsonRecordEnricher
{
    private static readonly string[] GeometrySubSections = ["Column", "Plate", "Flange", "Stiff"];

    public int Enrich(JsonArray arr)
    {
        if (arr == null || arr.Count == 0)
            return 0;

        var groups = new Dictionary<string, List<int>>(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < arr.Count; i++)
        {
            if (arr[i] is not JsonObject obj)
                continue;
            var code = obj["CONNECTION_CODE"]?.GetValue<string>()?.Trim();
            if (string.IsNullOrWhiteSpace(code))
                continue;
            if (!groups.TryGetValue(code, out var list))
            {
                list = [];
                groups[code] = list;
            }
            list.Add(i);
        }

        int enriched = 0;

        foreach (var (code, indices) in groups)
        {
            if (indices.Count < 2)
                continue;

            int templateIdx = indices[0];
            int templateScore = ScoreCompleteness(arr[templateIdx] as JsonObject);

            for (int k = 1; k < indices.Count; k++)
            {
                int score = ScoreCompleteness(arr[indices[k]] as JsonObject);
                if (score > templateScore)
                {
                    templateIdx = indices[k];
                    templateScore = score;
                }
            }

            if (templateScore == 0)
                continue;

            var template = (JsonObject)arr[templateIdx]!;

            foreach (var idx in indices)
            {
                if (idx == templateIdx)
                    continue;

                var obj = arr[idx] as JsonObject;
                if (obj == null)
                    continue;

                if (ScoreCompleteness(obj) >= templateScore)
                    continue;

                EnrichRecord(template, obj);
                enriched++;
            }
        }

        if (enriched == 0)
            return 0;

        return enriched;
    }

    private static void EnrichRecord(JsonObject template, JsonObject target)
    {
        var tGeom = template["Geometry"] as JsonObject;
        var oGeom = target  ["Geometry"] as JsonObject;
        if (tGeom != null && oGeom != null)
        {
            foreach (var section in GeometrySubSections)
                DeepCopyNode(tGeom, oGeom, section);
        }

        DeepCopyNode(template, target, "Bolts");
        DeepCopyNode(template, target, "Welds");

        CopyStringIfEmpty(template,target,"TableBrand");
        CopyStringIfEmpty(template,target,"Explanations");
        CopyStringIfEmpty(template,target,"TypeNode");

        #region Delete
        //// Копируем TableBrand, если он есть в шаблоне и пуст в целевой записи
        //if (template["TableBrand"] is JsonNode brandNode)
        //{
        //    var brandValue = brandNode.GetValue<string>();
        //    if (!string.IsNullOrWhiteSpace(brandValue))
        //    {
        //        var targetBrand = target["TableBrand"]?.GetValue<string>();
        //        if (string.IsNullOrWhiteSpace(targetBrand))
        //            target["TableBrand"] = brandValue;
        //    }
        //} 
        #endregion
    }
    /// <summary>
    /// Копирует строковое значение из шаблона в целевой объект, если в целевом оно пусто.
    /// </summary>
    /// <param name="template">Шаблонный JSON-объект.</param>
    /// <param name="target">Целевой JSON-объект.</param>
    /// <param name="key">Ключ строки для копирования.</param>
    private static void CopyStringIfEmpty(JsonObject template, JsonObject target, string key)
    {
        if (template[key] is JsonNode srcNode)
        {
            var srcValue = srcNode.GetValue<string>();
            if (!string.IsNullOrWhiteSpace(srcValue))
            {
                var targetValue = target[key]?.GetValue<string>();
                if (string.IsNullOrWhiteSpace(targetValue))
                    target[key] = srcValue;
            }
        }
    }

    /// <summary>
    /// Выполняет глубокое копирование узла JSON из источника в целевой объект.
    /// </summary>
    /// <param name="source">Исходный JSON-объект.</param>
    /// <param name="target">Целевой JSON-объект.</param>
    /// <param name="key">Ключ узла для копирования.</param>
    private static void DeepCopyNode(JsonObject source, JsonObject target, string key)
    {
        var src = source[key];
        if (src == null)
            return;
        target[key] = JsonNode.Parse(src.ToJsonString());
    }

    private static int ScoreCompleteness(JsonObject? obj)
    {
        if (obj == null)
            return 0;

        int score = 0;

        if (obj["Geometry"] is JsonObject geom)
        {
            foreach (var section in GeometrySubSections)
                score += CountNonZeroValues(geom[section] as JsonObject);
        }

        if (obj["Bolts"] is JsonObject bolts)
        {
            var f = bolts["DiameterBolt"]?["F"];
            if (f != null && GetNumericValue(f) != 0)
                score += 10;
            score += CountNonZeroValues(bolts["CoordinatesBolts"] as JsonObject);
        }

        score += CountNonZeroValues(obj["Welds"] as JsonObject);

        return score;
    }

    private static int CountNonZeroValues(JsonObject? obj)
    {
        if (obj == null)
            return 0;

        int count = 0;
        foreach (var prop in obj)
        {
            if (prop.Value is JsonValue val && GetNumericValue(val) != 0)
                count++;
            else if (prop.Value is JsonObject nested)
                count += CountNonZeroValues(nested);
        }
        return count;
    }

    /// <summary>
    /// Извлекает числовое значение из JSON-узла.
    /// </summary>
    /// <param name="node">JSON-узел.</param>
    /// <returns>Числовое значение (double или int) или 0, если узел не содержит число.</returns>
    private static double GetNumericValue(JsonNode? node)
    {
        if (node is not JsonValue val)
            return 0;
        if (val.TryGetValue<double>(out var d))
            return d;
        if (val.TryGetValue<int>(out var i))
            return i;
        return 0;
    }
}
