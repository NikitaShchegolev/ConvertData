using System;
using System.Collections.Generic;
using System.Linq;
using ConvertData.Domain;

namespace ConvertData.Infrastructure;

/// <summary>
/// Добавляет суффиксы _1, _2, ... к дублирующимся значениям CONNECTION_CODE в списке Row.
/// Обрабатывает список строк, модифицируя свойство CONNECTION_CODE у дубликатов.
/// </summary>
internal sealed class ConnectionCodeSuffixAppender
{
    /// <summary>
    /// Обрабатывает список строк, добавляя суффиксы к дублирующимся CONNECTION_CODE.
    /// </summary>
    /// <param name="rows">Список строк (будет модифицирован на месте).</param>
    /// <returns>Количество измененных строк.</returns>
    public int Process(List<Row> rows)
    {
        if (rows == null || rows.Count == 0)
            return 0;

        var changed = 0;
        var codeCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var usedCodes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // Первый проход: подсчет количества каждого кода
        foreach (var row in rows)
        {
            var code = row.CONNECTION_CODE?.Trim() ?? string.Empty;
            if (string.IsNullOrEmpty(code))
                continue;

            if (codeCounts.ContainsKey(code))
                codeCounts[code]++;
            else
                codeCounts[code] = 1;
        }

        // Второй проход: добавление суффиксов
        foreach (var row in rows)
        {
            var originalCode = row.CONNECTION_CODE?.Trim() ?? string.Empty;
            if (string.IsNullOrEmpty(originalCode))
                continue;

            // Если код встречается только один раз, оставляем как есть
            if (codeCounts[originalCode] == 1)
            {
                usedCodes.Add(originalCode);
                continue;
            }

            // Ищем уникальный вариант с суффиксом
            string newCode = originalCode;
            int suffix = 1;
            while (usedCodes.Contains(newCode) || (suffix == 1 && codeCounts[originalCode] > 1))
            {
                newCode = $"{originalCode}_{suffix}";
                suffix++;
            }

            if (newCode != originalCode)
            {
                row.CONNECTION_CODE = newCode;
                changed++;
            }

            usedCodes.Add(newCode);
        }

        return changed;
    }
}