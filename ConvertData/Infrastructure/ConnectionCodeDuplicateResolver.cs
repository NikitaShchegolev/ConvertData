using System;
using System.Collections.Generic;
using System.Linq;

namespace ConvertData.Infrastructure;

/// <summary>
/// Разрешает дубликаты CONNECTION_CODE путём добавления суффиксов _1, _2, ...
/// </summary>
internal sealed class ConnectionCodeDuplicateResolver
{
    /// <summary>
    /// Режим обработки дубликатов.
    /// </summary>
    private enum DuplicateMode
    {
        /// <summary> Добавлять суффиксы к дубликатам, сохраняя оригиналы без изменений. </summary>
        AddSuffixToDuplicates
    }

    /// <summary>
    /// Обрабатывает массив кодов, заменяя дубликаты уникальными значениями.
    /// </summary>
    /// <param name="codes">Входной массив кодов (может содержать дубликаты).</param>
    /// <returns>Список той же длины, где дубликаты заменены на уникальные значения с суффиксами.</returns>
    public List<string> Process(IEnumerable<string> codes)
    {
        var inputList = codes?.ToList() ?? new List<string>();
        if (inputList.Count == 0)
            return new List<string>();

        // Выбор режима обработки (можно расширить)
        var mode = DuplicateMode.AddSuffixToDuplicates;
        switch (mode)
        {
            case DuplicateMode.AddSuffixToDuplicates:
                return ApplySuffixToDuplicates(inputList);
            default:
                throw new NotSupportedException($"Mode {mode} is not supported.");
        }
    }

    /// <summary>
    /// Применяет стратегию добавления суффиксов к дубликатам.
    /// Возвращает список CONNECTION_CODE с обработанными дубликатами.
    /// </summary>
    private List<string> ApplySuffixToDuplicates(List<string> inputList)
    {
        var result = new List<string>(inputList.Count);
        var countMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var usedCodes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // Первый проход: подсчёт количества каждого кода
        foreach (var code in inputList)
        {
            var trimmed = code?.Trim() ?? string.Empty;
            if (countMap.ContainsKey(trimmed))
                countMap[trimmed]++;
            else
                countMap[trimmed] = 1;
        }

        // Второй проход: обработка каждого кода с учётом уже использованных
        foreach (var code in inputList)
        {
            var original = code?.Trim() ?? string.Empty;
            if (string.IsNullOrEmpty(original))
            {
                result.Add(original);
                continue;
            }

            // Если код встречается только один раз, оставляем как есть
            if (countMap[original] == 1)
            {
                usedCodes.Add(original);
                result.Add(original);
                continue;
            }

            // Код повторяется, нужно определить суффикс
            string newCode = original;
            int suffix = 0;
            bool isFirstOccurrence = !usedCodes.Contains(original);
            
            if (isFirstOccurrence)
            {
                // Первое вхождение повторяющегося кода оставляем без суффикса
                newCode = original;
            }
            else
            {
                // Последующие вхождения: добавляем суффикс, начиная с 1
                suffix = 1;
                newCode = $"{original}_{suffix}";
                // Увеличиваем суффикс, пока не найдём свободный код
                while (usedCodes.Contains(newCode))
                {
                    suffix++;
                    newCode = $"{original}_{suffix}";
                }
            }

            usedCodes.Add(newCode);
            result.Add(newCode);
        }

        // Проверка, что длина результата равна длине входного списка
        if (result.Count != inputList.Count)
            throw new InvalidOperationException("Result length must equal input length.");

        return result;
    }
}