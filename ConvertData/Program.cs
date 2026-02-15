using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using OfficeOpenXml;

namespace ConvertData
{
    /// <summary>
    /// Точка входа приложения.
    /// Служит композиционным корнем: создаёт и запускает сценарий конвертации.
    /// </summary>
    internal class Program
    {
        /// <summary>
        /// Запускает сценарий конвертации.
        /// По умолчанию читает входные файлы из папки `EXCEL` проекта и пишет результаты в `JSON_OUT`.
        /// Если переданы аргументы командной строки — обрабатывает только указанные файлы.
        /// </summary>
        /// <param name="args">Список путей к входным файлам (опционально).</param>
        static void Main(string[] args)
        {
            new Application.ConvertApp().Run(args);
            Console.ReadKey();
        }
    }
}
