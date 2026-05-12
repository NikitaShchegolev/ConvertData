using ConvertData.Application;
using OfficeOpenXml;

namespace ConvertData.Infrastructure
{
    /// <summary>
    /// Настройка лицензии EPPlus для работы с Excel.
    /// Для EPPlus 7.x достаточно установить `ExcelPackage.LicenseContext`.
    /// </summary>
    internal sealed class EpplusLicenseConfigurator : ILicenseConfigurator
    {
        /// <summary>
        /// Устанавливает контекст лицензии EPPlus как NonCommercial.
        /// </summary>
        public void Configure()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }
    }
}
