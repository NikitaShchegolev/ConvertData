namespace ConvertData.Application
{
    /// <summary>
    /// Контракт настройки лицензирования/инициализации сторонних библиотек.
    /// В данном проекте используется для EPPlus.
    /// </summary>
    internal interface ILicenseConfigurator
    {
        /// <summary>
        /// Выполняет настройку лицензии/контекста использования.
        /// </summary>
        void Configure();
    }
}
