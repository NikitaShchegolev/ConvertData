using ConvertData.Domain;

namespace ConvertData.Infrastructure;

internal static partial class EpplusRowPropertyMaps
{
    private sealed class GeometrySectionDefinition
    {
        public required string FieldPrefix { get; init; }
        public required Action<Row, string> ProfileSetter { get; init; }
        public required string[] ProfileHeaders { get; init; }
        public required Action<Row, string> GostSetter { get; init; }
        public required string[] GostHeaders { get; init; }
        public required Action<Row, double> HSetter { get; init; }
        public required Action<Row, double> BSetter { get; init; }
        public required Action<Row, double> SSetter { get; init; }
        public required Action<Row, double> TSetter { get; init; }
        public required Action<Row, double> ASetter { get; init; }
        public required Action<Row, double> PSetter { get; init; }
        public required Action<Row, double> IzSetter { get; init; }
        public required Action<Row, double> IySetter { get; init; }
        public required Action<Row, double> IxSetter { get; init; }
        public required Action<Row, double> WzSetter { get; init; }
        public required Action<Row, double> WySetter { get; init; }
        public required Action<Row, double> WxSetter { get; init; }
        public required Action<Row, double> SzSetter { get; init; }
        public required Action<Row, double> SySetter { get; init; }
        public required Action<Row, double> izSetter { get; init; }
        public required Action<Row, double> iySetter { get; init; }
        public required Action<Row, double> xoSetter { get; init; }
        public required Action<Row, double> yoSetter { get; init; }
        public Dictionary<string, string[]> HeaderAliases { get; init; } = new(StringComparer.OrdinalIgnoreCase);

        public string[] GetHeaders(string fieldName)
        {
            return HeaderAliases.TryGetValue(fieldName, out var headers)
                ? headers
                : [$"{FieldPrefix}_{fieldName}"];
        }
    }
}
