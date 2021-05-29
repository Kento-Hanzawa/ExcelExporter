using System;
using System.Linq;
using System.Collections.Generic;
using CsvHelper;
using System.IO;

namespace ExcelInteropBridging
{
    public sealed record ExcelExporterResult
    {
        public string DestFilePath { get; init; }
        public string RangeName { get; init; }
        public string RangeString { get; init; }
        public IReadOnlyList<IDictionary<string, string>> ParsedData { get; init; }
        public IReadOnlyList<BadDataFoundArgs> BadDataList { get; init; }

        internal ExcelExporterResult(string destFilePath, string rangeName, string rangeString, IEnumerable<dynamic>? parsedData, IEnumerable<BadDataFoundArgs>? badDataList)
        {
            DestFilePath = destFilePath;
            RangeName = rangeName;
            RangeString = rangeString;
            ParsedData = new List<IDictionary<string, string>>(parsedData?.Cast<IDictionary<string, object>>().Select(x => x.ToDictionary(pair => pair.Key.Trim(), pair => pair.Value?.ToString()?.Trim() ?? string.Empty)) ?? Array.Empty<Dictionary<string, string>>());
            BadDataList = new List<BadDataFoundArgs>(badDataList ?? Array.Empty<BadDataFoundArgs>());
        }
    }
}
