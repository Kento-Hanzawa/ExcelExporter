using System;
using System.Linq;
using System.Collections.Generic;
using System.IO;
using ExcelInteropBridging.Core;
using CsvHelper;
using System.Text;
using CsvHelper.Configuration;
using System.Globalization;

namespace ExcelInteropBridging
{
    public class Converter : IDisposable
    {
        private readonly ExcelUnicodeTextExporter unicodeTextExporter;

        public Converter(FileInfo excelFile)
            : this(excelFile, Array.Empty<FileInfo>())
        {
        }

        public Converter(FileInfo excelFile, IEnumerable<FileInfo> referenceExcelFiles)
        {
            unicodeTextExporter = new ExcelUnicodeTextExporter(excelFile, referenceExcelFiles);
        }

        void AAA()
        {
            using (var temp = new TemporaryDirectoryScope())
            {
                var file = new FileInfo(Path.Combine(temp.TemporaryDirectoryName, Path.GetRandomFileName()));
                var result = unicodeTextExporter.ExportSheet("", file);

                using (var stream = file.OpenRead())
                {

                }
            }
        }

        private static (IEnumerable<dynamic> parsedData, IEnumerable<BadDataFoundArgs> badDatas) ReadUnicodeText(in ExcelExporterResult result)
        {
            // UnicodeText ファイルを TSV フォーマットとして読み取り、IEnumerable<dynamic> として再取得します。
            // TODO: UnicodeText ファイルは通常 TSV フォーマットで保存されるようですが、Excel のバージョンによっては変更される可能性があります。
            var badDataList = new List<BadDataFoundArgs>();
            var configuration = UnicodeTextHelper.GetUnicodeTextConfiguration();
            configuration.BadDataFound = badData => badDataList.Add(badData);

            using var stream = result.ExportFile.OpenRead();
            using var streamReader = new StreamReader(stream, Encoding.Unicode);
            using var csvReader = new CsvReader(streamReader, configuration);
            return (csvReader.GetRecords<dynamic>().ToArray(), badDataList);
        }

        // sheet
        public ConverterResult ExportSheet(string sheetName, FileInfo dest)
        {
            using var tempDirScope = new TemporaryDirectoryScope();
            var result = unicodeTextExporter.ExportSheet(sheetName, new FileInfo(Path.Combine(tempDirScope.TemporaryDirectoryName, sheetName)));
            var (parsedData, badDatas) = ReadUnicodeText(result);

            // 書き込み処理
            using var fileStream = dest.OpenWrite();
            using var streamWriter = new StreamWriter(fileStream, new UTF8Encoding(false));
            var configuration = new CsvConfiguration(CultureInfo.CurrentCulture)
            {
                //TrimOptions = TrimOptions.Trim,
                Encoding = new UTF8Encoding(false)
            };
            using var csvWriter = new CsvWriter(streamWriter, configuration);
            csvWriter.WriteRecords(parsedData);

            return new ConverterResult(true, result.RangeName, result.RangeString, parsedData, badDatas);
        }

        public IEnumerable<ConverterResult> ExportAllSheet(DirectoryInfo destDir)
        {
            using var tempDirScope = new TemporaryDirectoryScope();
            foreach (var result in unicodeTextExporter.ExportAllSheet(new DirectoryInfo(tempDirScope.TemporaryDirectoryName)))
            {
                var (parsedData, badDatas) = ReadUnicodeText(result);

                // 書き込み処理
                using var fileStream = new FileInfo(Path.Combine(destDir.FullName, result.RangeName)).OpenWrite();
                using var streamWriter = new StreamWriter(fileStream, new UTF8Encoding(false));
                var configuration = new CsvConfiguration(CultureInfo.CurrentCulture)
                {
                    TrimOptions = TrimOptions.Trim,
                    Encoding = new UTF8Encoding(false)
                };
                using var csvWriter = new CsvWriter(streamWriter, configuration);
                csvWriter.WriteRecords(parsedData);

                yield return new ConverterResult(true, result.RangeName, result.RangeString, parsedData, badDatas);
            }
        }

        public void Dispose()
        {
            unicodeTextExporter?.Dispose();
        }

        private struct A
        {
            public ExcelExporterResult Result;
            public IEnumerable<dynamic> UnicodeTextParsedData;
        }

        public sealed class ConverterResult
        {
            public bool Success { get; } = false;
            public string RangeName { get; } = string.Empty;
            public string RangeString { get; } = string.Empty;
            public IReadOnlyList<IDictionary<string, string>> ParsedData { get; }
            public IReadOnlyList<BadDataFoundArgs> BadDataList { get; }

            public ConverterResult(bool success, string rangeName, string rangeString, IEnumerable<dynamic>? parsedData, IEnumerable<BadDataFoundArgs>? badDataList)
            {
                Success = success;
                RangeName = rangeName;
                RangeString = rangeString;
                ParsedData = new List<IDictionary<string, string>>(parsedData?.Cast<IDictionary<string, object>>().Select(x => x.ToDictionary(pair => pair.Key.Trim(), pair => pair.Value?.ToString()?.Trim() ?? string.Empty)) ?? Array.Empty<Dictionary<string, string>>());
                BadDataList = new List<BadDataFoundArgs>(badDataList ?? Array.Empty<BadDataFoundArgs>());
            }
        }
    }

    public sealed class CsvExporter : IDisposable
    {
        private readonly ExcelUnicodeTextExporter exporter;

        public CsvExporter(FileInfo excelFile)
            : this(excelFile, Array.Empty<FileInfo>())
        {
        }

        public CsvExporter(FileInfo excelFile, IEnumerable<FileInfo> referenceExcelFiles)
        {
            exporter = new ExcelUnicodeTextExporter(excelFile, referenceExcelFiles);
        }

        public void Export()
        {

        }

        public void Dispose()
        {
            exporter?.Dispose();
        }
    }
}
