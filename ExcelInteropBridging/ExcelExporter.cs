using System;
using System.Linq;
using System.Collections.Generic;
using System.IO;
using ExcelInteropBridging.Core;
using CsvHelper;
using System.Text;

namespace ExcelInteropBridging
{
    public abstract class ExcelExporter : IDisposable
    {
        private readonly ExcelUnicodeTextExporter unicodeTextExporter;



        public ExcelExporter(FileInfo source)
        {
            unicodeTextExporter = new ExcelUnicodeTextExporter(source);
        }

        public ExcelExporter(FileInfo source, IEnumerable<FileInfo> references)
        {
            unicodeTextExporter = new ExcelUnicodeTextExporter(source, references);
        }



        protected abstract void ConvertCore(IEnumerable<dynamic> parsedData, FileInfo dest);



        private static (IEnumerable<dynamic> parsedData, IEnumerable<BadDataFoundArgs> badDatas) ParseUnicodeText(FileInfo source)
        {
            // UnicodeText ファイルを TSV フォーマットとして読み取り、IEnumerable<dynamic> として再取得します。
            // TODO: UnicodeText ファイルは通常 TSV フォーマットで保存されるようですが、Excel のバージョンによっては変更される可能性があります。
            var badDataList = new List<BadDataFoundArgs>();
            var configuration = UnicodeTextHelper.GetUnicodeTextConfiguration();
            configuration.BadDataFound = badData => badDataList.Add(badData);

            using var stream = source.OpenRead();
            using var sReader = new StreamReader(stream, Encoding.Unicode);
            using var cReader = new CsvReader(sReader, configuration);
            return (cReader.GetRecords<dynamic>().ToArray(), badDataList);
        }



        // =====
        // Sheet
        // =====
        public ExcelExporterResult ExportFromSheet(string sheetName, FileInfo dest)
        {
            using var tempScope = new TemporaryDirectoryScope();

            var tempTsvFile = new FileInfo(Path.Combine(tempScope.TemporaryDirectoryName, Path.GetRandomFileName()));
            var result = unicodeTextExporter.ExportSheet(sheetName, tempTsvFile);

            var (parsedData, badDatas) = ParseUnicodeText(tempTsvFile);

            var tempDestFile = new FileInfo(Path.Combine(tempScope.TemporaryDirectoryName, Path.GetRandomFileName()));
            ConvertCore(parsedData, tempDestFile);
            tempDestFile.CopyTo(dest.FullName, true);

            return new ExcelExporterResult(dest.FullName, result.RangeName, result.RangeString, parsedData, badDatas);
        }

        public ExcelExporterResult[] ExportFromSheetAny(Func<RangeInfo, FileInfo> destSelector, Func<RangeInfo, bool>? where = null)
        {
            return ExportFromSheetAnyEnumerable(destSelector, where).ToArray();
        }

        public IEnumerable<ExcelExporterResult> ExportFromSheetAnyEnumerable(Func<RangeInfo, FileInfo> destSelector, Func<RangeInfo, bool>? where = null)
        {
            using var tempScope = new TemporaryDirectoryScope();

            FileInfo? dest = default;
            FileInfo? tempTsvFile = default;
            FileInfo DestSelector(RangeInfo info)
            {
                dest = destSelector(info);
                tempTsvFile = new FileInfo(Path.Combine(tempScope.TemporaryDirectoryName, Path.GetRandomFileName()));
                return tempTsvFile;
            }

            foreach (var result in unicodeTextExporter.ExportSheetAnyEnumerable(DestSelector, where))
            {
                var (parsedData, badDatas) = ParseUnicodeText(tempTsvFile);

                var tempDestFile = new FileInfo(Path.Combine(tempScope.TemporaryDirectoryName, Path.GetRandomFileName()));
                ConvertCore(parsedData, tempDestFile);
                tempDestFile.CopyTo(dest.FullName, true);

                yield return new ExcelExporterResult(dest.FullName, result.RangeName, result.RangeString, parsedData, badDatas);
            }
        }

        // =====
        // Table
        // =====
        public ExcelExporterResult ExportFromTable(string tableName, FileInfo dest)
        {
            using var tempScope = new TemporaryDirectoryScope();

            var tempTsvFile = new FileInfo(Path.Combine(tempScope.TemporaryDirectoryName, Path.GetRandomFileName()));
            var result = unicodeTextExporter.ExportTable(tableName, tempTsvFile);

            var (parsedData, badDatas) = ParseUnicodeText(tempTsvFile);

            var tempDestFile = new FileInfo(Path.Combine(tempScope.TemporaryDirectoryName, Path.GetRandomFileName()));
            ConvertCore(parsedData, tempDestFile);
            tempDestFile.CopyTo(dest.FullName, true);

            return new ExcelExporterResult(dest.FullName, result.RangeName, result.RangeString, parsedData, badDatas);
        }

        public ExcelExporterResult[] ExportFromTableAny(Func<RangeInfo, FileInfo> destSelector, Func<RangeInfo, bool>? where = null)
        {
            return ExportFromTableAnyEnumerable(destSelector, where).ToArray();
        }

        public IEnumerable<ExcelExporterResult> ExportFromTableAnyEnumerable(Func<RangeInfo, FileInfo> destSelector, Func<RangeInfo, bool>? where = null)
        {
            using var tempScope = new TemporaryDirectoryScope();

            FileInfo? dest = default;
            FileInfo? tempTsvFile = default;
            FileInfo DestSelector(RangeInfo info)
            {
                dest = destSelector(info);
                tempTsvFile = new FileInfo(Path.Combine(tempScope.TemporaryDirectoryName, Path.GetRandomFileName()));
                return tempTsvFile;
            }

            foreach (var result in unicodeTextExporter.ExportSheetAnyEnumerable(DestSelector, where))
            {
                var (parsedData, badDatas) = ParseUnicodeText(tempTsvFile);

                var tempDestFile = new FileInfo(Path.Combine(tempScope.TemporaryDirectoryName, Path.GetRandomFileName()));
                ConvertCore(parsedData, tempDestFile);
                tempDestFile.CopyTo(dest.FullName, true);

                yield return new ExcelExporterResult(dest.FullName, result.RangeName, result.RangeString, parsedData, badDatas);
            }
        }



        public void Dispose()
        {
            ((IDisposable)unicodeTextExporter).Dispose();
        }
    }
}
