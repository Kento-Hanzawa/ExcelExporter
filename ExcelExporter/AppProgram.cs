using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Text.Unicode;
using System.Threading;
using System.Threading.Tasks;
using ConsoleAppFramework;
using CsvHelper;
using CsvHelper.Configuration;
using ExcelExporter.ConsoleAppCore;
using ExcelExporter.ConsoleAppCore.Filters;
using ExcelExporterCore;
using ZLogger;

namespace ExcelExporter
{
    [SupportedOSPlatform("windows")]
    internal sealed class CsvExporter : ConsoleApp
    {
        public CsvExporter(Microsoft.Extensions.Options.IOptions<AppSettings> config, Microsoft.Extensions.Logging.ILogger<CsvExporter> logger)
            : base(config, logger) { }

        [Command("sheet", "エクセルシートから変換するモード。")]
        public void FromSheet(
            [Option("i", "読み取り対象となるエクセルファイルのパス。")] string Input,
            [Option("o", "出力先ディレクトリのパス。")] string Output,
            [Option("n", "CSV フォーマットに変換するシート名。")] string RangeName,
            [Option("urx", "シート名に正規表現を使用するかどうか。使用する場合は true を指定します。")] bool UseRegex = false)
        {
            var output = new DirectoryInfo(Output);

            if (UseRegex)
            {
                var regex = new Regex(RangeName, RegexOptions.Singleline);

                using (var package = new ExcelPackage(new FileInfo(Input)))
                using (var tempDir = new TemporaryDirectoryScope())
                {
                    var results = package.ExportSheetAny(
                        info => new FileInfo(Path.Combine(tempDir.DirectoryName, $"{info.RangeName}.tmp")),
                        _ => 42,
                        info => regex.IsMatch(info.RangeName));
                    foreach (var result in results)
                    {
                        var (ParsedData, BadDatas) = UnicodeTextParser.Parse(result.DestFileInfo);
                        ExportCore(new FileInfo(Path.Combine(output.FullName, $"{RangeName}.csv")), result, ParsedData, BadDatas, Context);
                    }
                }
            }
            else
            {
                using (var package = new ExcelPackage(new FileInfo(Input)))
                using (var tempDir = new TemporaryDirectoryScope())
                {
                    var result = package.ExportSheet(
                        RangeName,
                        new FileInfo(Path.Combine(tempDir.DirectoryName, $"{RangeName}.tmp")),
                        42);
                    var (ParsedData, BadDatas) = UnicodeTextParser.Parse(result.DestFileInfo);
                    ExportCore(new FileInfo(Path.Combine(output.FullName, $"{RangeName}.csv")), result, ParsedData, BadDatas, Context);
                }
            }
        }

        [Command("table", "エクセルテーブルから変換するモード。")]
        public void FromTable(
            [Option("i", "読み取り対象となるエクセルファイルのパス。")] string Input,
            [Option("o", "出力先ディレクトリのパス。")] string Output,
            [Option("n", "CSV フォーマットに変換するテーブル名。")] string RangeName,
            [Option("urx", "テーブル名に正規表現を使用するかどうか。使用する場合は true を指定します。")] bool UseRegex = false)
        {
            var output = new DirectoryInfo(Output);

            if (UseRegex)
            {
                var regex = new Regex(RangeName, RegexOptions.Singleline);

                using (var package = new ExcelPackage(new FileInfo(Input)))
                using (var tempDir = new TemporaryDirectoryScope())
                {
                    var results = package.ExportTableAny(
                        info => new FileInfo(Path.Combine(tempDir.DirectoryName, $"{info.RangeName}.tmp")),
                        _ => 42,
                        info => regex.IsMatch(info.RangeName));
                    foreach (var result in results)
                    {
                        var (ParsedData, BadDatas) = UnicodeTextParser.Parse(result.DestFileInfo);
                        ExportCore(new FileInfo(Path.Combine(output.FullName, $"{RangeName}.csv")), result, ParsedData, BadDatas, Context);
                    }
                }
            }
            else
            {
                using (var package = new ExcelPackage(new FileInfo(Input)))
                using (var tempDir = new TemporaryDirectoryScope())
                {
                    var result = package.ExportTable(
                        RangeName,
                        new FileInfo(Path.Combine(tempDir.DirectoryName, $"{RangeName}.tmp")),
                        42);
                    var (ParsedData, BadDatas) = UnicodeTextParser.Parse(result.DestFileInfo);
                    ExportCore(new FileInfo(Path.Combine(output.FullName, $"{RangeName}.csv")), result, ParsedData, BadDatas, Context);
                }
            }
        }

        private static void ExportCore(FileInfo dest, ExportResult result, IReadOnlyList<dynamic> ParsedData, IReadOnlyList<BadDataFoundArgs> BadDatas, ConsoleAppContext context)
        {
            context.Logger.ZLogDebug("-----");

            Writer.WriteCsv(dest, ParsedData, new CsvConfiguration(CultureInfo.InvariantCulture) { Encoding = new UTF8Encoding(false) });
            context.Logger.ZLogInformation("Range: {0} ({1})", result.RangeInfo.RangeName, result.RangeInfo.RangeString);
            context.Logger.ZLogInformation("DestTo: {0}", dest.FullName);

            foreach (var v in BadDatas)
            {
                context.Logger.ZLogWarning("BadData: \nRawRecord: {0}\nField: {1}", v.RawRecord, v.Field);
            }

            context.Logger.ZLogDebug("-----");
        }
    }

    [SupportedOSPlatform("windows")]
    internal sealed class JsonExporter : ConsoleApp
    {
        public JsonExporter(Microsoft.Extensions.Options.IOptions<AppSettings> config, Microsoft.Extensions.Logging.ILogger<CsvExporter> logger)
            : base(config, logger) { }

        [Command("sheet", "エクセルシートから変換するモード。")]
        public void FromSheet(
            [Option("i", "読み取り対象となるエクセルファイルのパス。")] string Input,
            [Option("o", "出力先ディレクトリのパス。")] string Output,
            [Option("n", "CSV フォーマットに変換するシート名。")] string RangeName,
            [Option("urx", "シート名に正規表現を使用するかどうか。使用する場合は true を指定します。")] bool UseRegex = false)
        {
            var output = new DirectoryInfo(Output);

            if (UseRegex)
            {
                var regex = new Regex(RangeName, RegexOptions.Singleline);

                using (var package = new ExcelPackage(new FileInfo(Input)))
                using (var tempDir = new TemporaryDirectoryScope())
                {
                    var results = package.ExportSheetAny(
                        info => new FileInfo(Path.Combine(tempDir.DirectoryName, $"{info.RangeName}.tmp")),
                        _ => 42,
                        info => regex.IsMatch(info.RangeName));
                    foreach (var result in results)
                    {
                        var (ParsedData, BadDatas) = UnicodeTextParser.Parse(result.DestFileInfo);
                        ExportCore(new FileInfo(Path.Combine(output.FullName, $"{RangeName}.csv")), result, ParsedData, BadDatas, Context);
                    }
                }
            }
            else
            {
                using (var package = new ExcelPackage(new FileInfo(Input)))
                using (var tempDir = new TemporaryDirectoryScope())
                {
                    var result = package.ExportSheet(
                        RangeName,
                        new FileInfo(Path.Combine(tempDir.DirectoryName, $"{RangeName}.tmp")),
                        42);
                    var (ParsedData, BadDatas) = UnicodeTextParser.Parse(result.DestFileInfo);
                    ExportCore(new FileInfo(Path.Combine(output.FullName, $"{RangeName}.csv")), result, ParsedData, BadDatas, Context);
                }
            }
        }

        [Command("table", "エクセルテーブルから変換するモード。")]
        public void FromTable(
            [Option("i", "読み取り対象となるエクセルファイルのパス。")] string Input,
            [Option("o", "出力先ディレクトリのパス。")] string Output,
            [Option("n", "CSV フォーマットに変換するテーブル名。")] string RangeName,
            [Option("urx", "テーブル名に正規表現を使用するかどうか。使用する場合は true を指定します。")] bool UseRegex = false)
        {
            var output = new DirectoryInfo(Output);

            if (UseRegex)
            {
                var regex = new Regex(RangeName, RegexOptions.Singleline);

                using (var package = new ExcelPackage(new FileInfo(Input)))
                using (var tempDir = new TemporaryDirectoryScope())
                {
                    var results = package.ExportTableAny(
                        info => new FileInfo(Path.Combine(tempDir.DirectoryName, $"{info.RangeName}.tmp")),
                        _ => 42,
                        info => regex.IsMatch(info.RangeName));
                    foreach (var result in results)
                    {
                        var (ParsedData, BadDatas) = UnicodeTextParser.Parse(result.DestFileInfo);
                        ExportCore(new FileInfo(Path.Combine(output.FullName, $"{RangeName}.csv")), result, ParsedData, BadDatas, Context);
                    }
                }
            }
            else
            {
                using (var package = new ExcelPackage(new FileInfo(Input)))
                using (var tempDir = new TemporaryDirectoryScope())
                {
                    var result = package.ExportTable(
                        RangeName,
                        new FileInfo(Path.Combine(tempDir.DirectoryName, $"{RangeName}.tmp")),
                        42);
                    var (ParsedData, BadDatas) = UnicodeTextParser.Parse(result.DestFileInfo);
                    ExportCore(new FileInfo(Path.Combine(output.FullName, $"{RangeName}.csv")), result, ParsedData, BadDatas, Context);
                }
            }
        }

        private static void ExportCore(FileInfo dest, ExportResult result, IReadOnlyList<dynamic> ParsedData, IReadOnlyList<BadDataFoundArgs> BadDatas, ConsoleAppContext context)
        {
            context.Logger.ZLogDebug("-----");

            Writer.WriteJson(dest, ParsedData, new JsonWriterOptions() { Indented = true, Encoder = JavaScriptEncoder.Create(UnicodeRanges.All) });
            context.Logger.ZLogInformation("Range: {0} ({1})", result.RangeInfo.RangeName, result.RangeInfo.RangeString);
            context.Logger.ZLogInformation("DestTo: {0}", dest.FullName);

            foreach (var v in BadDatas)
            {
                context.Logger.ZLogWarning("BadData: \nRawRecord: {0}\nField: {1}", v.RawRecord, v.Field);
            }

            context.Logger.ZLogDebug("-----");
        }
    }
}
