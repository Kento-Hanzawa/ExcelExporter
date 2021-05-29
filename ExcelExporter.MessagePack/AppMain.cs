using System;
using ConsoleAppFramework;
using ZLogger;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using ExcelInteropBridging;
using System.IO;
using MessagePack;
using System.Text.RegularExpressions;

namespace ExcelExporter.MessagePack
{
    #region Filters
    // NOTE: 処理実行前後のイベントをフックするためのフィルターを設定します。Order パラメータが低いものが先に実行されます。
    // フィルターは Swagger 上で使用することが出来ないため、ビルド構成が Debug の場合は無効になります。
    // 詳しくは ConsoleAppFramework: Filter からも確認することが出来ます。（https://github.com/Cysharp/ConsoleAppFramework#filter）
#if !DEBUG
    [ConsoleAppFilter(typeof(Filters.Log_AssemblyInfoFilter), Order = 0)]
    //[ConsoleAppFilter(typeof(Filters.MutexFilter), Order = 1)]
    [ConsoleAppFilter(typeof(Filters.Log_RunningTimeFilter), Order = 2)]
#endif
    #endregion
    partial class AppMain
    {
        [Command("sheet", "エクセルシートから変換するモード。")]
        public void FromSheet(
            [Option("i", "読み取り対象となるエクセルファイルのパス。")] string excel,
            [Option("n", "MessagePack フォーマットに変換するシート名。")] string rangeName,
            [Option("o", "出力先ディレクトリのパス。")] string output,
            [Option("r", "シート名に正規表現を使用するかどうか。使用する場合は true を指定します。")] bool useRegex = false,
            [Option("ir", "正規表現の条件を反転するかどうか。反転させる場合は true を指定します。正規表現を使用する場合のみ使用されます。")] bool inverseRegex = false,
            [Option("z", "LZ4圧縮を使用するかどうか。使用する場合は true を指定します。")] bool useLZ4 = false)
        {
            using var exporter = new ExcelMessagePackExporter(new FileInfo(excel), MessagePackSerializerOptions.Standard.WithCompression(useLZ4 ? MessagePackCompression.Lz4BlockArray : MessagePackCompression.None));
            var outputDir = new DirectoryInfo(output);
            if (useRegex)
            {
                var regex = new Regex(rangeName, RegexOptions.Singleline);
                Logger.ZLogInformation("----- ----- ----- -----");
                foreach (var result in exporter.ExportFromSheetAnyEnumerable(
                    info => new FileInfo(Path.Combine(outputDir.FullName, info.RangeName + ".mpack")),
                    info => inverseRegex ? !regex.IsMatch(info.RangeName) : regex.IsMatch(info.RangeName)))
                {
                    Logger.ZLogInformation("出力成功: {0}", result.RangeName);
                    Logger.ZLogInformation("Result: {0}", result);
                    Logger.ZLogInformation("----- ----- ----- -----");
                }
            }
            else
            {
                var result = exporter.ExportFromSheet(rangeName, new FileInfo(Path.Combine(outputDir.FullName, rangeName + ".mpack")));
                Logger.ZLogInformation("----- ----- ----- -----");
                Logger.ZLogInformation("出力成功: {0}", rangeName);
                Logger.ZLogInformation("Result: {0}", result);
                Logger.ZLogInformation("----- ----- ----- -----");
            }
        }

        [Command("table", "エクセルテーブルから変換するモード。")]
        public void FromTable(
            [Option("i", "読み取り対象となるエクセルファイルのパス。")] string excel,
            [Option("n", "MessagePack フォーマットに変換するテーブル名。")] string rangeName,
            [Option("o", "出力先ディレクトリのパス。")] string output,
            [Option("r", "テーブル名に正規表現を使用するかどうか。使用する場合は true を指定します。")] bool useRegex = false,
            [Option("ir", "正規表現の条件を反転するかどうか。反転させる場合は true を指定します。正規表現を使用する場合のみ使用されます。")] bool inverseRegex = false,
            [Option("z", "LZ4圧縮を使用するかどうか。使用する場合は true を指定します。")] bool useLZ4 = false)
        {
            using var exporter = new ExcelMessagePackExporter(new FileInfo(excel), MessagePackSerializerOptions.Standard.WithCompression(useLZ4 ? MessagePackCompression.Lz4BlockArray : MessagePackCompression.None));
            var outdir = new DirectoryInfo(output);
            if (useRegex)
            {
                var regex = new Regex(rangeName, RegexOptions.Singleline);
                Logger.ZLogInformation("----- ----- ----- -----");
                foreach (var result in exporter.ExportFromTableAnyEnumerable(
                    info => new FileInfo(Path.Combine(outdir.FullName, info.RangeName + ".mpack")),
                    info => inverseRegex ? !regex.IsMatch(info.RangeName) : regex.IsMatch(info.RangeName)))
                {
                    Logger.ZLogInformation("出力成功: {0}", result.RangeName);
                    Logger.ZLogInformation("Result: {0}", result);
                    Logger.ZLogInformation("----- ----- ----- -----");
                }
            }
            else
            {
                var result = exporter.ExportFromTable(rangeName, new FileInfo(Path.Combine(outdir.FullName, rangeName + ".mpack")));
                Logger.ZLogInformation("----- ----- ----- -----");
                Logger.ZLogInformation("出力成功: {0}", rangeName);
                Logger.ZLogInformation("Result: {0}", result);
                Logger.ZLogInformation("----- ----- ----- -----");
            }
        }
    }
}

#region 既定実装
namespace ExcelExporter.MessagePack
{
    internal sealed partial class AppMain : ConsoleAppBase
    {
        // NOTE: ビルド構成が Debug の場合、Swagger によるデバッグが可能になります。
        // F5 からデバッグを開始したのち、この URL にアクセスすることで Swagger ページを開くことができます。
        // 詳しくは ConsoleAppFramework: Web Interface with Swagger からも確認することが出来ます。（https://github.com/Cysharp/ConsoleAppFramework#web-interface-with-swagger）
        public const string SwaggerUrl = "http://localhost:12345";

        internal IOptions<AppSettings> Config { get; }
        internal ILogger<AppMain> Logger { get; }

        public AppMain(IOptions<AppSettings> config, ILogger<AppMain> logger)
        {
            this.Config = config;
            this.Logger = logger;
        }
    }
}
#endregion
