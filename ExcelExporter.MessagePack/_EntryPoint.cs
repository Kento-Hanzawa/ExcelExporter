//
// ExcelExporter.MessagePack-1.0.6
//

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Unicode;
using ConsoleAppFramework;
using Cysharp.Text;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using ZLogger;

// NOTE: shift-jis などの文字コードを扱う場合、CodePagesEncodingProvider を事前に登録しておく必要があります。
Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

#pragma warning disable RS0030 // 禁止された API を使用しない
Encoding prev = Console.OutputEncoding;
#pragma warning restore RS0030 // 禁止された API を使用しない

try
{
#pragma warning disable RS0030 // 禁止された API を使用しない
    // NOTE: Console の既定文字コードは shift-jis のため、外国の文字や特殊文字を出力すると高確率で文字化けします。
    Console.OutputEncoding = Encoding.UTF8;
#pragma warning restore RS0030 // 禁止された API を使用しない

    await Host.CreateDefaultBuilder()
        .ConfigureServices(static (HostBuilderContext context, IServiceCollection service) =>
        {
            // NOTE: 設定ファイル（appsettings.json）が存在しない場合は、初期値となる Json を構築してファイルを作成します。
            var fileInfo = new FileInfo(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ExcelExporter.MessagePack.AppSettings.FileName));
            if (!fileInfo.Exists)
            {
                var settings = new Dictionary<string, object> { [context.HostingEnvironment.ApplicationName] = new ExcelExporter.MessagePack.AppSettings(), };
                var options = new JsonSerializerOptions { Encoder = JavaScriptEncoder.Create(UnicodeRanges.All), WriteIndented = true };
                File.WriteAllText(fileInfo.FullName, JsonSerializer.Serialize(settings, options));
            }

            var builder = new ConfigurationBuilder().SetBasePath(fileInfo.DirectoryName).AddJsonFile(fileInfo.Name);
            service.Configure<ExcelExporter.MessagePack.AppSettings>(builder.Build().GetSection(context.HostingEnvironment.ApplicationName));
        })
        .ConfigureLogging(static (HostBuilderContext context, ILoggingBuilder logging) =>
        {
            logging.ClearProviders();
            logging.SetMinimumLevel(LogLevel.Trace);
            logging.AddZLoggerConsole();

            var fileInfo = new FileInfo(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, context.HostingEnvironment.ApplicationName + ".log"));
            File.WriteAllText(fileInfo.FullName, string.Empty);
            logging.AddZLoggerFile(fileInfo.FullName, options =>
            {
                // "yyyy-MM-dd HH:mm:ss.fff±zzz  LogLevel  "
                var prefixFormat = ZString.PrepareUtf8<int, int, int, int, int, int, int, char, int, int, LogLevel>("{0:D4}-{1:D2}-{2:D2} {3:D2}:{4:D2}:{5:D2}.{6:D3}{7}{8:D2}:{9:D2}  {10,-11}  ");
                options.PrefixFormatter = (writer, info) =>
                {
                    var local = info.Timestamp.ToLocalTime();
                    prefixFormat.FormatTo(
                        ref writer,
                        local.Year, local.Month, local.Day,                        // 日付部
                        local.Hour, local.Minute, local.Second, local.Millisecond, // 時刻部
                        local.Offset.Ticks < 0 ? '-' : '+', Math.Abs(local.Offset.Hours), Math.Abs(local.Offset.Minutes), // オフセット部
                        info.LogLevel);
                };
            });
        })
#if DEBUG
        .RunConsoleAppFrameworkWebHostingAsync(ExcelExporter.MessagePack.AppMain.SwaggerUrl);
#else
        .RunConsoleAppFrameworkAsync<ExcelExporter.MessagePack.AppMain>(args);
#endif
}
finally
{
#pragma warning disable RS0030 // 禁止された API を使用しない
    Console.OutputEncoding = prev;
#pragma warning restore RS0030 // 禁止された API を使用しない
}
