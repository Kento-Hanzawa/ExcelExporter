using System;
using System.Threading.Tasks;
using ConsoleAppFramework;
using ZLogger;

namespace ExcelExporter.MessagePack.Filters
{
    /// <summary>
    /// プロセスの実行を開始した時刻と、そのプロセスの完了までに要した時間をログに出力します。
    /// </summary>
    internal sealed class Log_RunningTimeFilter : ConsoleAppFilter
    {
        public override async ValueTask Invoke(ConsoleAppContext context, Func<ConsoleAppContext, ValueTask> next)
        {
            context.Logger.ZLogInformation("# 処理開始: {0}", context.Timestamp.ToLocalTime().ToString("yyyy/MM/dd HH:mm:ss"));
            try
            {
                await next(context);
                context.Logger.ZLogInformation("# 処理は正常に終了しました。経過時間: {0}", DateTimeOffset.UtcNow - context.Timestamp);
            }
            catch
            {
                context.Logger.ZLogInformation("# 処理は失敗で終了しました。経過時間: {0}", DateTimeOffset.UtcNow - context.Timestamp);
                throw;
            }
        }
    }
}
