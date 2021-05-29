using System;
using System.Reflection;
using System.Threading.Tasks;
using ConsoleAppFramework;
using ZLogger;

namespace ExcelExporter.MessagePack.Filters
{
    /// <summary>
    /// アプリケーション実行前に <see cref="Assembly.GetExecutingAssembly"/> の情報をログに出力します。
    /// </summary>
    internal sealed class Log_AssemblyInfoFilter : ConsoleAppFilter
    {
        public override async ValueTask Invoke(ConsoleAppContext context, Func<ConsoleAppContext, ValueTask> next)
        {
            context.Logger.ZLogInformation("# Assembly Info: {0}", Assembly.GetExecutingAssembly().FullName);
            await next(context);
        }
    }
}
