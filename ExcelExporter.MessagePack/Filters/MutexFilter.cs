using System;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using ConsoleAppFramework;
using Cysharp.Text;
using ZLogger;

namespace ExcelExporter.MessagePack.Filters
{
    /// <summary>
    /// 複数のプロセスによる重複実行を避けるためのフィルター機能を提供します。
    /// </summary>
    // 参考: https://github.com/Cysharp/ConsoleAppFramework#filter
    internal sealed class MutexFilter : ConsoleAppFilter
    {
        public async override ValueTask Invoke(ConsoleAppContext context, Func<ConsoleAppContext, ValueTask> next)
        {
            var name = $"{Assembly.GetExecutingAssembly().GetName().Name}.{context.MethodInfo.DeclaringType?.Name ?? "(UnknownDeclaringType)"}.{context.MethodInfo.Name}";
            context.Logger.ZLogInformation("# Mutex Identifier: {0}", name);

            using var mutex = new Mutex(true, name, out var createdNew);
            if (!createdNew)
            {
                throw new MultipleExecutionException(name);
            }
            await next(context);
        }
    }

    internal sealed class MultipleExecutionException : Exception
    {
        private const string MessageFormat = "別のプロセスが既に実行されています。プロセス名=[{0}]";

        public string? MutexName { get; }

        public MultipleExecutionException()
            : this(null)
        {
        }

        public MultipleExecutionException(string? mutexName)
            : base(ZString.Format(MessageFormat, mutexName))
        {
            MutexName = mutexName;
        }

        public MultipleExecutionException(string? mutexName, Exception? innerException)
            : base(ZString.Format(MessageFormat, mutexName), innerException)
        {
            MutexName = mutexName;
        }

        public MultipleExecutionException(string? mutexName, string? message)
            : base(message)
        {
            MutexName = mutexName;
        }

        public MultipleExecutionException(string? mutexName, string? message, Exception? innerException)
            : base(message, innerException)
        {
            MutexName = mutexName;
        }
    }
}
