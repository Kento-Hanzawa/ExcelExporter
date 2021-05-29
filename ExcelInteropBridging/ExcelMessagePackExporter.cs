using System.Collections.Generic;
using System.IO;
using MessagePack;

namespace ExcelInteropBridging
{
    public sealed class ExcelMessagePackExporter : ExcelExporter
    {
        private readonly MessagePackSerializerOptions options;

        public ExcelMessagePackExporter(FileInfo source, MessagePackSerializerOptions? options = null)
            : base(source)
        {
            this.options = options ?? MessagePackSerializerOptions.Standard;
        }

        public ExcelMessagePackExporter(FileInfo source, IEnumerable<FileInfo> references)
            : base(source, references)
        {
            this.options = options ?? MessagePackSerializerOptions.Standard;
        }

        protected override void ConvertCore(IEnumerable<dynamic> parsedData, FileInfo dest)
        {
            using var stream = dest.OpenWrite();
            MessagePackSerializer.Serialize(stream, parsedData, options);
        }
    }
}
