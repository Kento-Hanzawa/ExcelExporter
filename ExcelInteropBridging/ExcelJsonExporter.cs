using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace ExcelInteropBridging
{
    public sealed class ExcelJsonExporter : ExcelExporter
    {
        private readonly JsonWriterOptions options;

        public ExcelJsonExporter(FileInfo source, JsonWriterOptions jsonOptions = default)
            : base(source)
        {
            this.options = jsonOptions;
        }

        public ExcelJsonExporter(FileInfo source, IEnumerable<FileInfo> references, JsonWriterOptions jsonOptions = default)
            : base(source, references)
        {
            this.options = jsonOptions;
        }

        protected override void ConvertCore(IEnumerable<dynamic> parsedData, FileInfo dest)
        {
            using var stream = dest.OpenWrite();
            using var writer = new Utf8JsonWriter(stream, options);
            JsonSerializer.Serialize(writer, parsedData);
        }
    }
}
