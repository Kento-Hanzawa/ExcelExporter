using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using CsvHelper;
using CsvHelper.Configuration;

namespace ExcelInteropBridging
{
    public sealed class ExcelCsvExporter : ExcelExporter
    {
        private readonly CsvConfiguration option;

        public ExcelCsvExporter(FileInfo source, CsvConfiguration configuration)
            : base(source)
        {
            option = configuration;
        }

        public ExcelCsvExporter(FileInfo source, IEnumerable<FileInfo> references, CsvConfiguration configuration)
            : base(source, references)
        {
            option = configuration;
        }

        protected override void ConvertCore(IEnumerable<dynamic> parsedData, FileInfo dest)
        {
            using var stream = dest.OpenWrite();
            using var sWriter = new StreamWriter(stream, option.Encoding);
            using var csvWriter = new CsvWriter(sWriter, option);
            csvWriter.WriteRecords(parsedData);
        }
    }
}
