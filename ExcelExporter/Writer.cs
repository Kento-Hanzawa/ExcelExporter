using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Unicode;
using System.Threading.Tasks;
using CsvHelper;
using CsvHelper.Configuration;

namespace ExcelExporter
{
    public static class Writer
    {
        public static void WriteCsv(FileInfo dest, IReadOnlyList<dynamic> data, CsvConfiguration configuration)
        {
            using (var stream = dest.OpenWrite())
            using (var sWriter = new StreamWriter(stream, new UTF8Encoding(false)))
            using (var csvWriter = new CsvWriter(sWriter, new CsvConfiguration(CultureInfo.InvariantCulture) { Encoding = new UTF8Encoding(false) }))
            {
                csvWriter.WriteRecords(data);
            }
        }

        public static void WriteJson(FileInfo dest, IReadOnlyList<dynamic> data, JsonWriterOptions options)
        {
            using (var stream = dest.OpenWrite())
            using (var writer = new Utf8JsonWriter(stream, new JsonWriterOptions() { Indented = true, Encoder = JavaScriptEncoder.Create(UnicodeRanges.All) }))
            {
                JsonSerializer.Serialize(writer, data);
            }
        }
    }
}
