using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CsvHelper;
using CsvHelper.Configuration;

namespace ExcelExporter
{
    public static class UnicodeTextParser
    {
        public static (IReadOnlyList<dynamic> ParsedData, IReadOnlyList<BadDataFoundArgs> BadDatas) Parse(FileInfo file)
        {
            // UnicodeText ファイルを TSV フォーマットとして読み取り、IEnumerable<dynamic> として再取得します。
            // TODO: UnicodeText ファイルは通常 TSV フォーマットで保存されるようですが、Excel のバージョンによっては変更される可能性があります。
            var badDataList = new List<BadDataFoundArgs>();
            var configuration = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                Delimiter = "\t",
                ShouldSkipRecord = record => record.Record.All(string.IsNullOrWhiteSpace),
                Encoding = Encoding.Unicode,
            };
            configuration.BadDataFound = badData => badDataList.Add(badData);

            using (var stream = file.OpenRead())
            using (var sReader = new StreamReader(stream, Encoding.Unicode))
            using (var cReader = new CsvReader(sReader, configuration))
            {
                var records = cReader.GetRecords<dynamic>().ToList();
                return (records, badDataList);
            }
        }
    }
}
