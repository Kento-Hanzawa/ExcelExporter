using System.Linq;
using System.Globalization;
using CsvHelper.Configuration;

namespace ExcelInteropBridging
{
	internal static class TsvHelper
	{
		public static CsvConfiguration UnicodeTextConfiguration { get; } = new CsvConfiguration(CultureInfo.CurrentCulture)
		{
			Delimiter = "\t",
			TrimOptions = TrimOptions.Trim,
			ShouldSkipRecord = record => record.Record.All(string.IsNullOrWhiteSpace),
		};
	}
}
