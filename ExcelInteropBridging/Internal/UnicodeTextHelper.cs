using System.Linq;
using System.Globalization;
using CsvHelper.Configuration;

namespace ExcelInteropBridging
{
	internal static class UnicodeTextHelper
	{
		public static CsvConfiguration GetUnicodeTextConfiguration()
		{
			return new CsvConfiguration(CultureInfo.CurrentCulture)
			{
				Delimiter = "\t",
				ShouldSkipRecord = record => record.Record.All(string.IsNullOrWhiteSpace),
			};
		}
	}
}
