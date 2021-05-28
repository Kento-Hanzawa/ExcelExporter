using Microsoft.Office.Interop.Excel;

namespace ExcelInteropBridging.Core
{
	internal static class ExcelExtensions
	{
		/// <summary>
		/// 指定した <see cref="Range"/> の範囲を、エクセル表現の範囲文字列に変換します。
		/// </summary>
		public static string GetRangeString(this Range range)
		{
			using (var cmRows = ComManaged.AsManaged(range.Rows))
			using (var cmColumns = ComManaged.AsManaged(range.Columns))
			{
				string beginAddress = $"{ToColumnName(range.Column)}{range.Row}";
				string endAddress = $"{ToColumnName(range.Column + cmColumns.ComObject.Count - 1)}{range.Row + cmRows.ComObject.Count - 1}";
				return $"{beginAddress}:{endAddress}";
			}

			string ToColumnName(int source)
			{
				if (source < 1)
				{
					return string.Empty;
				}
				return ToColumnName((source - 1) / 26) + char.ToString((char)('A' + ((source - 1) % 26)));
			}
		}
	}
}
