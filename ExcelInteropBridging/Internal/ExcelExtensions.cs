using Microsoft.Office.Interop.Excel;

namespace ExcelInteropBridging.Internal
{
    internal static class ExcelExtensions
    {
        /// <summary>
        /// 指定した <see cref="Worksheet"/> の使用範囲を、エクセル表現の範囲文字列に変換します。
        /// </summary>
        public static string GetRangeString(this IComManaged<Worksheet> sheet)
        {
            using (var mgRange = ComManaged.AsManaged(sheet.ComObject.UsedRange))
            {
                return GetRangeString(mgRange);
            }
        }

        /// <summary>
        /// 指定した <see cref="ListObject"/> の範囲を、エクセル表現の範囲文字列に変換します。
        /// </summary>
        public static string GetRangeString(this IComManaged<ListObject> table)
        {
            using (var mgRange = ComManaged.AsManaged(table.ComObject.Range))
            {
                return GetRangeString(mgRange);
            }
        }

        /// <summary>
        /// 指定した <see cref="Range"/> の範囲を、エクセル表現の範囲文字列に変換します。
        /// </summary>
        public static string GetRangeString(this IComManaged<Range> mgRange)
        {
            using (var mgRows = ComManaged.AsManaged(mgRange.ComObject.Rows))
            using (var mgColumns = ComManaged.AsManaged(mgRange.ComObject.Columns))
            {
                string beginAddress = $"{ToColumnName(mgRange.ComObject.Column)}{mgRange.ComObject.Row}";
                string endAddress = $"{ToColumnName(mgRange.ComObject.Column + mgColumns.ComObject.Count - 1)}{mgRange.ComObject.Row + mgRows.ComObject.Count - 1}";
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
