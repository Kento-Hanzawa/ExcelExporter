using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;

namespace InteropBridging.Internal
{
	internal static class ExcelExtensions
	{
		/// <summary>
		/// 指定したシート番号の <see cref="Worksheet"/> を取得します。シートが存在しない場合は <see langword="null"/> が返されます。
		/// </summary>
		[Obsolete("COM リソースの参照エラーが発生する場合があります。")]
		public static Worksheet GetWorksheet(this Workbook workbook, int sheetIndex)
		{
			using (IComManaged<Sheets> mdWorksheets = ComManaged.Manage(workbook.Worksheets))
			{
				return (Worksheet)mdWorksheets.ComObject[sheetIndex];
			}
		}

		/// <summary>
		/// 指定したシート名に一致する <see cref="Worksheet"/> を取得します。シートが存在しない場合は <see langword="null"/> が返されます。
		/// </summary>
		[Obsolete("COM リソースの参照エラーが発生する場合があります。")]
		public static Worksheet GetWorksheet(this Workbook workbook, string sheetName)
		{
			using (IComManaged<Sheets> mdWorksheets = ComManaged.Manage(workbook.Worksheets))
			{
				return (Worksheet)mdWorksheets.ComObject[sheetName];
			}
		}

		/// <summary>
		/// 指定した条件を満たす全ての <see cref="Worksheet"/> を取得します。（遅延実行専用）
		/// </summary>
		public static IEnumerable<IComManaged<Worksheet>> GetWorksheets(this Workbook workbook, Predicate<Worksheet> predicate)
		{
			using (IComManaged<Sheets> mdWorksheets = ComManaged.Manage(workbook.Worksheets))
			{
				for (var i = 1; i <= mdWorksheets.ComObject.Count; i++)
				{
					using (IComManaged<Worksheet> mdWorksheet = ComManaged.Manage((Worksheet)mdWorksheets.ComObject[i]))
					{
						if (predicate(mdWorksheet.ComObject))
						{
							yield return mdWorksheet;
						}
					}
				}
			}
		}

		/// <summary>
		/// 指定したテーブル名に一致する <see cref="ListObject"/> を取得します。テーブルが存在しない場合は <see langword="null"/> が返されます。
		/// </summary>
		[Obsolete("COM リソースの参照エラーが発生する場合があります。")]
		public static ListObject GetListObject(this Workbook workbook, string tableName)
		{
			using (IComManaged<Sheets> mdWorksheets = ComManaged.Manage(workbook.Worksheets))
			{
				for (var sheetIndex = 1; sheetIndex <= mdWorksheets.ComObject.Count; sheetIndex++)
				{
					using (IComManaged<Worksheet> mdWorksheet = ComManaged.Manage((Worksheet)mdWorksheets.ComObject[sheetIndex]))
					using (IComManaged<ListObjects> mdListObjects = ComManaged.Manage(mdWorksheet.ComObject.ListObjects))
					{
						for (var listObjIndex = 1; listObjIndex <= mdListObjects.ComObject.Count; listObjIndex++)
						{
							using (IComManaged<ListObject> listObject = ComManaged.Manage(mdListObjects.ComObject[listObjIndex]))
							{
								if (listObject.ComObject.Name == tableName)
								{
									return mdListObjects.ComObject[listObjIndex];
								}
							}
						}
					}
				}
			}
			return null;
		}

		/// <summary>
		/// 指定した条件を満たす全ての <see cref="ListObject"/> を取得します。（遅延実行専用）
		/// </summary>
		public static IEnumerable<IComManaged<ListObject>> GetListObjects(this Workbook workbook, Predicate<ListObject> predicate)
		{
			using (IComManaged<Sheets> mdWorksheets = ComManaged.Manage(workbook.Worksheets))
			{
				for (var sheetIndex = 1; sheetIndex <= mdWorksheets.ComObject.Count; sheetIndex++)
				{
					using (IComManaged<Worksheet> mdWorksheet = ComManaged.Manage((Worksheet)mdWorksheets.ComObject[sheetIndex]))
					using (IComManaged<ListObjects> mdListObjects = ComManaged.Manage(mdWorksheet.ComObject.ListObjects))
					{
						for (var listObjIndex = 1; listObjIndex <= mdListObjects.ComObject.Count; listObjIndex++)
						{
							using (IComManaged<ListObject> mdListObject = ComManaged.Manage(mdListObjects.ComObject[listObjIndex]))
							{
								if (predicate(mdListObject.ComObject))
								{
									yield return mdListObject;
								}
							}
						}
					}
				}
			}
		}

		/// <summary>
		/// 指定した <see cref="Range"/> の範囲を、エクセル表現の範囲文字列に変換します。
		/// </summary>
		public static string GetRangeString(this Range range)
		{
			string ToColumnName(int source)
			{
				if (source < 1)
				{
					return string.Empty;
				}
				return ToColumnName((source - 1) / 26) + (char)('A' + ((source - 1) % 26));
			}

			using (IComManaged<Range> mdRows = ComManaged.Manage(range.Rows))
			using (IComManaged<Range> mdColumns = ComManaged.Manage(range.Columns))
			{
				string beginAddress = $"{ToColumnName(range.Column)}{range.Row}";
				string endAddress = $"{ToColumnName(range.Column + mdColumns.ComObject.Count - 1)}{range.Row + mdRows.ComObject.Count - 1}";
				return $"{beginAddress}:{endAddress}";
			}
		}
	}
}
