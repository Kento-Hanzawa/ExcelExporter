using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;

namespace ExcelInteropBridging.Core
{
	public sealed class UnicodeTextExporter : ExcelBridge
	{
		private const string DefaultFileExtension = ".tsv";

		public UnicodeTextExporter(FileInfo excelFile)
			: base(excelFile, Array.Empty<FileInfo>())
		{
		}

		public UnicodeTextExporter(FileInfo excelFile, IEnumerable<FileInfo> referenceExcelFiles)
			: base(excelFile, referenceExcelFiles)
		{
		}

		// ==========
		// Sheet
		// ==========
		private bool InternalTryExportSheet(in string sheetName, in FileInfo outputFile, out string rangeString)
		{
			using (IComManaged<Sheets> managedSheets = ComManaged.Manage(managedWorkbook.ComObject.Worksheets))
			using (IComManaged<Worksheet> managedWorksheet = ComManaged.Manage((Worksheet)managedSheets.ComObject[sheetName]))
			{
				return InternalTryExportSheet(managedWorksheet.ComObject, outputFile, out rangeString);
			}
		}

		private bool InternalTryExportSheet(in Worksheet worksheet, in FileInfo outputFile, out string rangeString)
		{
			if (worksheet == null)
			{
				rangeString = string.Empty;
				return false;
			}

			// 新しく Excel Workbook を作成し、シートの使用領域をコピーします。
			// コピーした Workbook を UnicodeText (FileFormat: 42) で一時領域へ保存します。
			// この手順をおこなう理由は、worksheet をそのまま SaveAs した場合、
			// 読み取り対象の Workbook (managedWorkbook) を閉じるまで UnicodeText ファイルにアクセスできなくなるためです。
			using (IComManaged<Range> managedUsedRange = ComManaged.Manage(worksheet.UsedRange))
			using (IComManaged<Workbook> managedTempWorkbook = ComManaged.Manage(managedWorkbooks.ComObject.Add()))
			using (IComManaged<Worksheet> managedTempWorksheet = ComManaged.Manage((Worksheet)managedTempWorkbook.ComObject.ActiveSheet))
			using (IComManaged<Range> managedDestination = ComManaged.Manage(managedTempWorksheet.ComObject.Range["A1"]))
			{
				managedUsedRange.ComObject.Copy();
				managedDestination.ComObject.PasteSpecial(XlPasteType.xlPasteValuesAndNumberFormats);
				outputFile.Directory.Create();
				managedTempWorkbook.ComObject.SaveAs(outputFile.FullName, FileFormat: 42);
				rangeString = managedUsedRange.ComObject.GetRangeString();
				return true;
			}
		}

		public ExportResult ExportSheet(string sheetName, FileInfo outputFile)
		{
			if (sheetName == null) throw new ArgumentNullException(nameof(sheetName));
			if (outputFile == null) throw new ArgumentNullException(nameof(outputFile));

			if (InternalTryExportSheet(sheetName, outputFile, out string rangeString))
			{
				return new ExportResult(true, sheetName, rangeString, outputFile);
			}
			else
			{
				return new ExportResult(false, sheetName, rangeString, outputFile);
			}
		}

		public IEnumerable<ExportResult> ExportSheet(string sheetNamePattern, DirectoryInfo outputDirectory, bool useAntiPattern = false)
		{
			if (sheetNamePattern == null) throw new ArgumentNullException(nameof(sheetNamePattern));
			if (outputDirectory == null) throw new ArgumentNullException(nameof(outputDirectory));

			var sheetNameRegex = new Regex(sheetNamePattern, RegexOptions.Singleline);
			foreach (IComManaged<Worksheet> managedWorksheet in managedWorkbook.ComObject.GetWorksheets(worksheet => useAntiPattern ? !sheetNameRegex.IsMatch(worksheet.Name) : sheetNameRegex.IsMatch(worksheet.Name)))
			{
				string sheetName = managedWorksheet.ComObject.Name;
				var outputFile = new FileInfo(Path.Combine(outputDirectory.FullName, sheetName + DefaultFileExtension));
				if (InternalTryExportSheet(managedWorksheet.ComObject, outputFile, out string rangeString))
				{
					yield return new ExportResult(true, sheetName, rangeString, outputFile);
				}
				else
				{
					yield return new ExportResult(false, sheetName, rangeString, outputFile);
				}
			}
		}

		public IEnumerable<ExportResult> ExportSheet(IEnumerable<string> sheetNames, DirectoryInfo outputDirectory)
		{
			if (sheetNames == null) throw new ArgumentNullException(nameof(sheetNames));
			if (outputDirectory == null) throw new ArgumentNullException(nameof(outputDirectory));

			foreach (string sheetName in sheetNames)
			{
				var outputFile = new FileInfo(Path.Combine(outputDirectory.FullName, sheetName + DefaultFileExtension));
				if (InternalTryExportSheet(sheetName, outputFile, out string rangeString))
				{
					yield return new ExportResult(true, sheetName, rangeString, outputFile);
				}
				else
				{
					yield return new ExportResult(false, sheetName, rangeString, outputFile);
				}
			}
		}

		// ==========
		// Table
		// ==========
		private bool InternalTryExportTable(in string tableName, in FileInfo outputFile, out string rangeString)
		{
			using (IComManaged<Sheets> managedSheets = ComManaged.Manage(managedWorkbook.ComObject.Worksheets))
			{
				for (var sheetIndex = 1; sheetIndex <= managedSheets.ComObject.Count; sheetIndex++)
				{
					using (IComManaged<Worksheet> managedWorksheet = ComManaged.Manage((Worksheet)managedSheets.ComObject[sheetIndex]))
					using (IComManaged<ListObjects> managedListObjects = ComManaged.Manage(managedWorksheet.ComObject.ListObjects))
					{
						for (var listObjIndex = 1; listObjIndex <= managedListObjects.ComObject.Count; listObjIndex++)
						{
							using (IComManaged<ListObject> managedListObject = ComManaged.Manage(managedListObjects.ComObject[listObjIndex]))
							{
								if (managedListObject.ComObject.Name == tableName)
								{
									return InternalTryExportTable(managedListObject.ComObject, outputFile, out rangeString);
								}
							}
						}
					}
				}
			}

			// 指定された名前のテーブルが見つからなかった。
			rangeString = string.Empty;
			return false;
		}

		private bool InternalTryExportTable(in ListObject listObject, in FileInfo outputFile, out string rangeString)
		{
			if (listObject == null)
			{
				rangeString = string.Empty;
				return false;
			}

			// 新しく Excel Workbook を作成し、テーブルの領域をコピーします。
			// コピーした Workbook を UnicodeText (FileFormat: 42) で一時領域へ保存します。
			// この手順をおこなう理由は、worksheet をそのまま SaveAs した場合、
			// 読み取り対象の Workbook (managedWorkbook) を閉じるまで UnicodeText ファイルにアクセスできなくなるためです。
			using (IComManaged<Range> managedTableRange = ComManaged.Manage(listObject.Range))
			using (IComManaged<Workbook> managedTempWorkbook = ComManaged.Manage(managedWorkbooks.ComObject.Add()))
			using (IComManaged<Worksheet> managedTempWorksheet = ComManaged.Manage((Worksheet)managedTempWorkbook.ComObject.ActiveSheet))
			using (IComManaged<Range> managedDestination = ComManaged.Manage(managedTempWorksheet.ComObject.Range["A1"]))
			{
				managedTableRange.ComObject.Copy();
				managedDestination.ComObject.PasteSpecial(XlPasteType.xlPasteValuesAndNumberFormats);
				outputFile.Directory.Create();
				managedTempWorkbook.ComObject.SaveAs(outputFile.FullName, FileFormat: 42);
				rangeString = managedTableRange.ComObject.GetRangeString();
				return true;
			}
		}

		public ExportResult ExportTable(string tableName, FileInfo outputFile)
		{
			if (tableName == null) throw new ArgumentNullException(nameof(tableName));
			if (outputFile == null) throw new ArgumentNullException(nameof(outputFile));

			if (InternalTryExportTable(tableName, outputFile, out string rangeString))
			{
				return new ExportResult(true, tableName, rangeString, outputFile);
			}
			else
			{
				return new ExportResult(false, tableName, rangeString, outputFile);
			}
		}

		public IEnumerable<ExportResult> ExportTable(string tableNamePattern, DirectoryInfo outputDirectory, bool useAntiPattern = false)
		{
			if (tableNamePattern == null) throw new ArgumentNullException(nameof(tableNamePattern));
			if (outputDirectory == null) throw new ArgumentNullException(nameof(outputDirectory));

			var tableNameRegex = new Regex(tableNamePattern, RegexOptions.Singleline);
			foreach (IComManaged<ListObject> managedListObject in managedWorkbook.ComObject.GetListObjects(listObject => useAntiPattern ? !tableNameRegex.IsMatch(listObject.Name) : tableNameRegex.IsMatch(listObject.Name)))
			{
				string tableName = managedListObject.ComObject.Name;
				var outputFile = new FileInfo(Path.Combine(outputDirectory.FullName, tableName + DefaultFileExtension));
				if (InternalTryExportTable(managedListObject.ComObject, outputFile, out string rangeString))
				{
					yield return new ExportResult(true, tableName, rangeString, outputFile);
				}
				else
				{
					yield return new ExportResult(false, tableName, rangeString, outputFile);
				}
			}
		}

		public IEnumerable<ExportResult> ExportTable(IEnumerable<string> tableNames, DirectoryInfo outputDirectory)
		{
			if (tableNames == null) throw new ArgumentNullException(nameof(tableNames));
			if (outputDirectory == null) throw new ArgumentNullException(nameof(outputDirectory));

			foreach (string tableName in tableNames)
			{
				var outputFile = new FileInfo(Path.Combine(outputDirectory.FullName, tableName + DefaultFileExtension));
				if (InternalTryExportTable(tableName, outputFile, out string rangeString))
				{
					yield return new ExportResult(true, tableName, rangeString, outputFile);
				}
				else
				{
					yield return new ExportResult(false, tableName, rangeString, outputFile);
				}
			}
		}

		public sealed class ExportResult
		{
			public bool Success { get; } = false;
			public string RangeName { get; } = string.Empty;
			public string RangeString { get; } = string.Empty;
			public FileInfo OutputFile { get; }

			public ExportResult(bool success, string rangeName, string rangeString, FileInfo outputFile)
			{
				Success = success;
				RangeName = rangeName;
				RangeString = rangeString;
				OutputFile = outputFile;
			}
		}
	}
}
