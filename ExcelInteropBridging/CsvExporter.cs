using System;
using System.Collections.Generic;
using System.IO;
using ExcelInteropBridging.Core;

namespace ExcelInteropBridging
{
	public sealed class CsvExporter : IDisposable
	{
		private readonly UnicodeTextExporter exporter;

		public CsvExporter(FileInfo excelFile)
			: this(excelFile, Array.Empty<FileInfo>())
		{
		}

		public CsvExporter(FileInfo excelFile, IEnumerable<FileInfo> referenceExcelFiles)
		{
			exporter = new UnicodeTextExporter(excelFile, referenceExcelFiles);
		}

		public void Export()
		{

		}

		public void Dispose()
		{
			exporter?.Dispose();
		}
	}
}
