using System.IO;

namespace ExcelInteropBridging.Core
{
	public readonly struct ExcelExporterResult
	{
		private readonly string rangeName;
		private readonly string rangeString;
		private readonly string exportPath;

		public string RangeName => rangeName;
		public string RangeString => rangeString;
		public FileInfo ExportFile => new FileInfo(exportPath);

		internal ExcelExporterResult(string rangeName, string rangeString, string exportPath)
		{
			this.rangeName = rangeName;
			this.rangeString = rangeString;
			this.exportPath = exportPath;
		}
	}
}
