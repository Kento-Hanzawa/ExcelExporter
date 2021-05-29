using System.IO;

namespace ExcelInteropBridging.Core
{
    public sealed class ExcelExporterResult
    {
        private readonly string destFile;

        public string RangeName { get; }
        public string RangeString { get; }
        public FileInfo DestFile { get { return new FileInfo(destFile); } }

        internal ExcelExporterResult(string rangeName, string rangeString, string destPath)
        {
            this.RangeName = rangeName;
            this.RangeString = rangeString;
            this.destFile = destPath;
        }
    }
}
