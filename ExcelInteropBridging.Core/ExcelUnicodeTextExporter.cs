using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;

namespace ExcelInteropBridging.Core
{
    public sealed class ExcelUnicodeTextExporter : IDisposable
    {
        private readonly ExcelBridger bridger;



        public ExcelUnicodeTextExporter(FileInfo source)
        {
            bridger = new ExcelBridger(source);
        }

        public ExcelUnicodeTextExporter(FileInfo source, IEnumerable<FileInfo> references)
        {
            bridger = new ExcelBridger(source, references);
        }



        // ==========
        // Sheet
        // ==========
        private string ExportSheetCore(IComManaged<Worksheet> mgWorksheet, FileInfo dest)
        {
            // 新しく Excel Workbook を作成し、シートの使用領域をコピーします。
            // コピーした Workbook を UnicodeText (FileFormat: 42) で一時領域へ保存します。
            // この手順をおこなう理由は、worksheet をそのまま SaveAs した場合、
            // 読み取り対象の Workbook (managedWorkbook) を閉じるまで UnicodeText ファイルにアクセスできなくなるためです。
            using (var mgUsedRange = ComManaged.AsManaged(mgWorksheet.ComObject.UsedRange))
            using (var mgTempWorkbook = ComManaged.AsManaged(bridger.MgWorkbooks.ComObject.Add()))
            using (var mgTempWorksheet = ComManaged.AsManaged((Worksheet)mgTempWorkbook.ComObject.ActiveSheet))
            using (var mgDestinationRange = ComManaged.AsManaged(mgTempWorksheet.ComObject.Range["A1"]))
            {
                mgUsedRange.ComObject.Copy();
                mgDestinationRange.ComObject.PasteSpecial(XlPasteType.xlPasteValuesAndNumberFormats);
                dest.Directory.Create();
                mgTempWorkbook.ComObject.SaveAs(dest.FullName, FileFormat: 42);
                return mgUsedRange.GetRangeString();
            }
        }

        public ExcelExporterResult ExportSheet(int sheetIndex, FileInfo dest)
        {
            if (dest == null) throw new ArgumentNullException(nameof(dest));

            using (var mgWorksheet = bridger.GetWorksheet(sheetIndex))
            {
                if (mgWorksheet == null)
                {
                    throw new Exception($"ワークシート番号 {sheetIndex} は存在しません。");
                }
                var rangeString = ExportSheetCore(mgWorksheet, dest);
                return new ExcelExporterResult(mgWorksheet.ComObject.Name, rangeString, dest.FullName);
            }
        }

        public ExcelExporterResult ExportSheet(string sheetName, FileInfo dest)
        {
            if (sheetName == null) throw new ArgumentNullException(nameof(sheetName));
            if (dest == null) throw new ArgumentNullException(nameof(dest));

            using (var mgWorksheet = bridger.GetWorksheet(sheetName))
            {
                if (mgWorksheet == null)
                {
                    throw new Exception($"ワークシート名 {sheetName} は存在しません。");
                }
                var rangeString = ExportSheetCore(mgWorksheet, dest);
                return new ExcelExporterResult(sheetName, rangeString, dest.FullName);
            }
        }

        public ExcelExporterResult[] ExportSheetAny(Func<RangeInfo, FileInfo> destSelector, Func<RangeInfo, bool> where = null)
        {
            return ExportSheetAnyEnumerable(destSelector, where).ToArray();
        }

        public IEnumerable<ExcelExporterResult> ExportSheetAnyEnumerable(Func<RangeInfo, FileInfo> destSelector, Func<RangeInfo, bool> where = null)
        {
            if (destSelector == null) throw new ArgumentNullException(nameof(destSelector));

            foreach (var mgWorksheet in bridger.GetWorksheetAnyEnumerable())
            {
                var info = new RangeInfo(mgWorksheet);
                if (where?.Invoke(info) ?? true)
                {
                    var dest = destSelector(info);
                    var rangeString = ExportSheetCore(mgWorksheet, dest);
                    yield return new ExcelExporterResult(mgWorksheet.ComObject.Name, rangeString, dest.FullName);
                }
            }
        }



        // ==========
        // Table
        // ==========
        private string ExportTableCore(IComManaged<ListObject> mgListObject, FileInfo dest)
        {
            // 新しく Excel Workbook を作成し、テーブルの領域をコピーします。
            // コピーした Workbook を UnicodeText (FileFormat: 42) で一時領域へ保存します。
            // この手順をおこなう理由は、worksheet をそのまま SaveAs した場合、
            // 読み取り対象の Workbook (managedWorkbook) を閉じるまで UnicodeText ファイルにアクセスできなくなるためです。
            using (var mgTableRange = ComManaged.AsManaged(mgListObject.ComObject.Range))
            using (var mgTempWorkbook = ComManaged.AsManaged(bridger.MgWorkbooks.ComObject.Add()))
            using (var mgTempWorksheet = ComManaged.AsManaged((Worksheet)mgTempWorkbook.ComObject.ActiveSheet))
            using (var mgDestinationRange = ComManaged.AsManaged(mgTempWorksheet.ComObject.Range["A1"]))
            {
                mgTableRange.ComObject.Copy();
                mgDestinationRange.ComObject.PasteSpecial(XlPasteType.xlPasteValuesAndNumberFormats);
                dest.Directory.Create();
                mgTempWorkbook.ComObject.SaveAs(dest.FullName, FileFormat: 42);
                return mgTableRange.GetRangeString();
            }
        }

        public ExcelExporterResult ExportTable(string tableName, FileInfo dest)
        {
            if (tableName == null) throw new ArgumentNullException(nameof(tableName));
            if (dest == null) throw new ArgumentNullException(nameof(dest));

            using (var mgListObject = bridger.GetListObject(tableName))
            {
                if (mgListObject == null)
                {
                    throw new Exception($"テーブル名 {tableName} は存在しません。");
                }
                var rangeString = ExportTableCore(mgListObject, dest);
                return new ExcelExporterResult(tableName, rangeString, dest.FullName);
            }
        }

        public ExcelExporterResult[] ExportTableAny(Func<RangeInfo, FileInfo> destSelector, Func<RangeInfo, bool> where = null)
        {
            return ExportTableAnyEnumerable(destSelector, where).ToArray();
        }

        public IEnumerable<ExcelExporterResult> ExportTableAnyEnumerable(Func<RangeInfo, FileInfo> destSelector, Func<RangeInfo, bool> where = null)
        {
            if (destSelector == null) throw new ArgumentNullException(nameof(destSelector));

            foreach (var mgListObject in bridger.GetListObjectAnyEnumerable())
            {
                var info = new RangeInfo(mgListObject);
                if (where?.Invoke(info) ?? true)
                {
                    var dest = destSelector(info);
                    var rangeString = ExportTableCore(mgListObject, dest);
                    yield return new ExcelExporterResult(mgListObject.ComObject.Name, rangeString, dest.FullName);
                }
            }
        }



        #region IDisposable Support
        private bool disposed = false;

        private void Dispose(bool disposing)
        {
            if (disposed) return;

            if (disposing)
            {
                bridger.Dispose();
            }

            disposed = true;
        }

        ~ExcelUnicodeTextExporter()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        #endregion
    }
}
