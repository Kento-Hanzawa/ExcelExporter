using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;

namespace ExcelInteropBridging.Core
{
    public sealed class ExcelUnicodeTextExporter : ExcelBridger
    {
        private const string DefaultFileExtension = ".txt";

        public ExcelUnicodeTextExporter(FileInfo excelFile)
            : base(excelFile, Array.Empty<FileInfo>())
        {
        }

        public ExcelUnicodeTextExporter(FileInfo excelFile, IEnumerable<FileInfo> referenceExcelFiles)
            : base(excelFile, referenceExcelFiles)
        {
        }

        // ==========
        // Sheet
        // ==========
        private string ExportSheetCore(IComManaged<Worksheet> mWorksheet, FileInfo dest)
        {
            // 新しく Excel Workbook を作成し、シートの使用領域をコピーします。
            // コピーした Workbook を UnicodeText (FileFormat: 42) で一時領域へ保存します。
            // この手順をおこなう理由は、worksheet をそのまま SaveAs した場合、
            // 読み取り対象の Workbook (managedWorkbook) を閉じるまで UnicodeText ファイルにアクセスできなくなるためです。
            using (var cmUsedRange = ComManaged.AsManaged(mWorksheet.ComObject.UsedRange))
            using (var cmTempWorkbook = ComManaged.AsManaged(mWorkbooks.ComObject.Add()))
            using (var cmTempWorksheet = ComManaged.AsManaged((Worksheet)cmTempWorkbook.ComObject.ActiveSheet))
            using (var cmDestinationRange = ComManaged.AsManaged(cmTempWorksheet.ComObject.Range["A1"]))
            {
                cmUsedRange.ComObject.Copy();
                cmDestinationRange.ComObject.PasteSpecial(XlPasteType.xlPasteValuesAndNumberFormats);
                dest.Directory.Create();
                cmTempWorkbook.ComObject.SaveAs(dest.FullName, FileFormat: 42);
                return cmUsedRange.ComObject.GetRangeString();
            }
        }

        public ExcelExporterResult ExportSheet(string sheetName, FileInfo dest)
        {
            if (sheetName == null) throw new ArgumentNullException(nameof(sheetName));
            if (dest == null) throw new ArgumentNullException(nameof(dest));
            using (var cmWorksheet = GetWorksheet(sheetName))
            {
                if (cmWorksheet == null)
                {
                    throw new Exception($"ワークシート {sheetName} が存在しません。");
                }
                var rangeString = ExportSheetCore(cmWorksheet, dest);
                return new ExcelExporterResult(sheetName, rangeString, dest.FullName);
            }
        }

        public IEnumerable<ExcelExporterResult> ExportSheetAny(DirectoryInfo dest)
        {
            return ExportSheetAny(dest, null, null);
        }

        public IEnumerable<ExcelExporterResult> ExportSheetAny(DirectoryInfo dest, Func<SheetInfo, string> destNameSelector)
        {
            return ExportSheetAny(dest, null, destNameSelector);
        }

        public IEnumerable<ExcelExporterResult> ExportSheetAny(DirectoryInfo dest, Func<SheetInfo, bool> doExportSelector)
        {
            return ExportSheetAny(dest, doExportSelector, null);
        }

        public IEnumerable<ExcelExporterResult> ExportSheetAny(DirectoryInfo dest, Func<SheetInfo, bool> doExportSelector, Func<SheetInfo, string> destNameSelector)
        {
            if (dest == null) throw new ArgumentNullException(nameof(dest));
            foreach (var cmWorksheet in GetWorksheetAny())
            {
                var info = new SheetInfo();
                if (doExportSelector?.Invoke(info) ?? true)
                {
                    var file = new FileInfo(Path.Combine(dest.FullName, destNameSelector?.Invoke(info) ?? Path.GetRandomFileName()));
                    var rangeString = ExportSheetCore(cmWorksheet, file);
                    yield return new ExcelExporterResult(cmWorksheet.ComObject.Name, rangeString, dest.FullName);
                }
            }
        }

        public class SheetInfo
        {
            public string Name { get; }
        }

        public IEnumerable<ExcelExporterResult> ExportSheet(string sheetNamePattern, DirectoryInfo destDir, bool inversion = false)
        {
            if (sheetNamePattern == null) throw new ArgumentNullException(nameof(sheetNamePattern));
            if (destDir == null) throw new ArgumentNullException(nameof(destDir));
            var sheetNameRegex = new Regex(sheetNamePattern, RegexOptions.Singleline);
            foreach (var cmWorksheet in GetWorksheetAny(x => inversion ? !sheetNameRegex.IsMatch(x.ComObject.Name) : sheetNameRegex.IsMatch(x.ComObject.Name)))
            {
                using (cmWorksheet)
                {
                    var dest = new FileInfo(Path.Combine(destDir.FullName, cmWorksheet.ComObject.Name + DefaultFileExtension));
                    var rangeString = ExportSheetCore(cmWorksheet, dest);
                    yield return new ExcelExporterResult(cmWorksheet.ComObject.Name, rangeString, dest.FullName);
                }
            }
        }

        public IEnumerable<ExcelExporterResult> ExportSheet(IEnumerable<string> sheetNames, DirectoryInfo destDir)
        {
            if (sheetNames == null) throw new ArgumentNullException(nameof(sheetNames));
            if (destDir == null) throw new ArgumentNullException(nameof(destDir));
            foreach (var sheetName in sheetNames)
            {
                var dest = new FileInfo(Path.Combine(destDir.FullName, sheetName + DefaultFileExtension));
                yield return ExportSheet(sheetName, dest);
            }
        }

        public IEnumerable<ExcelExporterResult> ExportAllSheet(DirectoryInfo destDir)
        {
            if (destDir == null) throw new ArgumentNullException(nameof(destDir));
            foreach (var cmWorksheet in GetWorksheetAny())
            {
                using (cmWorksheet)
                {
                    var dest = new FileInfo(Path.Combine(destDir.FullName, cmWorksheet.ComObject.Name + DefaultFileExtension));
                    var rangeString = ExportSheetCore(cmWorksheet, dest);
                    yield return new ExcelExporterResult(cmWorksheet.ComObject.Name, rangeString, dest.FullName);
                }
            }
        }



        // ==========
        // Table
        // ==========
        private string ExportTableCore(IComManaged<ListObject> cmListObject, FileInfo dest)
        {
            // 新しく Excel Workbook を作成し、テーブルの領域をコピーします。
            // コピーした Workbook を UnicodeText (FileFormat: 42) で一時領域へ保存します。
            // この手順をおこなう理由は、worksheet をそのまま SaveAs した場合、
            // 読み取り対象の Workbook (managedWorkbook) を閉じるまで UnicodeText ファイルにアクセスできなくなるためです。
            using (var cmTableRange = ComManaged.AsManaged(cmListObject.ComObject.Range))
            using (var cmTempWorkbook = ComManaged.AsManaged(mWorkbooks.ComObject.Add()))
            using (var cmTempWorksheet = ComManaged.AsManaged((Worksheet)cmTempWorkbook.ComObject.ActiveSheet))
            using (var cmDestinationRange = ComManaged.AsManaged(cmTempWorksheet.ComObject.Range["A1"]))
            {
                cmTableRange.ComObject.Copy();
                cmDestinationRange.ComObject.PasteSpecial(XlPasteType.xlPasteValuesAndNumberFormats);
                dest.Directory.Create();
                cmTempWorkbook.ComObject.SaveAs(dest.FullName, FileFormat: 42);
                return cmTableRange.ComObject.GetRangeString();
            }
        }

        public ExcelExporterResult ExportTable(string tableName, FileInfo dest)
        {
            if (tableName == null) throw new ArgumentNullException(nameof(tableName));
            if (dest == null) throw new ArgumentNullException(nameof(dest));
            using (var cmSheets = ComManaged.AsManaged(mWorkbook.ComObject.Worksheets))
            {
                for (var sheetIndex = 1; sheetIndex <= cmSheets.ComObject.Count; sheetIndex++)
                {
                    using (var cmWorksheet = ComManaged.AsManaged((Worksheet)cmSheets.ComObject[sheetIndex]))
                    using (var cmListObjects = ComManaged.AsManaged(cmWorksheet.ComObject.ListObjects))
                    {
                        for (var listObjIndex = 1; listObjIndex <= cmListObjects.ComObject.Count; listObjIndex++)
                        {
                            using (var cmListObject = ComManaged.AsManaged(cmListObjects.ComObject[listObjIndex]))
                            {
                                if (cmListObject.ComObject.Name == tableName)
                                {
                                    string rangeString = ExportTableCore(cmListObject, dest);
                                    return new ExcelExporterResult(tableName, rangeString, dest.FullName);
                                }
                            }
                        }
                    }
                }
            }

            // 指定された名前のテーブルが見つからなかった。
            throw new Exception($"テーブル {tableName} が存在しません。");
        }

        public IEnumerable<ExcelExporterResult> ExportTable(string tableNamePattern, DirectoryInfo destDir, bool inversion = false)
        {
            if (tableNamePattern == null) throw new ArgumentNullException(nameof(tableNamePattern));
            if (destDir == null) throw new ArgumentNullException(nameof(destDir));
            var tableNameRegex = new Regex(tableNamePattern, RegexOptions.Singleline);
            foreach (var cmListObject in GetListObjectAny(x => inversion ? !tableNameRegex.IsMatch(x.ComObject.Name) : tableNameRegex.IsMatch(x.ComObject.Name)))
            {
                using (cmListObject)
                {
                    var dest = new FileInfo(Path.Combine(destDir.FullName, cmListObject.ComObject.Name + DefaultFileExtension));
                    string rangeString = ExportTableCore(cmListObject, dest);
                    yield return new ExcelExporterResult(cmListObject.ComObject.Name, rangeString, dest.FullName);
                }
            }
        }

        public IEnumerable<ExcelExporterResult> ExportTable(IEnumerable<string> tableNames, DirectoryInfo destDir)
        {
            if (tableNames == null) throw new ArgumentNullException(nameof(tableNames));
            if (destDir == null) throw new ArgumentNullException(nameof(destDir));
            foreach (string tableName in tableNames)
            {
                var dest = new FileInfo(Path.Combine(destDir.FullName, tableName + DefaultFileExtension));
                yield return ExportTable(tableName, dest);
            }
        }

        public IEnumerable<ExcelExporterResult> ExportAllTable(DirectoryInfo destDir)
        {
            if (destDir == null) throw new ArgumentNullException(nameof(destDir));
            foreach (var cmListObject in GetListObjectAny())
            {
                using (cmListObject)
                {
                    var dest = new FileInfo(Path.Combine(destDir.FullName, cmListObject.ComObject.Name + DefaultFileExtension));
                    var rangeString = ExportTableCore(cmListObject, dest);
                    yield return new ExcelExporterResult(cmListObject.ComObject.Name, rangeString, dest.FullName);
                }
            }
        }
    }
}
