using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reactive.Disposables;
using Microsoft.Office.Interop.Excel;

namespace ExcelInteropBridging.Internal
{
    internal sealed class ExcelBridger : IDisposable
    {
        private readonly IComManaged<Application> mgApplication;
        private readonly IComManaged<Workbooks> mgWorkbooks;
        private readonly IComManaged<Workbook> mgWorkbook;
        private readonly CompositeDisposable mgWorkbookReferences;

        public IComManaged<Application> MgApplication => mgApplication;
        public IComManaged<Workbooks> MgWorkbooks => mgWorkbooks;
        public IComManaged<Workbook> MgWorkbook => mgWorkbook;



        public ExcelBridger(FileInfo source)
            : this(source, Array.Empty<FileInfo>())
        {
        }

        public ExcelBridger(FileInfo source, IEnumerable<FileInfo> references)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (!source.Exists) throw new FileNotFoundException($"エクセルファイルが見つかりません。", source.FullName);

            references = references != null ? references.Distinct(new FileFullNameComparer()) : Array.Empty<FileInfo>();
            if (references.Any(file => !file.Exists))
            {
                FileInfo notFound = references.FirstOrDefault(file => !file.Exists);
                throw new FileNotFoundException($"外部参照エクセルファイルのいずれかが見つかりません。", notFound.FullName);
            }

            try
            {
                // Application
                mgApplication = ComManaged.AsManaged(new Application());
                mgApplication.ComObject.DisplayAlerts = false;
                mgApplication.ComObject.ScreenUpdating = false;
                mgApplication.ComObject.AskToUpdateLinks = false;
                // Workbooks
                mgWorkbooks = ComManaged.AsManaged(mgApplication.ComObject.Workbooks);

                // 外部参照エクセルは、ターゲットエクセル内のリンクデータが [#REF!] に更新されるのを防ぐために使用します。
                // 先に外部参照エクセルを開いておくことで、データ更新を未然に防げます。
                mgWorkbookReferences = new CompositeDisposable();
                foreach (var file in references)
                {
                    mgWorkbookReferences.Add(ComManaged.AsManaged(mgWorkbooks.ComObject.Open(file.FullName, 3, true)));
                }

                // Workbook
                mgWorkbook = ComManaged.AsManaged(mgWorkbooks.ComObject.Open(source.FullName, 3, true));
            }
            catch
            {
                Dispose();
                throw;
            }
        }



        /// <summary>
        /// 指定したシート番号の <see cref="Worksheet"/> を取得します。シートが存在しない場合は <see langword="null"/> が返されます。
        /// </summary>
        public IComManaged<Worksheet> GetWorksheet(int sheetIndex)
        {
            using (var mgSheets = ComManaged.AsManaged(mgWorkbook.ComObject.Worksheets))
            {
                var mgWorksheet = ComManaged.AsManaged((Worksheet)mgSheets.ComObject[sheetIndex]);
                if (mgWorksheet.ComObject == null)
                {
                    mgWorksheet.Dispose();
                    return null;
                }
                return mgWorksheet;
            }
        }

        /// <summary>
        /// 指定したシート名に一致する <see cref="Worksheet"/> を取得します。シートが存在しない場合は <see langword="null"/> が返されます。
        /// </summary>
        public IComManaged<Worksheet> GetWorksheet(string sheetName)
        {
            using (var mgSheets = ComManaged.AsManaged(mgWorkbook.ComObject.Worksheets))
            {
                var mgWorksheet = ComManaged.AsManaged((Worksheet)mgSheets.ComObject[sheetName]);
                if (mgWorksheet.ComObject == null)
                {
                    mgWorksheet.Dispose();
                    return null;
                }
                return mgWorksheet;
            }
        }

        /// <summary>
        /// 全ての <see cref="Worksheet"/> を取得します。（遅延実行専用）
        /// </summary>
        public IEnumerable<IComManaged<Worksheet>> GetWorksheetAnyEnumerable()
        {
            return GetWorksheetAnyEnumerable(null);
        }

        /// <summary>
        /// 指定した条件を満たす全ての <see cref="Worksheet"/> を取得します。（遅延実行専用）
        /// </summary>
        /// <exception cref="ArgumentNullException"><paramref name="predicate"/> が <see langword="null"/> です。</exception>
        public IEnumerable<IComManaged<Worksheet>> GetWorksheetAnyEnumerable(Predicate<IComManaged<Worksheet>> predicate)
        {
            using (var mgSheets = ComManaged.AsManaged(mgWorkbook.ComObject.Worksheets))
            {
                for (var i = 1; i <= mgSheets.ComObject.Count; i++)
                {
                    using (var mgWorksheet = ComManaged.AsManaged((Worksheet)mgSheets.ComObject[i]))
                    {
                        if ((mgWorksheet.ComObject != null) && (predicate?.Invoke(mgWorksheet) ?? true))
                        {
                            yield return mgWorksheet;
                        }
                    }
                }
            }
        }



        /// <summary>
        /// 指定したテーブル名に一致する <see cref="ListObject"/> を取得します。テーブルが存在しない場合は <see langword="null"/> が返されます。
        /// </summary>
        public IComManaged<ListObject> GetListObject(string tableName)
        {
            using (var mgSheets = ComManaged.AsManaged(mgWorkbook.ComObject.Worksheets))
            {
                for (var sheetIndex = 1; sheetIndex <= mgSheets.ComObject.Count; sheetIndex++)
                {
                    using (var mgWorksheet = ComManaged.AsManaged((Worksheet)mgSheets.ComObject[sheetIndex]))
                    using (var mgListObjects = ComManaged.AsManaged(mgWorksheet.ComObject.ListObjects))
                    {
                        for (var listObjIndex = 1; listObjIndex <= mgListObjects.ComObject.Count; listObjIndex++)
                        {
                            var mgListObject = ComManaged.AsManaged(mgListObjects.ComObject[listObjIndex]);
                            if ((mgListObject.ComObject == null) || (mgListObject.ComObject.Name != tableName))
                            {
                                mgListObject.Dispose();
                                continue;
                            }
                            return mgListObject;
                        }
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// 全ての <see cref="ListObject"/> を取得します。（遅延実行専用）
        /// </summary>
        public IEnumerable<IComManaged<ListObject>> GetListObjectAnyEnumerable()
        {
            return GetListObjectAnyEnumerable(null);
        }

        /// <summary>
        /// 指定した条件を満たす全ての <see cref="ListObject"/> を取得します。（遅延実行専用）
        /// </summary>
        /// <exception cref="ArgumentNullException"><paramref name="predicate"/> が <see langword="null"/> です。</exception>
        public IEnumerable<IComManaged<ListObject>> GetListObjectAnyEnumerable(Predicate<IComManaged<ListObject>> predicate)
        {
            using (var mgSheets = ComManaged.AsManaged(mgWorkbook.ComObject.Worksheets))
            {
                for (var sheetIndex = 1; sheetIndex <= mgSheets.ComObject.Count; sheetIndex++)
                {
                    using (var mgWorksheet = ComManaged.AsManaged((Worksheet)mgSheets.ComObject[sheetIndex]))
                    using (var mgListObjects = ComManaged.AsManaged(mgWorksheet.ComObject.ListObjects))
                    {
                        for (var listObjIndex = 1; listObjIndex <= mgListObjects.ComObject.Count; listObjIndex++)
                        {
                            using (var mgListObject = ComManaged.AsManaged(mgListObjects.ComObject[listObjIndex]))
                            {
                                if ((mgListObject.ComObject != null) && (predicate?.Invoke(mgListObject) ?? true))
                                {
                                    yield return mgListObject;
                                }
                            }
                        }
                    }
                }
            }
        }



        /// <summary>
        /// エクセルファイルに含まれる全てのシート名を取得します。
        /// </summary>
        /// <returns>全シート名の列挙。</returns>
        public string[] GetSheetNames()
        {
            return GetWorksheetAnyEnumerable().Select(x => x.ComObject.Name).ToArray();
        }

        /// <summary>
        /// エクセルファイルに指定したシートが存在するかを判断します。
        /// </summary>
        /// <param name="sheetName">シート名。</param>
        /// <returns>シートが存在する場合は <see langword="true"/>。存在しない場合は <see langword="false"/>。</returns>
        public bool ContainsSheet(string sheetName)
        {
            using (var mgWorksheet = GetWorksheet(sheetName))
            {
                return mgWorksheet != null;
            }
        }



        /// <summary>
        /// エクセルファイルに含まれる全てのテーブルの名前を取得します。
        /// </summary>
        /// <returns>全テーブル名の列挙。</returns>
        public string[] GetTableNames()
        {
            return GetListObjectAnyEnumerable().Select(x => x.ComObject.Name).ToArray();
        }

        /// <summary>
        /// エクセルファイルに指定したテーブルが存在するかを判断します。
        /// </summary>
        /// <param name="tableName">テーブル名。</param>
        /// <returns>テーブルが存在する場合は <see langword="true"/>。存在しない場合は <see langword="false"/>。</returns>
        public bool ContainsTable(string tableName)
        {
            using (var mgListObject = GetListObject(tableName))
            {
                return mgListObject != null;
            }
        }



        #region IDisposable Support
        private bool disposed = false;

        private void Dispose(bool disposing)
        {
            if (disposed) return;

            if (disposing)
            {
                // Dispose の順番はコンストラクタ―内の作成順と逆になるようにします。
                mgWorkbook?.Dispose();
                mgWorkbookReferences?.Dispose();
                mgWorkbooks?.Dispose();
                mgApplication?.Dispose();
            }

            disposed = true;
        }

        ~ExcelBridger()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        #endregion

        private sealed class FileFullNameComparer : IEqualityComparer<FileInfo>
        {
            public FileFullNameComparer()
            {
            }

            public bool Equals(FileInfo x, FileInfo y)
            {
                return x.FullName == y.FullName;
            }

            public int GetHashCode(FileInfo obj)
            {
                return obj.FullName.GetHashCode();
            }
        }
    }
}
