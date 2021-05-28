using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reactive.Disposables;
using Microsoft.Office.Interop.Excel;

namespace ExcelInteropBridging.Core
{
    public abstract class ExcelBridger : IDisposable
    {
        private protected readonly IComManaged<Application> mApplication;
        private protected readonly IComManaged<Workbooks> mWorkbooks;
        private protected readonly IComManaged<Workbook> mWorkbook;
        private readonly CompositeDisposable mReferenceWorkbookList;



        public ExcelBridger(FileInfo source)
            : this(source, Array.Empty<FileInfo>())
        {
        }

        public ExcelBridger(FileInfo source, IEnumerable<FileInfo> reference)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (!source.Exists) throw new FileNotFoundException($"エクセルファイルが見つかりません。", source.FullName);

            reference = reference != null ? reference.Distinct(new FileFullNameComparer()) : Array.Empty<FileInfo>();
            if (reference.Any(file => !file.Exists))
            {
                FileInfo notFound = reference.FirstOrDefault(file => !file.Exists);
                throw new FileNotFoundException($"外部参照エクセルファイルのいずれかが見つかりません。", notFound.FullName);
            }

            try
            {
                // Application
                mApplication = ComManaged.AsManaged(new Application());
                mApplication.ComObject.DisplayAlerts = false;
                mApplication.ComObject.ScreenUpdating = false;
                mApplication.ComObject.AskToUpdateLinks = false;
                // Workbooks
                mWorkbooks = ComManaged.AsManaged(mApplication.ComObject.Workbooks);

                // 外部参照エクセルは、ターゲットエクセル内のリンクデータが [#REF!] に更新されるのを防ぐために使用します。
                // 先に外部参照エクセルを開いておくことで、データ更新を未然に防げます。
                mReferenceWorkbookList = new CompositeDisposable();
                foreach (var file in reference)
                {
                    mReferenceWorkbookList.Add(ComManaged.AsManaged(mWorkbooks.ComObject.Open(file.FullName, 3, true)));
                }

                // Workbook
                mWorkbook = ComManaged.AsManaged(mWorkbooks.ComObject.Open(source.FullName, 3, true));
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
        private protected IComManaged<Worksheet> GetWorksheet(int sheetIndex)
        {
            using (var mSheets = ComManaged.AsManaged(mWorkbook.ComObject.Worksheets))
            {
                var mWorksheet = ComManaged.AsManaged((Worksheet)mSheets.ComObject[sheetIndex]);
                if (mWorksheet.ComObject == null)
                {
                    mWorksheet.Dispose();
                    return null;
                }
                return mWorksheet;
            }
        }

        /// <summary>
        /// 指定したシート名に一致する <see cref="Worksheet"/> を取得します。シートが存在しない場合は <see langword="null"/> が返されます。
        /// </summary>
        private protected IComManaged<Worksheet> GetWorksheet(string sheetName)
        {
            using (var mSheets = ComManaged.AsManaged(mWorkbook.ComObject.Worksheets))
            {
                var mWorksheet = ComManaged.AsManaged((Worksheet)mSheets.ComObject[sheetName]);
                if (mWorksheet.ComObject == null)
                {
                    mWorksheet.Dispose();
                    return null;
                }
                return mWorksheet;
            }
        }

        /// <summary>
        /// 全ての <see cref="Worksheet"/> を取得します。（遅延実行専用）
        /// </summary>
        private protected IEnumerable<IComManaged<Worksheet>> GetWorksheetAny()
        {
            return GetWorksheetAny(null);
        }

        /// <summary>
        /// 指定した条件を満たす全ての <see cref="Worksheet"/> を取得します。（遅延実行専用）
        /// </summary>
        /// <exception cref="ArgumentNullException"><paramref name="predicate"/> が <see langword="null"/> です。</exception>
        private protected IEnumerable<IComManaged<Worksheet>> GetWorksheetAny(Predicate<IComManaged<Worksheet>> predicate)
        {
            using (var mSheets = ComManaged.AsManaged(mWorkbook.ComObject.Worksheets))
            {
                for (var i = 1; i <= mSheets.ComObject.Count; i++)
                {
                    using (var mWorksheet = ComManaged.AsManaged((Worksheet)mSheets.ComObject[i]))
                    {
                        if ((mWorksheet.ComObject != null) && (predicate?.Invoke(mWorksheet) ?? true))
                        {
                            yield return mWorksheet;
                        }
                    }
                }
            }
        }



        /// <summary>
        /// 指定したテーブル名に一致する <see cref="ListObject"/> を取得します。テーブルが存在しない場合は <see langword="null"/> が返されます。
        /// </summary>
        private protected IComManaged<ListObject> GetListObject(string tableName)
        {
            using (var mSheets = ComManaged.AsManaged(mWorkbook.ComObject.Worksheets))
            {
                for (var sheetIndex = 1; sheetIndex <= mSheets.ComObject.Count; sheetIndex++)
                {
                    using (var mWorksheet = ComManaged.AsManaged((Worksheet)mSheets.ComObject[sheetIndex]))
                    using (var mListObjects = ComManaged.AsManaged(mWorksheet.ComObject.ListObjects))
                    {
                        for (var listObjIndex = 1; listObjIndex <= mListObjects.ComObject.Count; listObjIndex++)
                        {
                            var mListObject = ComManaged.AsManaged(mListObjects.ComObject[listObjIndex]);
                            if ((mListObject.ComObject == null) || (mListObject.ComObject.Name != tableName))
                            {
                                mListObject.Dispose();
                                continue;
                            }
                            return mListObject;
                        }
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// 全ての <see cref="ListObject"/> を取得します。（遅延実行専用）
        /// </summary>
        private protected IEnumerable<IComManaged<ListObject>> GetListObjectAny()
        {
            return GetListObjectAny(null);
        }

        /// <summary>
        /// 指定した条件を満たす全ての <see cref="ListObject"/> を取得します。（遅延実行専用）
        /// </summary>
        /// <exception cref="ArgumentNullException"><paramref name="predicate"/> が <see langword="null"/> です。</exception>
        private protected IEnumerable<IComManaged<ListObject>> GetListObjectAny(Predicate<IComManaged<ListObject>> predicate)
        {
            using (var mSheets = ComManaged.AsManaged(mWorkbook.ComObject.Worksheets))
            {
                for (var sheetIndex = 1; sheetIndex <= mSheets.ComObject.Count; sheetIndex++)
                {
                    using (var mWorksheet = ComManaged.AsManaged((Worksheet)mSheets.ComObject[sheetIndex]))
                    using (var mListObjects = ComManaged.AsManaged(mWorksheet.ComObject.ListObjects))
                    {
                        for (var listObjIndex = 1; listObjIndex <= mListObjects.ComObject.Count; listObjIndex++)
                        {
                            using (var mListObject = ComManaged.AsManaged(mListObjects.ComObject[listObjIndex]))
                            {
                                if ((mListObject.ComObject != null) && (predicate?.Invoke(mListObject) ?? true))
                                {
                                    yield return mListObject;
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
            return GetWorksheetAny().Select(x => x.ComObject.Name).ToArray();
        }

        /// <summary>
        /// エクセルファイルに指定したシートが存在するかを判断します。
        /// </summary>
        /// <param name="sheetName">シート名。</param>
        /// <returns>シートが存在する場合は <see langword="true"/>。存在しない場合は <see langword="false"/>。</returns>
        public bool ContainsSheet(string sheetName)
        {
            using (var mWorksheet = GetWorksheet(sheetName))
            {
                return mWorksheet != null;
            }
        }



        /// <summary>
        /// エクセルファイルに含まれる全てのテーブルの名前を取得します。
        /// </summary>
        /// <returns>全テーブル名の列挙。</returns>
        public string[] GetTableNames()
        {
            return GetListObjectAny().Select(x => x.ComObject.Name).ToArray();
        }

        /// <summary>
        /// エクセルファイルに指定したテーブルが存在するかを判断します。
        /// </summary>
        /// <param name="tableName">テーブル名。</param>
        /// <returns>テーブルが存在する場合は <see langword="true"/>。存在しない場合は <see langword="false"/>。</returns>
        public bool ContainsTable(string tableName)
        {
            using (var mListObject = GetListObject(tableName))
            {
                return mListObject != null;
            }
        }



        #region IDisposable Support
        private bool disposed = false;

        protected virtual void Dispose(bool disposing)
        {
            if (disposed) return;

            if (disposing)
            {
                // Dispose の順番はコンストラクタ―内の作成順と逆になるようにします。
                mWorkbook?.Dispose();
                mReferenceWorkbookList?.Dispose();
                mWorkbooks?.Dispose();
                mApplication?.Dispose();
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
