using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reactive.Disposables;
using InteropBridging.Internal;
using Microsoft.Office.Interop.Excel;

namespace InteropBridging
{
	public abstract class ExcelBridge : IDisposable
	{
		private protected readonly IComManaged<Application> managedApplication;
		private protected readonly IComManaged<Workbooks> managedWorkbooks;
		private protected readonly IComManaged<Workbook> managedWorkbook;
		private readonly CompositeDisposable managedReferenceWorkbookList;

		public ExcelBridge(FileInfo excelFile)
			: this(excelFile, Array.Empty<FileInfo>())
		{
		}

		public ExcelBridge(FileInfo excelFile, IEnumerable<FileInfo> referenceExcelFiles)
		{
			if (excelFile == null) throw new ArgumentNullException(nameof(excelFile));
			if (!excelFile.Exists) throw new FileNotFoundException($"指定されたエクセルファイルが見つかりません。", excelFile.FullName);

			if (referenceExcelFiles == null) referenceExcelFiles = Array.Empty<FileInfo>();
			if (referenceExcelFiles.Any(file => !file.Exists))
			{
				FileInfo notFound = referenceExcelFiles.FirstOrDefault(file => !file.Exists);
				throw new FileNotFoundException($"指定された外部参照エクセルファイルのいずれかが見つかりません。", notFound.FullName);
			}

			try
			{
				managedApplication = ComManaged.Manage(new Application());
				managedApplication.ComObject.DisplayAlerts = false;
				managedApplication.ComObject.ScreenUpdating = false;
				managedApplication.ComObject.AskToUpdateLinks = false;
				managedWorkbooks = ComManaged.Manage(managedApplication.ComObject.Workbooks);
				// 外部参照エクセルは、ターゲットエクセル内のリンクデータが [#REF!] に更新されるのを防ぐために使用します。
				// 先に外部参照エクセルを開いておくことで、データ更新を未然に防げます。
				managedReferenceWorkbookList = new CompositeDisposable();
				foreach (var file in referenceExcelFiles.Distinct(FileFullNameComparer.Instance))
				{
					managedReferenceWorkbookList.Add(ComManaged.Manage(managedWorkbooks.ComObject.Open(file.FullName, 3, true)));
				}
				managedWorkbook = ComManaged.Manage(managedWorkbooks.ComObject.Open(excelFile.FullName, 3, true));
			}
			catch
			{
				Dispose();
				throw;
			}
		}

		#region IDisposable Support
		private bool disposedValue = false;

		void Dispose(bool disposing)
		{
			if (!disposedValue)
			{
				if (disposing)
				{
				}

				// Dispose の順番はコンストラクタ―内の作成順と逆になるようにします。
				managedWorkbook?.Dispose();
				managedReferenceWorkbookList?.Dispose();
				managedWorkbooks?.Dispose();
				managedApplication?.Dispose();

				disposedValue = true;
			}
		}

		~ExcelBridge()
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
			public static IEqualityComparer<FileInfo> Instance { get; } = new FileFullNameComparer();

			private FileFullNameComparer()
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
