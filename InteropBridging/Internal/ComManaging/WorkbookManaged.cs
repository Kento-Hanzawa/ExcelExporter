using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace InteropBridging.Internal
{
	/// <summary>
	/// <see cref="Workbook"/> オブジェクト専用の COM リソース解放機能を保証する <see cref="IComManaged{T}"/> を提供します。
	/// </summary>
	internal sealed class WorkbookManaged : IComManaged<Workbook>
	{
		private Workbook internalComObject;

		/// <summary>
		/// ラッピングされた <see cref="Workbook"/> オブジェクト。
		/// </summary>
		public ref readonly Workbook ComObject { get { return ref internalComObject; } }

		/// <summary>
		/// ラッピングする <see cref="Workbook"/> オブジェクトを指定して <see cref="WorkbookManaged"/> クラスの新しいインスタンスを作成します。
		/// </summary>
		/// <param name="workbook"><see cref="Workbook"/> オブジェクト。</param>
		public WorkbookManaged(in Workbook workbook)
		{
			internalComObject = workbook;
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

				if (internalComObject != null)
				{
					internalComObject.Close(false);
					Marshal.FinalReleaseComObject(internalComObject);
					internalComObject = null;
				}

				disposedValue = true;
			}
		}

		~WorkbookManaged()
		{
			Dispose(false);
		}

		/// <summary>
		/// ラッピングする <see cref="Workbook"/> オブジェクトを解放します。
		/// </summary>
		public void Dispose()
		{
			Dispose(true);
			GC.SuppressFinalize(this);
		}
		#endregion
	}
}
