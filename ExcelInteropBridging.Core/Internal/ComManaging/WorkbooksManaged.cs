using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelInteropBridging.Core
{
	/// <summary>
	/// <see cref="Workbooks"/> オブジェクト専用の COM リソース解放機能を保証する <see cref="IComManaged{T}"/> を提供します。
	/// </summary>
	internal sealed class WorkbooksManaged : IComManaged<Workbooks>
	{
		private Workbooks internalComObject;

		/// <summary>
		/// ラッピングされた <see cref="Workbooks"/> オブジェクト。
		/// </summary>
		public ref readonly Workbooks ComObject { get { return ref internalComObject; } }

		/// <summary>
		/// ラッピングする <see cref="Workbooks"/> オブジェクトを指定して <see cref="WorkbooksManaged"/> クラスの新しいインスタンスを作成します。
		/// </summary>
		/// <param name="workbooks"><see cref="Workbooks"/> オブジェクト。</param>
		public WorkbooksManaged(in Workbooks workbooks)
		{
			internalComObject = workbooks;
		}

		#region IDisposable Support
		private bool disposed = false;

		void Dispose(bool disposing)
		{
			if (disposed) return;

			if (disposing)
			{
			}

			if (internalComObject != null)
			{
				internalComObject.Close();
				Marshal.FinalReleaseComObject(internalComObject);
				internalComObject = null;
			}

			disposed = true;
		}

		~WorkbooksManaged()
		{
			Dispose(false);
		}

		/// <summary>
		/// ラッピングする <see cref="Workbooks"/> オブジェクトを解放します。
		/// </summary>
		public void Dispose()
		{
			Dispose(true);
			GC.SuppressFinalize(this);
		}
		#endregion
	}
}
