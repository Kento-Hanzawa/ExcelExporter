using System;
using System.Runtime.InteropServices;

namespace ExcelInteropBridging.Core
{
	/// <summary>
	/// COM リソース解放の最低限の機能のみを保証する <see cref="IComManaged{T}"/> を提供します。
	/// </summary>
	/// <typeparam name="T">ラッピングする COM オブジェクトの型。</typeparam>
	internal sealed class GeneralComManaged<T> : IComManaged<T>
	{
		private T internalComObject;

		/// <summary>
		/// ラッピングされた COM オブジェクト。
		/// </summary>
		public ref readonly T ComObject { get { return ref internalComObject; } }

		/// <summary>
		/// ラッピングする COM オブジェクトを指定して <see cref="GeneralComManaged{T}"/> クラスの新しいインスタンスを作成します。
		/// </summary>
		/// <param name="comObject">COM オブジェクト。</param>
		public GeneralComManaged(in T comObject)
		{
			internalComObject = comObject;
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
				Marshal.FinalReleaseComObject(internalComObject);
				internalComObject = default;
			}

			disposed = true;
		}

		~GeneralComManaged()
		{
			Dispose(false);
		}

		/// <summary>
		/// ラッピングする COM オブジェクトを解放します。
		/// </summary>
		public void Dispose()
		{
			Dispose(true);
			GC.SuppressFinalize(this);
		}
		#endregion
	}
}
