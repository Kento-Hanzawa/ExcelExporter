using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace InteropBridging.Internal
{
	/// <summary>
	/// <see cref="Application"/> オブジェクト専用の COM リソース解放機能を保証する <see cref="IComManaged{T}"/> を提供します。
	/// </summary>
	internal sealed class ApplicationComManaged : IComManaged<Application>
	{
		private Application internalComObject;

		/// <summary>
		/// ラッピングされた <see cref="Application"/> オブジェクト。
		/// </summary>
		public ref readonly Application ComObject { get { return ref internalComObject; } }

		/// <summary>
		/// ラッピングする <see cref="Application"/> オブジェクトを指定して <see cref="ApplicationComManaged"/> クラスの新しいインスタンスを作成します。
		/// </summary>
		/// <param name="application"><see cref="Application"/> オブジェクト。</param>
		public ApplicationComManaged(in Application application)
		{
			internalComObject = application;
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

				// Microsoft.Office.Interop.Excel.Application の COM オブジェクト解放手順は、以下のサイトを参考にしています。
				// https://blogs.msdn.microsoft.com/office_client_development_support_blog/2012/02/09/office-5/

				// アプリケーションの終了前にガベージコレクトを強制します。
				GC.Collect();
				GC.WaitForPendingFinalizers();
				GC.Collect();

				if (internalComObject != null)
				{
					internalComObject.DisplayAlerts = true;
					internalComObject.ScreenUpdating = true;
					internalComObject.AskToUpdateLinks = true;
					internalComObject.Quit();
					Marshal.FinalReleaseComObject(internalComObject);
					internalComObject = null;

					// アプリケーションの終了後にガベージコレクトを強制します。
					GC.Collect();
					GC.WaitForPendingFinalizers();
					GC.Collect();
				}

				disposedValue = true;
			}
		}

		~ApplicationComManaged()
		{
			Dispose(false);
		}

		/// <summary>
		/// ラッピングする <see cref="Application"/> オブジェクトを解放します。
		/// </summary>
		public void Dispose()
		{
			Dispose(true);
			GC.SuppressFinalize(this);
		}
		#endregion
	}
}
