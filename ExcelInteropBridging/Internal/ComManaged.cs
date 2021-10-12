using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelInteropBridging.Internal
{
	/// <summary>
	/// <see cref="IComManaged{T}"/> を作成するためのユーティリティを提供します。
	/// </summary>
	internal static class ComManaged
	{
		/// <summary>
		/// 対象の COM オブジェクトをラッピングした <see cref="IComManaged{T}"/> のインスタンスを取得します。
		/// </summary>
		/// <param name="comObject">ラッピングする COM オブジェクト。</param>
		public static IComManaged<T> AsManaged<T>(T comObject)
			where T : class
		{
			return new DefaultComManaged<T>(comObject);
		}

		/// <summary>
		/// 対象の COM オブジェクトをラッピングした <see cref="IComManaged{T}"/> のインスタンスを取得します。
		/// </summary>
		/// <param name="comObject">ラッピングする COM オブジェクト。</param>
		public static IComManaged<Application> AsManaged(Application comObject)
		{
			return new ExcelApplicationManaged(comObject);
		}

		/// <summary>
		/// 対象の COM オブジェクトをラッピングした <see cref="IComManaged{T}"/> のインスタンスを取得します。
		/// </summary>
		/// <param name="comObject">ラッピングする COM オブジェクト。</param>
		public static IComManaged<Workbooks> AsManaged(Workbooks comObject)
		{
			return new ExcelWorkbooksManaged(comObject);
		}

		/// <summary>
		/// 対象の COM オブジェクトをラッピングした <see cref="IComManaged{T}"/> のインスタンスを取得します。
		/// </summary>
		/// <param name="comObject">ラッピングする COM オブジェクト。</param>
		public static IComManaged<Workbook> AsManaged(Workbook comObject)
		{
			return new ExcelWorkbookManaged(comObject);
		}



		private sealed class DefaultComManaged<T> : IComManaged<T>
			where T : class
		{
			private T comObject;
			public T ComObject { get { return comObject; } }

			public DefaultComManaged(T comObject)
			{
				this.comObject = comObject;
			}

			#region IDisposable Support
			private bool disposed = false;

			void Dispose(bool disposing)
			{
				if (disposed) return;
				if (disposing) { }
				if (comObject != null)
				{
					Marshal.FinalReleaseComObject(comObject);
					comObject = default;
				}
				disposed = true;
			}

			~DefaultComManaged()
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

		private sealed class ExcelApplicationManaged : IComManaged<Application>
		{
			private Application application;
			public Application ComObject { get { return application; } }

			public ExcelApplicationManaged(Application application)
			{
				this.application = application;
			}

			#region IDisposable Support
			private bool disposed = false;

			void Dispose(bool disposing)
			{
				if (disposed) return;
				if (disposing) { }

				// Microsoft.Office.Interop.Excel.Application の COM オブジェクト解放手順は、以下のサイトを参考にしています。
				// https://blogs.msdn.microsoft.com/office_client_development_support_blog/2012/02/09/office-5/

				// アプリケーションの終了前にガベージコレクトを強制します。
				GC.Collect();
				GC.WaitForPendingFinalizers();
				GC.Collect();

				if (application != null)
				{
					application.DisplayAlerts = true;
					application.ScreenUpdating = true;
					application.AskToUpdateLinks = true;
					application.Quit();
					Marshal.FinalReleaseComObject(application);
					application = null;

					// アプリケーションの終了後にガベージコレクトを強制します。
					GC.Collect();
					GC.WaitForPendingFinalizers();
					GC.Collect();
				}

				disposed = true;
			}

			~ExcelApplicationManaged()
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

		private sealed class ExcelWorkbooksManaged : IComManaged<Workbooks>
		{
			private Workbooks workbooks;
			public Workbooks ComObject { get { return workbooks; } }

			public ExcelWorkbooksManaged(Workbooks workbooks)
			{
				this.workbooks = workbooks;
			}

			#region IDisposable Support
			private bool disposed = false;

			void Dispose(bool disposing)
			{
				if (disposed) return;
				if (disposing) { }
				if (workbooks != null)
				{
					workbooks.Close();
					Marshal.FinalReleaseComObject(workbooks);
					workbooks = null;
				}
				disposed = true;
			}

			~ExcelWorkbooksManaged()
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

		private sealed class ExcelWorkbookManaged : IComManaged<Workbook>
		{
			private Workbook workbook;
			public Workbook ComObject { get { return workbook; } }

			public ExcelWorkbookManaged(Workbook workbook)
			{
				this.workbook = workbook;
			}

			#region IDisposable Support
			private bool disposed = false;

			void Dispose(bool disposing)
			{
				if (disposed) return;
				if (disposing) { }
				if (workbook != null)
				{
					workbook.Close(false);
					Marshal.FinalReleaseComObject(workbook);
					workbook = null;
				}
				disposed = true;
			}

			~ExcelWorkbookManaged()
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
}
