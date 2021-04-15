using System;
using System.IO;

namespace InteropBridging.Internal
{
	/// <summary>
	/// ディスク上に一時ディレクトリを作成し、スコープを抜けた際に自動でフォルダを削除する機能を提供します。
	/// </summary>
	internal sealed class TemporaryDirectoryScope : IDisposable
	{
		/// <summary>
		/// 一時ディレクトリの完全パスを取得します。
		/// </summary>
		public string TemporaryDirectoryPath { get; private set; }

		public TemporaryDirectoryScope()
		{
			string tempPath = Path.GetTempPath();
			do
			{
				TemporaryDirectoryPath = Path.Combine(tempPath, Path.GetRandomFileName());
			}
			while (Directory.Exists(TemporaryDirectoryPath));
			Directory.CreateDirectory(TemporaryDirectoryPath);
		}

		#region IDisposable Support
		private bool disposedValue = false;

		void Dispose(bool disposing)
		{
			if (!disposedValue)
			{
				if (disposing)
				{
					if (Directory.Exists(TemporaryDirectoryPath))
					{
						Directory.Delete(TemporaryDirectoryPath, true);
					}
				}
				disposedValue = true;
			}
		}

		public void Dispose()
		{
			Dispose(true);
		}
		#endregion
	}
}
