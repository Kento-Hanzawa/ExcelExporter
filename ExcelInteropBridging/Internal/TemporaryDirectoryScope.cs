using System;
using System.IO;

namespace ExcelInteropBridging
{
	/// <summary>
	/// ディスク上に一時ディレクトリを作成し、スコープを抜けた際に自動でフォルダを削除する機能を提供します。
	/// </summary>
	internal sealed class TemporaryDirectoryScope : IDisposable
	{
		/// <summary>
		/// 一時ディレクトリの完全パスを取得します。
		/// </summary>
		public string TemporaryDirectoryName { get; private set; }

		public TemporaryDirectoryScope()
		{
			string tempPath = Path.GetTempPath();
			do
			{
				TemporaryDirectoryName = Path.Combine(tempPath, Path.GetRandomFileName());
			}
			while (Directory.Exists(TemporaryDirectoryName));
			Directory.CreateDirectory(TemporaryDirectoryName);
		}

		public void Dispose()
		{
			if (Directory.Exists(TemporaryDirectoryName))
			{
				Directory.Delete(TemporaryDirectoryName, true);
			}
		}
	}
}
