﻿using System;

namespace InteropBridging.Internal
{
	/// <summary>
	/// COM オブジェクトをラッピングする機能を提供します。このインターフェースを実装するクラスは、<see cref="IDisposable.Dispose"/> からの COM リソース解放を保証する必要があります。
	/// </summary>
	/// <typeparam name="T">ラッピングする COM オブジェクトの型。</typeparam>
	internal interface IComManaged<T> : IDisposable
	{
		ref readonly T ComObject { get; }
	}
}