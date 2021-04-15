namespace InteropBridging.Internal
{
	/// <summary>
	/// <see cref="IComManaged{T}"/> を作成するためのユーティリティを提供します。
	/// </summary>
	internal static class ComManaged
	{
		public static IComManaged<T> Manage<T>(in T comObject)
		{
			return new GeneralComManaged<T>(comObject);
		}

		public static IComManaged<Microsoft.Office.Interop.Excel.Application> Manage(in Microsoft.Office.Interop.Excel.Application comObject)
		{
			return new ApplicationComManaged(comObject);
		}

		public static IComManaged<Microsoft.Office.Interop.Excel.Workbooks> Manage(in Microsoft.Office.Interop.Excel.Workbooks comObject)
		{
			return new WorkbooksManaged(comObject);
		}

		public static IComManaged<Microsoft.Office.Interop.Excel.Workbook> Manage(in Microsoft.Office.Interop.Excel.Workbook comObject)
		{
			return new WorkbookManaged(comObject);
		}
	}
}
