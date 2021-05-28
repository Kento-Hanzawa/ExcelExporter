using System;
using System.Globalization;
using System.IO;
using System.Text;
using CsvHelper;
using CsvHelper.Configuration;
using ExcelInteropBridging;

namespace ExcelInteropBridging.Csv
{
	class Program
	{
		static void Main(string[] args)
		{
			//var fi = new FileInfo(@"D:\hanzawa\TemporaryDocuments\2021-05-27\【CustomCast】翻訳マスター.xlsx");
			var fi = new FileInfo(@"D:\hanzawa\TemporaryDocuments\2021-05-27\Book1.xlsx");

			using (var expo = new Converter(fi))
			{
                foreach (var result in expo.ExportAllSheet(new DirectoryInfo(Path.Combine(fi.DirectoryName, "out"))))
                {
                    Console.WriteLine($"{result.RangeName} : {result.RangeString}");
                }
                //expo.ExportODC(new DirectoryInfo(Path.Combine(fi.DirectoryName, "out")));
			}
			//using (var e = new Converter(fi))
			//{

			//	foreach (var result in e.ConvertSheet("List"))
			//	{
			//		var outputConfiguration = new CsvConfiguration(CultureInfo.CurrentCulture)
			//		{
			//			TrimOptions = TrimOptions.Trim,
			//			Encoding = new UTF8Encoding(false)
			//		};
			//		using var fileStream = File.OpenWrite(Path.Combine(fi.DirectoryName, result.RangeName + ".csv"));
			//		using var streamWriter = new StreamWriter(fileStream, new UTF8Encoding(false));
			//		using var csvWriter = new CsvWriter(streamWriter, outputConfiguration);
			//		foreach (var record in result.ParsedData)
			//		{
			//			foreach (var field in record)
			//			{
			//				csvWriter.WriteField(field.Value);
			//			}
			//			csvWriter.NextRecord();
			//		}
			//	}
			//}
		}
	}
}
